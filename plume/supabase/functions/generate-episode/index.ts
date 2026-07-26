// PLUME — Edge Function `generate-episode`
// Le cœur du produit : génère l'épisode canonique du soir pour un enfant.
// Lancée chaque nuit par pg_cron (pré-génération : l'enfant n'attend jamais),
// ou à la demande si l'épisode du jour manque.
//
// Pipeline (docs/03-ARCHITECTURE.md §2) :
//   1. assembler le contexte canonique (monde, graines mûres, check-in du jour)
//   2. demander à Claude un épisode STRUCTURÉ (JSON), contraint par le canon
//   3. valider (sécurité, cohérence, longueur cible du Pacte de sommeil)
//   4. committer en transaction : scènes + événements + graines plantées/récoltées
//
// Les clés IA ne quittent jamais ce serveur.

import { createClient } from "npm:@supabase/supabase-js@2";
import Anthropic from "npm:@anthropic-ai/sdk";

const supabase = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);
const anthropic = new Anthropic({ apiKey: Deno.env.get("ANTHROPIC_API_KEY")! });

const MODEL = "claude-sonnet-5";

interface EpisodePlan {
  title: string;
  emotional_thread: string | null;
  scenes: {
    text: string;
    illustration_brief: string;
  }[];
  world_events: string[]; // ce qui devient canonique ce soir
  harvested_seed_ids: string[]; // promesses tenues ce soir
  next_recap: string; // « la dernière fois… » du prochain épisode
}

Deno.serve(async (req) => {
  try {
    const { child_id } = await req.json();
    if (!child_id) return json({ error: "child_id requis" }, 400);

    const context = await assembleContext(child_id);
    if (!context) return json({ error: "enfant inconnu" }, 404);

    const plan = await writeEpisode(context);
    validate(plan, context);
    const episodeId = await commit(child_id, context, plan);

    return json({ episode_id: episodeId, title: plan.title });
  } catch (error) {
    console.error("generate-episode failed:", error);
    return json({ error: "generation_failed" }, 500);
  }
});

// ─── 1. Contexte canonique ────────────────────────────────────────────────────

async function assembleContext(childId: string) {
  const { data: child } = await supabase
    .from("children").select("*").eq("id", childId).maybeSingle();
  if (!child) return null;

  const [companion, world, cast, checkin, lastEpisode] = await Promise.all([
    one("companions", "child_id", childId),
    one("worlds", "child_id", childId),
    many("family_cast", "child_id", childId),
    supabase.from("emotion_checkins").select("*").eq("child_id", childId)
      .eq("date", today()).maybeSingle().then((r) => r.data),
    supabase.from("episodes").select("number, title, next_recap")
      .eq("child_id", childId)
      .order("number", { ascending: false }).limit(1).maybeSingle()
      .then((r) => r.data),
  ]);

  // Graines mûres : moments de vie d'abord, puis priorité, puis ancienneté —
  // une promesse narrative ne doit jamais être oubliée.
  const { data: seeds } = await supabase
    .from("narrative_seeds").select("*").eq("child_id", childId)
    .is("harvested_at", null).lte("germinate_after", new Date().toISOString())
    .order("planted_at").limit(10);
  const ripeSeeds = (seeds ?? [])
    .sort((a, b) =>
      (b.kind === "lifeMoment" ? 1 : 0) - (a.kind === "lifeMoment" ? 1 : 0) ||
      prio(b.priority) - prio(a.priority))
    .slice(0, 2);

  // Les 12 derniers événements du monde + entités actives : le canon récent.
  const { data: events } = await supabase
    .from("world_events").select("summary, episode_number")
    .eq("world_id", world?.id ?? "").order("created_at", { ascending: false })
    .limit(12);
  const { data: entities } = await supabase
    .from("world_entities").select("kind, name, traits, state")
    .eq("world_id", world?.id ?? "").limit(40);

  const age = Math.floor(
    (Date.now() - new Date(child.birth_date).getTime()) / 31_557_600_000,
  );

  return {
    child, companion, world, cast, checkin, ripeSeeds,
    events: events ?? [], entities: entities ?? [],
    age, nextNumber: (lastEpisode?.number ?? 0) + 1,
    previousRecap: lastEpisode?.next_recap ?? null,
  };
}

// ─── 2. Écriture par Claude, contrainte par le canon ─────────────────────────

async function writeEpisode(ctx: NonNullable<Awaited<ReturnType<typeof assembleContext>>>): Promise<EpisodePlan> {
  // Registre littéraire choisi par la famille (worlds.style_bible.register) :
  // « réalisme doux » (défaut) = monde réel, poésie du quotidien, animaux vrais ;
  // « merveilleux » = conte classique. Jamais imposé, toujours réglable.
  const register =
    (ctx.world?.style_bible as { register?: string } | null)?.register ??
    "réalisme doux";

  const system = `Tu es le Conteur de Plume : tu écris l'épisode du soir d'un monde
persistant et unique appartenant à un enfant. Règles absolues :
0. REGISTRE — « ${register} ». En réalisme doux : monde réel uniquement (jardin,
   maison, saisons, animaux vrais), aucune créature fantastique, aucune magie
   explicite — la magie vient de la mémoire du monde et de la beauté du
   quotidien. Écriture sobre et chaleureuse, jamais mièvre ni « bizarre ».
1. CANON — tu n'inventes jamais rien qui contredise les entités et événements
   fournis. Les personnages se souviennent de tout. Tu peux introduire au plus
   UNE nouvelle entité par épisode.
2. FILIGRANE — si un fil émotionnel du jour est fourni (peur, victoire, tristesse),
   tisse-le par miroir dans l'histoire, JAMAIS frontalement. Peur du dentiste →
   le chien de l'histoire appréhende sa visite chez le vétérinaire et la
   surmonte. Ne nomme jamais la situation réelle de l'enfant.
3. GRAINES — si des graines mûres sont fournies, tiens ces promesses ce soir :
   c'est le moment le plus magique (« le dragon sauvé revient »). Marque-les récoltées.
4. ÂGE — vocabulaire, rythme et thèmes calibrés à l'âge indiqué. Fin TOUJOURS
   apaisante : l'épisode conduit au sommeil. Aucune violence, aucune peur non
   résolue, aucun contenu inadapté.
5. AUCUNE INTERACTION — jamais de choix, jamais de question posée à l'enfant.
   Le récit se déroule seul, comme un feuilleton qu'on écoute. C'est le canon
   (et les journées réelles de l'enfant) qui fait avancer l'histoire.
6. FEUILLETON — l'épisode COMMENCE par un court rappel (2 phrases maximum,
   « La dernière fois… ») qui reprend le résumé fourni, puis CONTINUE le récit
   là où il s'était arrêté. Il se termine par une note douce qui donne envie
   de demain, jamais un suspense angoissant. Fournis aussi next_recap : le
   rappel que lira l'épisode suivant.
7. DURÉE — le parent a demandé ${plannedMinutes(ctx)} minute(s) : longueur
   totale ≈ ${wordsForMinutes(plannedMinutes(ctx))} mots, répartis en 3 à 7
   scènes. Chaque scène a un illustration_brief (une phrase,
   style: ${JSON.stringify(ctx.world?.style_bible ?? {})}).
Réponds UNIQUEMENT avec le JSON demandé.`;

  const user = {
    enfant: {
      prenom: ctx.child.first_name,
      age: ctx.age,
      passions: ctx.child.passions,
      peurs_en_travail: ctx.child.fears,
    },
    compagnon: ctx.companion
      ? { nom: ctx.companion.name, adn: ctx.companion.dna, mots_appris: ctx.companion.learned_words }
      : null,
    famille_castee: ctx.cast,
    monde: { nom: ctx.world?.name, saison: ctx.world?.season },
    entites_canoniques: ctx.entities,
    evenements_recents: ctx.events,
    fil_emotionnel_du_jour: ctx.checkin
      ? {
        qualite: ctx.checkin.day_quality,
        emotion: ctx.checkin.dominant_emotion,
        victoire: ctx.checkin.victory,
        peur: ctx.checkin.fear_of_the_day,
      }
      : null,
    graines_a_recolter: ctx.ripeSeeds.map((s) => ({ id: s.id, resume: s.summary })),
    rappel_du_dernier_episode: ctx.previousRecap,
    numero_episode: ctx.nextNumber,
    format_attendu: {
      title: "string",
      emotional_thread: "string|null — résumé du filigrane, pour le parent",
      scenes: [{ text: "string", illustration_brief: "string" }],
      world_events: ["résumés canoniques de ce qui s'est passé ce soir"],
      harvested_seed_ids: ["ids des graines tenues ce soir"],
      next_recap: "le rappel « La dernière fois… » du prochain épisode",
    },
  };

  const response = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 4096,
    system,
    messages: [{ role: "user", content: JSON.stringify(user) }],
  });
  const text = response.content.find((b) => b.type === "text")?.text ?? "";
  return JSON.parse(text.slice(text.indexOf("{"), text.lastIndexOf("}") + 1));
}

// ─── 3. Validation ───────────────────────────────────────────────────────────

function validate(plan: EpisodePlan, ctx: { ripeSeeds: { id: string }[] }) {
  if (!plan.title || !plan.scenes?.length) throw new Error("plan incomplet");
  if (plan.scenes.length > 8) throw new Error("trop de scènes");
  if (!plan.next_recap) throw new Error("next_recap manquant (feuilleton)");
  const knownSeeds = new Set(ctx.ripeSeeds.map((s) => s.id));
  for (const id of plan.harvested_seed_ids ?? []) {
    if (!knownSeeds.has(id)) throw new Error(`graine inconnue récoltée: ${id}`);
  }
}

// ─── 4. Commit canonique ─────────────────────────────────────────────────────

async function commit(
  childId: string,
  ctx: { world: { id: string } | null; nextNumber: number },
  plan: EpisodePlan,
): Promise<string> {
  const { data: episode, error } = await supabase
    .from("episodes")
    .insert({
      child_id: childId,
      number: ctx.nextNumber,
      title: plan.title,
      emotional_thread: plan.emotional_thread,
      next_recap: plan.next_recap,
      status: "ready",
    })
    .select("id").single();
  if (error) throw error;

  await supabase.from("scenes").insert(
    plan.scenes.map((scene, index) => ({
      episode_id: episode.id,
      index,
      text: scene.text,
    })),
  );

  if (ctx.world) {
    await supabase.from("world_events").insert(
      plan.world_events.map((summary) => ({
        world_id: ctx.world!.id,
        episode_number: ctx.nextNumber,
        summary,
      })),
    );
  }

  if (plan.harvested_seed_ids?.length) {
    await supabase.from("narrative_seeds")
      .update({ harvested_at: new Date().toISOString(), harvested_episode: ctx.nextNumber })
      .in("id", plan.harvested_seed_ids);
  }

  // TODO pipeline média : TTS multi-voix + illustrations depuis les briefs,
  // upload en Storage, puis notification « épisode prêt ».
  return episode.id;
}

// ─── Utilitaires ─────────────────────────────────────────────────────────────

const prio = (p: string) => ({ low: 0, normal: 1, high: 2 }[p] ?? 1);
const today = () => new Date().toISOString().slice(0, 10);
const wordsForMinutes = (min: number) => Math.round(min * 120); // voix grave, débit lent
// Durée : choix du parent (children.story_minutes, borné 1-5 min),
// sinon durée par défaut selon l'âge — miroir du Pacte de sommeil de l'app.
const plannedMinutes = (ctx: { child: { story_minutes?: number }; age: number }) => {
  const parent = ctx.child.story_minutes;
  if (parent) return Math.min(5, Math.max(1, parent));
  return ctx.age <= 4 ? 2 : ctx.age <= 7 ? 3 : ctx.age <= 9 ? 4 : 5;
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });

async function one(table: string, key: string, value: string) {
  const { data } = await supabase.from(table).select("*").eq(key, value).maybeSingle();
  return data;
}
async function many(table: string, key: string, value: string) {
  const { data } = await supabase.from(table).select("*").eq(key, value);
  return data ?? [];
}
