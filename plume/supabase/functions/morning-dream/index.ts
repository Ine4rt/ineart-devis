// PLUME — Edge Function `morning-dream` (innovation #15)
// Chaque matin, le compagnon « a rêvé » : un teaser de 2 phrases de l'épisode
// du soir, envoyé en notification douce. La boucle d'attente de Netflix,
// pour les 4-9 ans — l'enfant y pense toute la journée.
// Déclenchée par pg_cron vers 7 h 30 heure locale de la famille.

import Anthropic from "npm:@anthropic-ai/sdk";

import { guardChildAccess, serviceClient } from "../_shared/guard.ts";

// Clé service (RLS court-circuitée) : la garde tourne avant toute lecture.
const supabase = serviceClient();
const anthropic = new Anthropic({ apiKey: Deno.env.get("ANTHROPIC_API_KEY")! });

Deno.serve(async (req) => {
  const { child_id } = await req.json();
  if (!child_id) {
    return new Response(JSON.stringify({ error: "child_id requis" }), {
      status: 400,
      headers: { "content-type": "application/json" },
    });
  }

  // Sécurité : le rêve du matin révèle le début de l'épisode du soir d'un
  // enfant. Sans cette garde, un uuid deviné suffisait à le lire.
  const guard = await guardChildAccess(req, child_id, supabase);
  if (!guard.ok) return guard.response;

  const [{ data: companion }, { data: episode }] = await Promise.all([
    supabase.from("companions").select("name, dna").eq("child_id", child_id).maybeSingle(),
    supabase.from("episodes").select("title, id").eq("child_id", child_id)
      .eq("status", "ready").order("number", { ascending: false }).limit(1).maybeSingle(),
  ]);

  if (!companion || !episode) {
    return new Response(JSON.stringify({ skipped: true }), { status: 200 });
  }

  const { data: scenes } = await supabase
    .from("scenes").select("text").eq("episode_id", episode.id).order("index").limit(2);

  const response = await anthropic.messages.create({
    model: "claude-haiku-4-5-20251001", // tâche courte : le petit modèle suffit
    max_tokens: 200,
    system:
      `Tu es ${companion.name}, un ${companion.dna?.species ?? "compagnon"} adorable. ` +
      "Raconte en 2 phrases maximum, à la première personne, un rêve mystérieux et " +
      "joyeux qui donne TRÈS envie de découvrir l'histoire de ce soir — sans rien " +
      "révéler d'important. Ton doux, langage d'enfant, une pointe de malice.",
    messages: [{
      role: "user",
      content: `Début de l'épisode de ce soir (« ${episode.title} ») : ` +
        (scenes ?? []).map((s) => s.text).join(" ").slice(0, 600),
    }],
  });

  const dream = response.content.find((b) => b.type === "text")?.text?.trim() ?? "";

  // TODO: push FCM/APNs vers l'appareil du parent (titre : « Le rêve de Pipo »).
  return new Response(JSON.stringify({ dream }), {
    headers: { "content-type": "application/json" },
  });
});
