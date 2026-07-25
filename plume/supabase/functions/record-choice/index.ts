// PLUME — Edge Function `record-choice` (innovation #4)
// Le choix de l'enfant devient une graine narrative : une promesse plantée
// dans le canon, qui germera dans 3 à 12 semaines. C'est ce délai qui crée
// la magie (« le dragon sauvé en janvier revient en juin »).

import { createClient } from "npm:@supabase/supabase-js@2";

const supabase = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

Deno.serve(async (req) => {
  const { episode_id, option_id, seed_summary } = await req.json();
  if (!episode_id || !seed_summary) {
    return new Response(JSON.stringify({ error: "paramètres manquants" }), { status: 400 });
  }

  const { data: episode } = await supabase
    .from("episodes").select("child_id, number").eq("id", episode_id).maybeSingle();
  if (!episode) {
    return new Response(JSON.stringify({ error: "épisode inconnu" }), { status: 404 });
  }

  // Germination entre 21 et 84 jours : assez loin pour être oublié,
  // assez proche pour être reconnu avec émerveillement.
  const days = 21 + Math.floor(Math.random() * 63);
  const germinate = new Date(Date.now() + days * 86_400_000);

  await supabase.from("narrative_seeds").insert({
    child_id: episode.child_id,
    kind: "choiceConsequence",
    summary: seed_summary,
    germinate_after: germinate.toISOString(),
  });

  await supabase.from("world_events").insert({
    world_id: (await supabase.from("worlds").select("id")
      .eq("child_id", episode.child_id).single()).data!.id,
    episode_number: episode.number,
    summary: `Choix (${option_id}) : ${seed_summary}`,
  });

  return new Response(JSON.stringify({ planted: true, germinates_in_days: days }), {
    headers: { "content-type": "application/json" },
  });
});
