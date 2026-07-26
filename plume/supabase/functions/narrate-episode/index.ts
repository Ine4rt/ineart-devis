// PLUME — Edge Function `narrate-episode`
// La voix de Plume : narration studio, profonde, calme, vraiment humaine.
// Appelée juste après generate-episode (la nuit) : chaque scène est convertie
// en audio via un TTS neuronal (ElevenLabs), stockée dans Supabase Storage,
// et l'app ne fait que lire des mp3 — aucun TTS robotique côté client.
//
// Secrets requis :
//   supabase secrets set ELEVENLABS_API_KEY=...
//   supabase secrets set ELEVENLABS_VOICE_ID=...   (voix française grave et chaleureuse)

import { createClient } from "npm:@supabase/supabase-js@2";

const supabase = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

const ELEVEN_KEY = Deno.env.get("ELEVENLABS_API_KEY")!;
// Voix cible : timbre masculin grave, débit lent, chaleur de conteur.
// À choisir dans la bibliothèque ElevenLabs (ou une voix clonée maison).
const VOICE_ID = Deno.env.get("ELEVENLABS_VOICE_ID")!;

const TTS_URL = (voiceId: string) =>
  `https://api.elevenlabs.io/v1/text-to-speech/${voiceId}?output_format=mp3_44100_64`;

Deno.serve(async (req) => {
  try {
    const { episode_id } = await req.json();
    if (!episode_id) return json({ error: "episode_id requis" }, 400);

    const { data: scenes, error } = await supabase
      .from("scenes")
      .select("id, index, text, audio_url")
      .eq("episode_id", episode_id)
      .order("index");
    if (error) throw error;
    if (!scenes?.length) return json({ error: "aucune scène" }, 404);

    for (const scene of scenes) {
      if (scene.audio_url) continue; // déjà narré (reprise après échec)

      const audio = await synthesize(scene.text);
      const path = `episodes/${episode_id}/scene-${scene.index}.mp3`;

      const { error: upErr } = await supabase.storage
        .from("narration")
        .upload(path, audio, { contentType: "audio/mpeg", upsert: true });
      if (upErr) throw upErr;

      const { data: pub } = supabase.storage.from("narration").getPublicUrl(path);
      await supabase.from("scenes")
        .update({ audio_url: pub.publicUrl })
        .eq("id", scene.id);
    }

    return json({ narrated: scenes.length });
  } catch (error) {
    console.error("narrate-episode failed:", error);
    return json({ error: "narration_failed" }, 500);
  }
});

async function synthesize(text: string): Promise<Uint8Array> {
  const response = await fetch(TTS_URL(VOICE_ID), {
    method: "POST",
    headers: {
      "xi-api-key": ELEVEN_KEY,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      text,
      model_id: "eleven_multilingual_v2",
      voice_settings: {
        stability: 0.65, // posé, sans monotonie
        similarity_boost: 0.8,
        style: 0.25, // légère intention de conteur, jamais théâtral
        speed: 0.9, // débit du soir
      },
    }),
  });
  if (!response.ok) {
    throw new Error(`TTS ${response.status}: ${await response.text()}`);
  }
  return new Uint8Array(await response.arrayBuffer());
}

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
