// PLUME — Edge Function `narrate-episode`
// La voix de Plume : narration humaine, calme, chaleureuse — GRATUITE.
// Appelée juste après generate-episode (la nuit) : chaque scène est convertie
// en audio, stockée dans Supabase Storage, et l'app ne fait que lire des
// mp3/wav — aucun TTS robotique côté client.
//
// Fournisseur par défaut (gratuit, open source) : un serveur Piper
// auto-hébergé (https://github.com/rhasspy/piper) exposé en HTTP —
// une petite VM à 5 €/mois narre des milliers d'épisodes par nuit ;
// voix françaises libres : fr_FR-tom (homme), fr_FR-siwis (femme),
// fr_FR-upmc (deux locuteurs). Validé sur le prototype.
//   supabase secrets set TTS_SERVER_URL=https://tts.example.com/api/tts
//
// Option premium (payante, encore plus expressive) : ElevenLabs —
// utilisée seulement si la clé est fournie.
//   supabase secrets set ELEVENLABS_API_KEY=... ELEVENLABS_VOICE_ID=...

import { createClient } from "npm:@supabase/supabase-js@2";

const supabase = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

const PIPER_URL = Deno.env.get("TTS_SERVER_URL"); // gratuit, par défaut
const ELEVEN_KEY = Deno.env.get("ELEVENLABS_API_KEY"); // optionnel
const VOICE_ID = Deno.env.get("ELEVENLABS_VOICE_ID");

const ELEVEN_URL = (voiceId: string) =>
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
  // 1. Piper auto-hébergé : gratuit, illimité, hors des API payantes.
  if (PIPER_URL) {
    const response = await fetch(PIPER_URL, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        text,
        voice: "fr_FR-tom-medium", // homme, calme ; siwis-medium pour femme
        length_scale: 1.2, // débit du soir
        sentence_silence: 0.45,
      }),
    });
    if (!response.ok) {
      throw new Error(`Piper ${response.status}: ${await response.text()}`);
    }
    return new Uint8Array(await response.arrayBuffer());
  }

  // 2. Option premium ElevenLabs, seulement si la clé est configurée.
  if (ELEVEN_KEY && VOICE_ID) {
    const response = await fetch(ELEVEN_URL(VOICE_ID), {
      method: "POST",
      headers: { "xi-api-key": ELEVEN_KEY, "content-type": "application/json" },
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

  throw new Error("aucun fournisseur TTS configuré (TTS_SERVER_URL ou ELEVENLABS_API_KEY)");
}

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
