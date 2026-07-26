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

import { guardEpisodeAccess, serviceClient } from "../_shared/guard.ts";

// Clé service (RLS court-circuitée) : la garde tourne avant toute action.
const supabase = serviceClient();

const PIPER_URL = Deno.env.get("TTS_SERVER_URL"); // gratuit, par défaut
const ELEVEN_KEY = Deno.env.get("ELEVENLABS_API_KEY"); // optionnel
const VOICE_ID = Deno.env.get("ELEVENLABS_VOICE_ID");

const ELEVEN_URL = (voiceId: string) =>
  `https://api.elevenlabs.io/v1/text-to-speech/${voiceId}?output_format=mp3_44100_64`;

// 30 jours : largement plus qu'un cycle veille/génération, sans être éternelle.
const SIGNED_URL_TTL_SECONDS = 60 * 60 * 24 * 30;

Deno.serve(async (req) => {
  try {
    const { episode_id } = await req.json();
    if (!episode_id) return json({ error: "episode_id requis" }, 400);

    // Sécurité : narrer un épisode consomme du TTS et expose son texte.
    // On remonte à l'enfant et on vérifie qu'il appartient à l'appelant
    // (parent authentifié) ou que l'appel vient de pg_cron.
    const guard = await guardEpisodeAccess(req, episode_id, supabase);
    if (!guard.ok) return guard.response;

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

      // Bucket privé : une URL signée, pas une URL publique — la voix
      // personnalisée d'un enfant ne doit pas être accessible par un uuid
      // deviné. Longue durée de vie : l'app la met en cache pour le soir.
      const { data: signed, error: signErr } = await supabase.storage
        .from("narration")
        .createSignedUrl(path, SIGNED_URL_TTL_SECONDS);
      if (signErr) throw signErr;

      await supabase.from("scenes")
        .update({ audio_url: signed.signedUrl })
        .eq("id", scene.id);
    }

    return json({ narrated: scenes.length });
  } catch (error) {
    console.error("narrate-episode failed:", error);
    return json({ error: "narration_failed" }, 500);
  }
});

/// Piper d'abord (gratuit), ElevenLabs en repli. Le repli était inatteignable :
/// un Piper en panne faisait `throw` au lieu de basculer, et la narration
/// échouait alors même qu'une clé premium était configurée.
async function synthesize(text: string): Promise<Uint8Array> {
  if (!PIPER_URL && !(ELEVEN_KEY && VOICE_ID)) {
    throw new Error(
      "aucun fournisseur TTS configuré (TTS_SERVER_URL ou ELEVENLABS_API_KEY)",
    );
  }

  // 1. Piper auto-hébergé : gratuit, illimité, hors des API payantes.
  if (PIPER_URL) {
    try {
      return await synthesizeWithPiper(text);
    } catch (piperError) {
      // Pas de clé premium : rien à tenter de plus, on remonte l'échec.
      if (!(ELEVEN_KEY && VOICE_ID)) throw piperError;
      console.warn("Piper indisponible, repli sur ElevenLabs:", piperError);
    }
  }

  // 2. Option premium ElevenLabs : repli, ou fournisseur unique.
  return await synthesizeWithElevenLabs(text);
}

async function synthesizeWithPiper(text: string): Promise<Uint8Array> {
  const response = await fetch(PIPER_URL!, {
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

async function synthesizeWithElevenLabs(text: string): Promise<Uint8Array> {
  const response = await fetch(ELEVEN_URL(VOICE_ID!), {
    method: "POST",
    headers: { "xi-api-key": ELEVEN_KEY!, "content-type": "application/json" },
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
    throw new Error(`ElevenLabs ${response.status}: ${await response.text()}`);
  }
  return new Uint8Array(await response.arrayBuffer());
}

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
