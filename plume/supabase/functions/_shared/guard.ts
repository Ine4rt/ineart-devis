// PLUME — Garde d'accès partagée des Edge Functions.
//
// Toutes ces fonctions utilisent la clé service (SUPABASE_SERVICE_ROLE_KEY),
// qui COURT-CIRCUITE la RLS : sans vérification de l'appelant, n'importe qui
// pouvait générer des épisodes ou lire le rêve d'un enfant en devinant un
// uuid. Ce module est le point de passage obligé avant toute action.
//
// Deux appelants légitimes, et deux seulement :
//
//  1. LE PARENT, depuis l'app. Il présente son JWT utilisateur dans
//     `Authorization: Bearer <token>` (supabase-flutter le joint
//     automatiquement à `functions.invoke`). On vérifie le token, puis on
//     vérifie que l'enfant visé appartient bien à SA famille.
//
//  2. pg_cron, en interne (pré-génération nocturne des épisodes, rêves du
//     matin). Un job cron n'a pas de session utilisateur : il présente un
//     secret partagé dédié dans l'en-tête `x-plume-cron-secret`.
//
//       -- côté Postgres, le secret est lu depuis Vault, jamais en clair :
//       select net.http_post(
//         url     := 'https://<projet>.functions.supabase.co/generate-episode',
//         headers := jsonb_build_object(
//           'content-type',        'application/json',
//           'x-plume-cron-secret', vault.decrypted_secret('plume_cron_secret')
//         ),
//         body    := jsonb_build_object('child_id', c.id)
//       ) from children c;
//
//       -- et côté fonctions :
//       supabase secrets set PLUME_CRON_SECRET="$(openssl rand -hex 32)"
//
//     Ce secret est distinct de la clé service : il n'ouvre QUE ces fonctions,
//     et il se révoque sans rien recréer. Si `PLUME_CRON_SECRET` n'est pas
//     configuré, la voie interne est fermée (aucun repli permissif).

import { createClient, SupabaseClient } from "npm:@supabase/supabase-js@2";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY")!;
const CRON_SECRET = Deno.env.get("PLUME_CRON_SECRET");

export type GuardResult =
  | { ok: true; internal: boolean; userId: string | null }
  | { ok: false; response: Response };

/// Client de service : il ignore la RLS, donc il ne sort jamais d'ici sans
/// qu'une garde ait tourné avant.
export function serviceClient(): SupabaseClient {
  return createClient(
    SUPABASE_URL,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  );
}

/// Vrai si l'appel vient de pg_cron avec le bon secret partagé.
export function isInternalCall(req: Request): boolean {
  const presented = req.headers.get("x-plume-cron-secret");
  return Boolean(CRON_SECRET) && presented === CRON_SECRET;
}

/// Autorise l'accès à [childId], ou renvoie la réponse d'erreur à retourner
/// telle quelle. À appeler AVANT toute lecture ou écriture.
export async function guardChildAccess(
  req: Request,
  childId: string,
  admin: SupabaseClient,
): Promise<GuardResult> {
  // Voie 1 — pg_cron : secret partagé dédié, pas de session utilisateur.
  if (isInternalCall(req)) return { ok: true, internal: true, userId: null };

  // Voie 2 — le parent : JWT utilisateur obligatoire.
  const authorization = req.headers.get("Authorization") ?? "";
  const token = authorization.toLowerCase().startsWith("bearer ")
    ? authorization.slice(7).trim()
    : "";
  if (!token) {
    return { ok: false, response: deny(401, "authentification requise") };
  }

  // On valide le token avec la clé anonyme : impossible de se faire passer
  // pour un autre utilisateur, c'est Supabase Auth qui tranche.
  const { data: { user }, error } = await createClient(SUPABASE_URL, ANON_KEY)
    .auth.getUser(token);
  if (error || !user) {
    return { ok: false, response: deny(401, "session invalide") };
  }

  // L'enfant visé doit appartenir à la famille de CET utilisateur.
  const { data: child, error: childError } = await admin
    .from("children")
    .select("id, families!inner(owner_id)")
    .eq("id", childId)
    .maybeSingle();
  if (childError) {
    return { ok: false, response: deny(500, "verification_impossible") };
  }

  const owner =
    (child as { families?: { owner_id?: string } } | null)?.families?.owner_id;
  // Même message et même code pour « inconnu » et « pas à vous » : on ne
  // laisse pas deviner l'existence d'un enfant par énumération d'uuid.
  if (!child || owner !== user.id) {
    return { ok: false, response: deny(403, "accès refusé") };
  }

  return { ok: true, internal: false, userId: user.id };
}

/// Variante pour les fonctions qui ne connaissent qu'un épisode : on remonte
/// à l'enfant, puis on applique la même garde.
export async function guardEpisodeAccess(
  req: Request,
  episodeId: string,
  admin: SupabaseClient,
): Promise<GuardResult & { childId?: string }> {
  const { data: episode } = await admin
    .from("episodes").select("child_id").eq("id", episodeId).maybeSingle();
  if (!episode) {
    // Pas d'information sur l'existence de l'épisode avant authentification.
    if (!isInternalCall(req) && !req.headers.get("Authorization")) {
      return { ok: false, response: deny(401, "authentification requise") };
    }
    return { ok: false, response: deny(404, "épisode inconnu") };
  }
  const guard = await guardChildAccess(req, episode.child_id, admin);
  return { ...guard, childId: episode.child_id };
}

function deny(status: number, error: string): Response {
  return new Response(JSON.stringify({ error }), {
    status,
    headers: { "content-type": "application/json" },
  });
}
