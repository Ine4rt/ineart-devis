import {
  createCipheriv,
  createDecipheriv,
  createHash,
  randomBytes,
  scryptSync,
  timingSafeEqual,
} from "node:crypto";

/**
 * Chiffrement des secrets stockés en base (identifiants FTP/SSH/CMS…).
 *
 * Algorithme : AES-256-GCM (chiffrement authentifié — une altération du
 * ciphertext est détectée au déchiffrement au lieu de produire des octets
 * silencieusement faux).
 *
 * La clé provient de APP_ENCRYPTION_KEY. Elle est dérivée par SHA-256 pour
 * accepter une phrase secrète de longueur quelconque tout en garantissant
 * 32 octets en sortie.
 *
 * ⚠️ Perdre APP_ENCRYPTION_KEY = perdre définitivement les secrets stockés.
 *    Le reste des données (clients, projets…) reste évidemment lisible.
 */

const ENCODING = "base64url";

function encryptionKey(): Buffer {
  const secret = process.env.APP_ENCRYPTION_KEY;
  if (!secret || secret.length < 16) {
    throw new Error(
      "APP_ENCRYPTION_KEY manquante ou trop courte (32 caractères minimum recommandés). " +
        "Générez-en une avec : openssl rand -base64 32",
    );
  }
  return createHash("sha256").update(secret).digest();
}

/** Chiffre une valeur en clair. Retourne `iv.tag.ciphertext` en base64url. */
export function encryptSecret(plain: string): string {
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", encryptionKey(), iv);
  const ciphertext = Buffer.concat([cipher.update(plain, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return [iv.toString(ENCODING), tag.toString(ENCODING), ciphertext.toString(ENCODING)].join(".");
}

/**
 * Déchiffre un payload produit par `encryptSecret`.
 * Retourne `null` si la valeur est absente, malformée ou si la clé ne correspond
 * pas — l'appelant affiche alors un message clair plutôt que de planter.
 */
export function decryptSecret(payload: string | null | undefined): string | null {
  if (!payload) return null;
  const parts = payload.split(".");
  if (parts.length !== 3) return null;

  try {
    const [iv, tag, ciphertext] = parts.map((p) => Buffer.from(p, ENCODING));
    const decipher = createDecipheriv("aes-256-gcm", encryptionKey(), iv);
    decipher.setAuthTag(tag);
    return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString("utf8");
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Mots de passe du compte
// ---------------------------------------------------------------------------

const SCRYPT_KEYLEN = 64;

/** Hache un mot de passe avec scrypt + sel aléatoire. Format `sel:hash`. */
export function hashPassword(password: string): string {
  const salt = randomBytes(16).toString("hex");
  const derived = scryptSync(password, salt, SCRYPT_KEYLEN).toString("hex");
  return `${salt}:${derived}`;
}

/** Vérifie un mot de passe en temps constant. */
export function verifyPassword(password: string, stored: string): boolean {
  const [salt, hash] = stored.split(":");
  if (!salt || !hash) return false;
  const derived = scryptSync(password, salt, SCRYPT_KEYLEN);
  const expected = Buffer.from(hash, "hex");
  if (expected.length !== derived.length) return false;
  return timingSafeEqual(derived, expected);
}
