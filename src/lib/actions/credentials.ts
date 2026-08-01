"use server";

import type { CredentialKind } from "@prisma/client";
import { revalidatePath } from "next/cache";

import { requireUser } from "@/lib/auth";
import { CREDENTIAL_KIND } from "@/lib/constants";
import { decryptSecret, encryptSecret } from "@/lib/crypto";
import { db } from "@/lib/db";
import { enumOf, optStr, str } from "@/lib/actions/helpers";

const KINDS = Object.keys(CREDENTIAL_KIND) as CredentialKind[];

/**
 * Coffre à identifiants.
 *
 * Le secret est chiffré à l'écriture et n'est déchiffré que par `revealSecret`,
 * sur demande explicite. Aucun rendu de page ne renvoie de mot de passe en
 * clair : même une fuite du HTML ne les exposerait pas.
 */

export async function createCredential(formData: FormData) {
  await requireUser();

  const label = str(formData, "label");
  const secret = str(formData, "secret");
  if (!label) throw new Error("Le libellé est obligatoire.");

  await db.credential.create({
    data: {
      label,
      kind: enumOf(formData, "kind", KINDS, "OTHER"),
      username: optStr(formData, "username"),
      secretEnc: secret ? encryptSecret(secret) : null,
      url: optStr(formData, "url"),
      notes: optStr(formData, "notes"),
      clientId: optStr(formData, "clientId"),
      projectId: optStr(formData, "projectId"),
      hostingId: optStr(formData, "hostingId"),
      domainId: optStr(formData, "domainId"),
    },
  });

  revalidatePathsFor(formData);
}

export async function updateCredential(credentialId: string, formData: FormData) {
  await requireUser();

  const secret = str(formData, "secret");

  await db.credential.update({
    where: { id: credentialId },
    data: {
      label: str(formData, "label"),
      kind: enumOf(formData, "kind", KINDS, "OTHER"),
      username: optStr(formData, "username"),
      // Un champ laissé vide ne doit pas effacer le secret existant : on ne
      // remplace que si une nouvelle valeur a été saisie.
      ...(secret ? { secretEnc: encryptSecret(secret) } : {}),
      url: optStr(formData, "url"),
      notes: optStr(formData, "notes"),
    },
  });

  revalidatePathsFor(formData);
}

export async function deleteCredential(credentialId: string, revalidate: string) {
  await requireUser();
  await db.credential.delete({ where: { id: credentialId } });
  revalidatePath(revalidate);
}

/**
 * Déchiffre un secret à la demande. Seul point de sortie en clair de
 * l'application — appelé depuis un bouton « Révéler », jamais au rendu.
 */
export async function revealSecret(credentialId: string): Promise<{ value: string | null; error?: string }> {
  await requireUser();

  const credential = await db.credential.findUnique({
    where: { id: credentialId },
    select: { secretEnc: true },
  });

  if (!credential?.secretEnc) return { value: null, error: "Aucun secret enregistré." };

  const value = decryptSecret(credential.secretEnc);
  if (value === null) {
    return {
      value: null,
      error: "Déchiffrement impossible — la clé APP_ENCRYPTION_KEY a-t-elle changé ?",
    };
  }

  return { value };
}

function revalidatePathsFor(formData: FormData) {
  const path = str(formData, "revalidate");
  if (path) revalidatePath(path);
}
