import "server-only";

import { randomUUID } from "node:crypto";
import { mkdir, unlink, writeFile } from "node:fs/promises";
import path from "node:path";

/**
 * Stockage des fichiers joints.
 *
 * Les binaires vivent hors de `public/` : rien n'est servi statiquement, tout
 * passe par une route authentifiée. Un contrat client ne doit pas être
 * accessible à qui devine son URL.
 */

const STORAGE_ROOT = process.env.STORAGE_DIR
  ? path.resolve(process.env.STORAGE_DIR)
  : path.join(process.cwd(), "storage", "documents");

export const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20 Mo

/** Nettoie un nom de fichier : pas de séparateur, pas de traversée de chemin. */
function safeName(name: string): string {
  return path
    .basename(name)
    .replace(/[^a-zA-Z0-9._-]/g, "_")
    .slice(-80);
}

export async function saveUploadedFile(file: File): Promise<{
  storageKey: string;
  size: number;
  mimeType: string;
  originalName: string;
}> {
  if (file.size > MAX_FILE_SIZE) {
    throw new Error(`Fichier trop volumineux (maximum ${MAX_FILE_SIZE / 1024 / 1024} Mo).`);
  }

  await mkdir(STORAGE_ROOT, { recursive: true });

  const storageKey = `${randomUUID()}-${safeName(file.name)}`;
  const buffer = Buffer.from(await file.arrayBuffer());
  await writeFile(path.join(STORAGE_ROOT, storageKey), buffer);

  return {
    storageKey,
    size: file.size,
    mimeType: file.type || "application/octet-stream",
    originalName: file.name,
  };
}

/** Chemin absolu d'un fichier stocké, en refusant toute sortie du dossier. */
export function resolveStoragePath(storageKey: string): string {
  const resolved = path.resolve(STORAGE_ROOT, storageKey);
  if (!resolved.startsWith(STORAGE_ROOT)) {
    throw new Error("Chemin de stockage invalide.");
  }
  return resolved;
}

export async function deleteStoredFile(storageKey: string): Promise<void> {
  try {
    await unlink(resolveStoragePath(storageKey));
  } catch {
    /* fichier déjà absent : la suppression de la fiche reste valide */
  }
}
