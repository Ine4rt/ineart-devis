import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import { NextResponse } from "next/server";
import type { ReadableStream as WebReadableStream } from "node:stream/web";
import { Readable } from "node:stream";

import { getSessionUser } from "@/lib/auth";
import { db } from "@/lib/db";
import { resolveStoragePath } from "@/lib/storage";

/**
 * Téléchargement d'un document. Seule porte de sortie des fichiers : la session
 * est vérifiée avant toute lecture disque, et le fichier est diffusé en flux
 * plutôt que chargé entièrement en mémoire.
 */
export async function GET(
  _request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  const user = await getSessionUser();
  if (!user) return new NextResponse("Non autorisé", { status: 401 });

  const { id } = await params;
  const document = await db.document.findUnique({ where: { id } });
  if (!document) return new NextResponse("Introuvable", { status: 404 });

  let filePath: string;
  try {
    filePath = resolveStoragePath(document.storageKey);
    await stat(filePath);
  } catch {
    return new NextResponse("Fichier absent du stockage", { status: 410 });
  }

  const stream = Readable.toWeb(createReadStream(filePath)) as WebReadableStream<Uint8Array>;

  return new NextResponse(stream as unknown as ReadableStream, {
    headers: {
      "Content-Type": document.mimeType,
      "Content-Length": String(document.size),
      // `inline` : les PDF et images s'ouvrent dans l'onglet, le reste se télécharge.
      "Content-Disposition": `inline; filename="${encodeURIComponent(document.originalName)}"`,
      "Cache-Control": "private, no-store",
    },
  });
}
