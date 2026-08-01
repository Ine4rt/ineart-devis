"use server";

import type { DocumentCategory } from "@prisma/client";
import { revalidatePath } from "next/cache";

import { requireUser } from "@/lib/auth";
import { DOCUMENT_CATEGORY } from "@/lib/constants";
import { db } from "@/lib/db";
import { deleteStoredFile, saveUploadedFile } from "@/lib/storage";
import { enumOf, logActivity, optStr, str } from "@/lib/actions/helpers";

const CATEGORIES = Object.keys(DOCUMENT_CATEGORY) as DocumentCategory[];

export async function uploadDocument(formData: FormData) {
  await requireUser();

  const file = formData.get("file");
  if (!(file instanceof File) || file.size === 0) {
    throw new Error("Aucun fichier sélectionné.");
  }

  const stored = await saveUploadedFile(file);
  const clientId = optStr(formData, "clientId");

  const document = await db.document.create({
    data: {
      // Le nom affiché est modifiable ; le nom d'origine reste conservé.
      name: str(formData, "name") || stored.originalName,
      originalName: stored.originalName,
      mimeType: stored.mimeType,
      size: stored.size,
      storageKey: stored.storageKey,
      category: enumOf(formData, "category", CATEGORIES, "OTHER"),
      notes: optStr(formData, "notes"),
      clientId,
      projectId: optStr(formData, "projectId"),
    },
  });

  await logActivity({
    entityType: "document",
    entityId: document.id,
    clientId,
    action: "uploaded",
    summary: `Document « ${document.name} » ajouté pour`,
  });

  revalidatePath("/documents");
  if (clientId) revalidatePath(`/clients/${clientId}/documents`);
}

export async function deleteDocument(documentId: string) {
  await requireUser();

  const document = await db.document.delete({ where: { id: documentId } });
  await deleteStoredFile(document.storageKey);

  revalidatePath("/documents");
  if (document.clientId) revalidatePath(`/clients/${document.clientId}/documents`);
}

export async function updateDocument(documentId: string, formData: FormData) {
  await requireUser();

  const document = await db.document.update({
    where: { id: documentId },
    data: {
      name: str(formData, "name"),
      category: enumOf(formData, "category", CATEGORIES, "OTHER"),
      notes: optStr(formData, "notes"),
    },
  });

  revalidatePath("/documents");
  if (document.clientId) revalidatePath(`/clients/${document.clientId}/documents`);
}
