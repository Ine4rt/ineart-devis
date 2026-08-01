"use server";

import type { ClientStatus, PipelineStage } from "@prisma/client";
import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";

import { requireUser } from "@/lib/auth";
import { CLIENT_STATUS, PIPELINE_STAGE } from "@/lib/constants";
import { db } from "@/lib/db";
import {
  bool,
  date,
  dec,
  logActivity,
  nextClientReference,
  optStr,
  str,
} from "@/lib/actions/helpers";

const CLIENT_STATUSES = Object.keys(CLIENT_STATUS) as ClientStatus[];
const PIPELINE_STAGES = Object.keys(PIPELINE_STAGE) as PipelineStage[];

/** Champs partagés par la création et la modification. */
function readClientFields(formData: FormData) {
  const status = (str(formData, "status") || "PROSPECT") as ClientStatus;
  const pipelineStage = (str(formData, "pipelineStage") || "IDENTIFIED") as PipelineStage;

  return {
    company: str(formData, "company"),
    firstName: optStr(formData, "firstName"),
    lastName: optStr(formData, "lastName"),
    email: optStr(formData, "email"),
    phone: optStr(formData, "phone"),
    mobile: optStr(formData, "mobile"),
    website: optStr(formData, "website"),
    addressLine1: optStr(formData, "addressLine1"),
    addressLine2: optStr(formData, "addressLine2"),
    postalCode: optStr(formData, "postalCode"),
    city: optStr(formData, "city"),
    country: optStr(formData, "country"),
    vatNumber: optStr(formData, "vatNumber"),
    companyNumber: optStr(formData, "companyNumber"),
    industry: optStr(formData, "industry"),
    logoUrl: optStr(formData, "logoUrl"),
    source: optStr(formData, "source"),
    notes: optStr(formData, "notes"),
    status: CLIENT_STATUSES.includes(status) ? status : "PROSPECT",
    // L'étape de pipeline n'a de sens que pour un prospect : on la neutralise
    // sinon, pour éviter des données contradictoires dans le kanban.
    pipelineStage:
      status === "PROSPECT" && PIPELINE_STAGES.includes(pipelineStage) ? pipelineStage : null,
    potentialValue: dec(formData, "potentialValue") || null,
  };
}

export async function createClient(formData: FormData) {
  await requireUser();
  const fields = readClientFields(formData);

  if (!fields.company) throw new Error("La société est obligatoire.");

  const client = await db.client.create({
    data: { ...fields, reference: await nextClientReference() },
  });

  await logActivity({
    entityType: "client",
    entityId: client.id,
    clientId: client.id,
    action: "created",
    summary: "Nouveau client créé :",
  });

  revalidatePath("/clients");
  redirect(`/clients/${client.id}`);
}

export async function updateClient(clientId: string, formData: FormData) {
  await requireUser();
  const fields = readClientFields(formData);

  if (!fields.company) throw new Error("La société est obligatoire.");

  const previous = await db.client.findUnique({
    where: { id: clientId },
    select: { status: true },
  });

  await db.client.update({
    where: { id: clientId },
    data: {
      ...fields,
      // Horodate la conversion pour pouvoir mesurer le délai de closing.
      wonAt: fields.status === "ACTIVE" && previous?.status !== "ACTIVE" ? new Date() : undefined,
      lostAt: fields.status === "LOST" && previous?.status !== "LOST" ? new Date() : undefined,
    },
  });

  await logActivity({
    entityType: "client",
    entityId: clientId,
    clientId,
    action: "updated",
    summary: "Fiche mise à jour :",
  });

  revalidatePath(`/clients/${clientId}`);
  revalidatePath("/clients");
  redirect(`/clients/${clientId}`);
}

/** Déplacement d'un prospect dans le pipeline (glisser-déposer du kanban). */
export async function moveProspect(clientId: string, stage: PipelineStage) {
  await requireUser();
  if (!PIPELINE_STAGES.includes(stage)) return;

  const isWon = stage === "WON";
  const isLost = stage === "LOST";

  await db.client.update({
    where: { id: clientId },
    data: {
      pipelineStage: stage,
      // Gagner une opportunité fait basculer le prospect en client actif :
      // le pipeline et la base clients ne peuvent pas se contredire.
      status: isWon ? "ACTIVE" : isLost ? "LOST" : "PROSPECT",
      wonAt: isWon ? new Date() : null,
      lostAt: isLost ? new Date() : null,
    },
  });

  await logActivity({
    entityType: "client",
    entityId: clientId,
    clientId,
    action: "pipeline",
    summary: `Passé à l'étape « ${PIPELINE_STAGE[stage].label} » :`,
  });

  revalidatePath("/prospection");
  revalidatePath("/clients");
}

export async function deleteClient(clientId: string) {
  await requireUser();
  await db.client.delete({ where: { id: clientId } });
  revalidatePath("/clients");
  redirect("/clients");
}

// ---------------------------------------------------------------------------
// Interlocuteurs
// ---------------------------------------------------------------------------

export async function createContact(clientId: string, formData: FormData) {
  await requireUser();

  await db.contact.create({
    data: {
      clientId,
      firstName: str(formData, "firstName"),
      lastName: optStr(formData, "lastName"),
      role: optStr(formData, "role"),
      email: optStr(formData, "email"),
      phone: optStr(formData, "phone"),
      isPrimary: bool(formData, "isPrimary"),
      notes: optStr(formData, "notes"),
    },
  });

  revalidatePath(`/clients/${clientId}`);
}

export async function deleteContact(contactId: string, clientId: string) {
  await requireUser();
  await db.contact.delete({ where: { id: contactId } });
  revalidatePath(`/clients/${clientId}`);
}

// ---------------------------------------------------------------------------
// Notes
// ---------------------------------------------------------------------------

export async function createNote(formData: FormData) {
  await requireUser();

  const clientId = optStr(formData, "clientId");
  const body = str(formData, "body");
  if (!body) return;

  await db.note.create({
    data: {
      clientId,
      projectId: optStr(formData, "projectId"),
      title: optStr(formData, "title"),
      body,
      pinned: bool(formData, "pinned"),
    },
  });

  if (clientId) revalidatePath(`/clients/${clientId}/suivi`);
}

export async function toggleNotePin(noteId: string, clientId: string | null) {
  await requireUser();
  const note = await db.note.findUnique({ where: { id: noteId }, select: { pinned: true } });
  if (!note) return;

  await db.note.update({ where: { id: noteId }, data: { pinned: !note.pinned } });
  if (clientId) revalidatePath(`/clients/${clientId}/suivi`);
}

export async function deleteNote(noteId: string, clientId: string | null) {
  await requireUser();
  await db.note.delete({ where: { id: noteId } });
  if (clientId) revalidatePath(`/clients/${clientId}/suivi`);
}

// ---------------------------------------------------------------------------
// Jalons de projet (historique des modifications)
// ---------------------------------------------------------------------------

export async function createMilestone(projectId: string, formData: FormData) {
  await requireUser();
  const title = str(formData, "title");
  if (!title) return;

  await db.milestone.create({
    data: {
      projectId,
      title,
      detail: optStr(formData, "detail"),
      happenedAt: date(formData, "happenedAt") ?? new Date(),
    },
  });

  revalidatePath(`/projets/${projectId}`);
}
