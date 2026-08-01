"use server";

import type { ProjectStatus, ProjectType } from "@prisma/client";
import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";

import { requireUser } from "@/lib/auth";
import { PROJECT_STATUS, PROJECT_TYPE } from "@/lib/constants";
import { db } from "@/lib/db";
import { bool, date, enumOf, int, jsonList, logActivity, optStr, str } from "@/lib/actions/helpers";

const STATUSES = Object.keys(PROJECT_STATUS) as ProjectStatus[];
const TYPES = Object.keys(PROJECT_TYPE) as ProjectType[];

function readProjectFields(formData: FormData) {
  return {
    name: str(formData, "name"),
    url: optStr(formData, "url"),
    stagingUrl: optStr(formData, "stagingUrl"),
    description: optStr(formData, "description"),
    type: enumOf(formData, "type", TYPES, "VITRINE"),
    status: enumOf(formData, "status", STATUSES, "DISCOVERY"),
    progress: Math.max(0, Math.min(100, int(formData, "progress"))),
    isCustom: bool(formData, "isCustom"),
    cms: optStr(formData, "cms"),
    techStack: jsonList(formData, "techStack"),
    repositoryUrl: optStr(formData, "repositoryUrl"),
    notes: optStr(formData, "notes"),
    startedAt: date(formData, "startedAt"),
    launchedAt: date(formData, "launchedAt"),
    deliveredAt: date(formData, "deliveredAt"),
  };
}

export async function createProject(formData: FormData) {
  await requireUser();

  const clientId = str(formData, "clientId");
  const fields = readProjectFields(formData);
  if (!clientId || !fields.name) throw new Error("Client et nom du projet sont obligatoires.");

  const project = await db.project.create({ data: { ...fields, clientId } });

  // Le premier jalon est créé d'office : l'historique d'un projet commence à
  // sa création, pas à la première modification qu'on pense à consigner.
  await db.milestone.create({
    data: { projectId: project.id, title: "Projet créé", happenedAt: new Date() },
  });

  await logActivity({
    entityType: "project",
    entityId: project.id,
    clientId,
    action: "created",
    summary: `Projet « ${project.name} » créé pour`,
  });

  revalidatePath("/projets");
  redirect(`/projets/${project.id}`);
}

export async function updateProject(projectId: string, formData: FormData) {
  await requireUser();

  const fields = readProjectFields(formData);
  const previous = await db.project.findUnique({
    where: { id: projectId },
    select: { status: true, clientId: true },
  });

  const project = await db.project.update({ where: { id: projectId }, data: fields });

  // Un changement d'état est un événement du projet : il rejoint l'historique.
  if (previous && previous.status !== fields.status) {
    await db.milestone.create({
      data: {
        projectId,
        title: `Statut : ${PROJECT_STATUS[previous.status].label} → ${PROJECT_STATUS[fields.status].label}`,
        happenedAt: new Date(),
      },
    });
  }

  await logActivity({
    entityType: "project",
    entityId: projectId,
    clientId: project.clientId,
    action: "updated",
    summary: `Projet « ${project.name} » mis à jour pour`,
  });

  revalidatePath(`/projets/${projectId}`);
  revalidatePath("/projets");
  redirect(`/projets/${projectId}`);
}

export async function deleteProject(projectId: string) {
  await requireUser();
  const project = await db.project.delete({ where: { id: projectId } });
  revalidatePath("/projets");
  redirect(`/clients/${project.clientId}/projets`);
}

/** Mise à jour rapide de l'avancement depuis la liste, sans ouvrir la fiche. */
export async function setProjectProgress(projectId: string, progress: number) {
  await requireUser();
  await db.project.update({
    where: { id: projectId },
    data: { progress: Math.max(0, Math.min(100, Math.round(progress))) },
  });
  revalidatePath(`/projets/${projectId}`);
  revalidatePath("/projets");
}
