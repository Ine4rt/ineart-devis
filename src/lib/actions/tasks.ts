"use server";

import type { TaskKind, TaskPriority, TaskStatus } from "@prisma/client";
import { revalidatePath } from "next/cache";

import { requireUser } from "@/lib/auth";
import { TASK_KIND, TASK_PRIORITY, TASK_STATUS } from "@/lib/constants";
import { db } from "@/lib/db";
import { date, enumOf, optStr, str } from "@/lib/actions/helpers";

const STATUSES = Object.keys(TASK_STATUS) as TaskStatus[];
const PRIORITIES = Object.keys(TASK_PRIORITY) as TaskPriority[];
const KINDS = Object.keys(TASK_KIND) as TaskKind[];

export async function createTask(formData: FormData) {
  await requireUser();

  const title = str(formData, "title");
  if (!title) return;

  const clientId = optStr(formData, "clientId");

  await db.task.create({
    data: {
      title,
      description: optStr(formData, "description"),
      status: enumOf(formData, "status", STATUSES, "TODO"),
      priority: enumOf(formData, "priority", PRIORITIES, "MEDIUM"),
      kind: enumOf(formData, "kind", KINDS, "TASK"),
      dueAt: date(formData, "dueAt"),
      clientId,
      projectId: optStr(formData, "projectId"),
    },
  });

  revalidatePath("/taches");
  if (clientId) revalidatePath(`/clients/${clientId}/suivi`);
}

export async function updateTask(taskId: string, formData: FormData) {
  await requireUser();

  const status = enumOf(formData, "status", STATUSES, "TODO");

  await db.task.update({
    where: { id: taskId },
    data: {
      title: str(formData, "title"),
      description: optStr(formData, "description"),
      status,
      priority: enumOf(formData, "priority", PRIORITIES, "MEDIUM"),
      kind: enumOf(formData, "kind", KINDS, "TASK"),
      dueAt: date(formData, "dueAt"),
      clientId: optStr(formData, "clientId"),
      completedAt: status === "DONE" ? new Date() : null,
    },
  });

  revalidatePath("/taches");
}

/** Coche / décoche une tâche. Action la plus utilisée : elle reste en un clic. */
export async function toggleTask(taskId: string) {
  await requireUser();

  const task = await db.task.findUnique({
    where: { id: taskId },
    select: { status: true, clientId: true },
  });
  if (!task) return;

  const done = task.status === "DONE";

  await db.task.update({
    where: { id: taskId },
    data: {
      status: done ? "TODO" : "DONE",
      completedAt: done ? null : new Date(),
    },
  });

  revalidatePath("/taches");
  if (task.clientId) revalidatePath(`/clients/${task.clientId}/suivi`);
}

export async function deleteTask(taskId: string) {
  await requireUser();
  const task = await db.task.delete({ where: { id: taskId } });
  revalidatePath("/taches");
  if (task.clientId) revalidatePath(`/clients/${task.clientId}/suivi`);
}
