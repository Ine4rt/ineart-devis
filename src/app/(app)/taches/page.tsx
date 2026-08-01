import type { Prisma, TaskStatus } from "@prisma/client";
import { ListChecks, Plus } from "lucide-react";
import type { Metadata } from "next";

import { ListToolbar } from "@/components/domain/list-toolbar";
import { filterOptions } from "@/lib/filters";
import { TaskList } from "@/components/domain/task-list";
import { Card, CardBody, CardHeader, PageHeader } from "@/components/ui/card";
import { EnumSelect, Field, Input, Select } from "@/components/ui/form";
import { SegmentedLinks } from "@/components/ui/tabs";
import { StatTile } from "@/components/ui/stat-tile";
import { SubmitButton } from "@/components/ui/submit-button";
import {
  TASK_KIND_LIST,
  TASK_PRIORITY_LIST,
  TASK_STATUS,
  TASK_STATUS_LIST,
} from "@/lib/constants";
import { createTask } from "@/lib/actions/tasks";
import { db } from "@/lib/db";
import { getClientOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Tâches" };
export const dynamic = "force-dynamic";

export default async function TasksPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; statut?: string; priorite?: string; filtre?: string }>;
}) {
  const params = await searchParams;
  const query = params.q?.trim();
  const now = new Date();

  const viewFilter: Prisma.TaskWhereInput =
    params.filtre === "retard"
      ? { status: { not: "DONE" }, dueAt: { lt: now } }
      : params.filtre === "terminees"
        ? { status: "DONE" }
        : params.filtre === "toutes"
          ? {}
          : { status: { not: "DONE" } };

  const where: Prisma.TaskWhereInput = {
    ...viewFilter,
    ...(params.statut && params.statut in TASK_STATUS
      ? { status: params.statut as TaskStatus }
      : {}),
    ...(params.priorite ? { priority: params.priorite as never } : {}),
    ...(query
      ? { OR: [{ title: { contains: query } }, { description: { contains: query } }] }
      : {}),
  };

  const [tasks, clients, counts] = await Promise.all([
    db.task.findMany({
      where,
      orderBy: [{ status: "asc" }, { dueAt: "asc" }, { priority: "desc" }],
      include: { client: { select: { id: true, company: true } } },
    }),
    getClientOptions(),
    Promise.all([
      db.task.count({ where: { status: { not: "DONE" } } }),
      db.task.count({ where: { status: { not: "DONE" }, dueAt: { lt: now } } }),
      db.task.count({
        where: {
          status: { not: "DONE" },
          dueAt: {
            gte: now,
            lte: new Date(now.getFullYear(), now.getMonth(), now.getDate() + 7),
          },
        },
      }),
      db.task.count({ where: { status: "DONE" } }),
    ]),
  ]);

  const [openCount, lateCount, weekCount, doneCount] = counts;

  const views = [
    { href: "/taches", label: "À faire", key: undefined },
    { href: "/taches?filtre=retard", label: "En retard", key: "retard" },
    { href: "/taches?filtre=terminees", label: "Terminées", key: "terminees" },
    { href: "/taches?filtre=toutes", label: "Toutes", key: "toutes" },
  ];

  return (
    <div className="space-y-4">
      <PageHeader
        title="Tâches & rappels"
        description="Relances, échéances et notes de travail, rattachées à vos clients."
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile label="Ouvertes" value={openCount} />
        <StatTile
          label="En retard"
          value={lateCount}
          tone={lateCount > 0 ? "danger" : "neutral"}
          href="/taches?filtre=retard"
        />
        <StatTile label="Cette semaine" value={weekCount} tone="warning" />
        <StatTile label="Terminées" value={doneCount} tone="success" />
      </div>

      <Card>
        <CardHeader
          title="Ajouter une tâche"
          description="Un titre suffit. Le reste est facultatif."
          icon={<Plus className="h-4 w-4" />}
        />
        <CardBody>
          <form action={createTask} className="flex flex-wrap items-end gap-2">
            <Field label="Intitulé" htmlFor="title" className="min-w-[220px] flex-1">
              <Input id="title" name="title" required placeholder="Relancer pour le renouvellement" />
            </Field>
            <Field label="Client" htmlFor="clientId" className="min-w-[160px]">
              <Select id="clientId" name="clientId">
                <option value="">Aucun</option>
                {clients.map((client) => (
                  <option key={client.id} value={client.id}>
                    {client.company}
                  </option>
                ))}
              </Select>
            </Field>
            <Field label="Type" htmlFor="kind" className="w-32">
              <EnumSelect id="kind" name="kind" options={TASK_KIND_LIST} defaultValue="TASK" />
            </Field>
            <Field label="Priorité" htmlFor="priority" className="w-32">
              <EnumSelect
                id="priority"
                name="priority"
                options={TASK_PRIORITY_LIST}
                defaultValue="MEDIUM"
              />
            </Field>
            <Field label="Échéance" htmlFor="dueAt">
              <Input id="dueAt" name="dueAt" type="date" />
            </Field>
            <SubmitButton>
              <Plus className="h-3.5 w-3.5" />
              Ajouter
            </SubmitButton>
          </form>
        </CardBody>
      </Card>

      <div className="flex flex-wrap items-center justify-between gap-3">
        <SegmentedLinks
          items={views.map((view) => ({
            href: view.href,
            label: view.label,
            active: params.filtre === view.key,
          }))}
        />
      </div>

      <ListToolbar
        searchPlaceholder="Rechercher une tâche…"
        filters={[
          { name: "statut", label: "Statut", options: filterOptions(TASK_STATUS_LIST) },
          { name: "priorite", label: "Priorité", options: filterOptions(TASK_PRIORITY_LIST) },
        ]}
      />

      <Card className="overflow-hidden">
        <TaskList
          tasks={tasks}
          emptyTitle={params.filtre === "retard" ? "Rien en retard" : "Aucune tâche"}
          emptyDescription={
            params.filtre === "retard"
              ? "Vous êtes à jour sur toutes vos échéances."
              : "Ajoutez une tâche avec le formulaire ci-dessus."
          }
        />
      </Card>

      <p className="flex items-center justify-center gap-1.5 text-[11px] text-ink-muted">
        <ListChecks className="h-3 w-3" />
        Les tâches en retard apparaissent aussi en pastille dans le menu latéral.
      </p>
    </div>
  );
}
