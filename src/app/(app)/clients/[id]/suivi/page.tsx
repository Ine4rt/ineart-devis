import { History, Pin, PinOff, Plus, StickyNote, Trash2 } from "lucide-react";

import { TaskList } from "@/components/domain/task-list";
import { Button } from "@/components/ui/button";
import { Card, CardBody, CardHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { EnumSelect, Field, Input, Textarea } from "@/components/ui/form";
import { SubmitButton } from "@/components/ui/submit-button";
import { TASK_KIND_LIST, TASK_PRIORITY_LIST } from "@/lib/constants";
import { createNote, deleteNote, toggleNotePin } from "@/lib/actions/clients";
import { createTask } from "@/lib/actions/tasks";
import { db } from "@/lib/db";
import { formatDateTime, formatRelative } from "@/lib/format";
import { cn } from "@/lib/utils";

export const dynamic = "force-dynamic";

/**
 * Onglet suivi : la mémoire de la relation client — notes, tâches et
 * journal d'activité réunis chronologiquement.
 */
export default async function ClientFollowUpPage({
  params,
}: {
  params: Promise<{ id: string }>;
}) {
  const { id } = await params;

  const [notes, tasks, activities] = await Promise.all([
    db.note.findMany({
      where: { clientId: id },
      orderBy: [{ pinned: "desc" }, { updatedAt: "desc" }],
    }),
    db.task.findMany({
      where: { clientId: id },
      orderBy: [{ status: "asc" }, { dueAt: "asc" }],
      include: { client: { select: { id: true, company: true } } },
    }),
    db.activity.findMany({
      where: { clientId: id },
      orderBy: { createdAt: "desc" },
      take: 40,
    }),
  ]);

  return (
    <div className="grid gap-4 lg:grid-cols-2">
      <div className="space-y-4">
        <Card>
          <CardHeader
            title="Notes"
            description="Comptes rendus d'appel, décisions, points d'attention."
            icon={<StickyNote className="h-4 w-4" />}
          />

          <CardBody className="border-b border-line">
            <form action={createNote} className="space-y-2">
              <input type="hidden" name="clientId" value={id} />
              <Field htmlFor="note-title">
                <Input id="note-title" name="title" placeholder="Titre (facultatif)" />
              </Field>
              <Field htmlFor="note-body">
                <Textarea
                  id="note-body"
                  name="body"
                  rows={3}
                  required
                  placeholder="Appel du 12/03 : souhaite ajouter une page « Nos réalisations » avant l'été."
                />
              </Field>
              <div className="flex justify-end">
                <SubmitButton variant="default">
                  <Plus className="h-3.5 w-3.5" />
                  Ajouter la note
                </SubmitButton>
              </div>
            </form>
          </CardBody>

          <div className="divide-y divide-line">
            {notes.length > 0 ? (
              notes.map((note) => {
                const pin = toggleNotePin.bind(null, note.id, id);
                const remove = deleteNote.bind(null, note.id, id);
                return (
                  <div
                    key={note.id}
                    className={cn("group px-4 py-3", note.pinned && "bg-surface-subtle")}
                  >
                    <div className="flex items-start gap-2">
                      <div className="min-w-0 flex-1">
                        {note.title ? (
                          <p className="text-[13px] font-semibold text-ink">{note.title}</p>
                        ) : null}
                        <p className="mt-0.5 whitespace-pre-wrap text-[13px] leading-relaxed text-ink-secondary">
                          {note.body}
                        </p>
                        <p className="mt-1 text-[11px] text-ink-muted">
                          {formatDateTime(note.createdAt)}
                        </p>
                      </div>

                      <div className="flex shrink-0 items-center gap-1">
                        <form action={pin}>
                          <Button
                            variant="ghost"
                            size="icon"
                            type="submit"
                            aria-label={note.pinned ? "Désépingler" : "Épingler"}
                            title={note.pinned ? "Désépingler" : "Épingler"}
                          >
                            {note.pinned ? (
                              <PinOff className="h-3.5 w-3.5" />
                            ) : (
                              <Pin className="h-3.5 w-3.5" />
                            )}
                          </Button>
                        </form>
                        <form
                          action={remove}
                          className="opacity-0 transition-opacity group-hover:opacity-100"
                        >
                          <Button variant="ghost" size="icon" type="submit" aria-label="Supprimer">
                            <Trash2 className="h-3.5 w-3.5" />
                          </Button>
                        </form>
                      </div>
                    </div>
                  </div>
                );
              })
            ) : (
              <EmptyState compact title="Aucune note" description="Consignez ici ce qui compte." />
            )}
          </div>
        </Card>
      </div>

      <div className="space-y-4">
        <Card>
          <CardHeader title="Tâches & rappels" />

          <CardBody className="border-b border-line">
            <form action={createTask} className="flex flex-wrap items-end gap-2">
              <input type="hidden" name="clientId" value={id} />
              <Field label="Intitulé" htmlFor="task-title" className="min-w-[160px] flex-1">
                <Input id="task-title" name="title" required placeholder="Relancer avant échéance" />
              </Field>
              <Field label="Type" htmlFor="task-kind" className="w-28">
                <EnumSelect
                  id="task-kind"
                  name="kind"
                  options={TASK_KIND_LIST}
                  defaultValue="FOLLOW_UP"
                />
              </Field>
              <Field label="Priorité" htmlFor="task-priority" className="w-28">
                <EnumSelect
                  id="task-priority"
                  name="priority"
                  options={TASK_PRIORITY_LIST}
                  defaultValue="MEDIUM"
                />
              </Field>
              <Field label="Échéance" htmlFor="task-due">
                <Input id="task-due" name="dueAt" type="date" />
              </Field>
              <SubmitButton variant="default">
                <Plus className="h-3.5 w-3.5" />
                Ajouter
              </SubmitButton>
            </form>
          </CardBody>

          <TaskList
            tasks={tasks}
            showClient={false}
            emptyTitle="Aucune tâche"
            emptyDescription="Ajoutez un rappel pour ce client."
          />
        </Card>

        <Card>
          <CardHeader
            title="Journal d'activité"
            description="Trace automatique des actions effectuées."
            icon={<History className="h-4 w-4" />}
          />
          <div className="divide-y divide-line">
            {activities.length > 0 ? (
              activities.map((activity) => (
                <div key={activity.id} className="px-4 py-2">
                  <p className="text-xs text-ink">{activity.summary.replace(/\s*:?\s*$/, "")}</p>
                  <p
                    className="mt-0.5 text-[11px] text-ink-muted"
                    title={formatDateTime(activity.createdAt)}
                  >
                    {formatRelative(activity.createdAt)}
                  </p>
                </div>
              ))
            ) : (
              <EmptyState compact title="Aucune activité enregistrée" />
            )}
          </div>
        </Card>
      </div>
    </div>
  );
}
