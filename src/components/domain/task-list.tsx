import type { Client, Task } from "@prisma/client";
import { Check, Trash2 } from "lucide-react";
import Link from "next/link";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { EmptyState } from "@/components/ui/empty-state";
import { TASK_KIND, TASK_PRIORITY, TASK_STATUS } from "@/lib/constants";
import { deleteTask, toggleTask } from "@/lib/actions/tasks";
import { formatDate } from "@/lib/format";
import { countdownLabel, daysUntil } from "@/lib/renewals";
import { cn } from "@/lib/utils";

export type TaskWithClient = Task & { client: Pick<Client, "id" | "company"> | null };

/**
 * Liste de tâches. La case à cocher est un vrai formulaire : cocher fonctionne
 * même avant l'hydratation, et il n'y a aucun état client à synchroniser.
 */
export function TaskList({
  tasks,
  emptyTitle = "Aucune tâche",
  emptyDescription,
  showClient = true,
}: {
  tasks: TaskWithClient[];
  emptyTitle?: string;
  emptyDescription?: string;
  showClient?: boolean;
}) {
  if (tasks.length === 0) {
    return <EmptyState compact title={emptyTitle} description={emptyDescription} />;
  }

  return (
    <div className="divide-y divide-line">
      {tasks.map((task) => {
        const toggle = toggleTask.bind(null, task.id);
        const remove = deleteTask.bind(null, task.id);
        const done = task.status === "DONE";
        const days = task.dueAt ? daysUntil(task.dueAt) : null;
        const late = days !== null && days < 0 && !done;

        return (
          <div
            key={task.id}
            className="group flex items-start gap-3 px-4 py-2.5 transition-colors hover:bg-surface-inset"
          >
            <form action={toggle} className="pt-0.5">
              <button
                type="submit"
                aria-label={done ? "Rouvrir la tâche" : "Marquer comme terminée"}
                className={cn(
                  "flex h-4 w-4 items-center justify-center rounded-[5px] border transition-all duration-150",
                  done
                    ? "border-[var(--accent)] bg-[var(--accent)] text-white"
                    : "border-line-strong hover:border-[var(--accent)]",
                )}
              >
                {done ? <Check className="h-3 w-3" strokeWidth={3} /> : null}
              </button>
            </form>

            <div className="min-w-0 flex-1">
              <div className="flex flex-wrap items-center gap-2">
                <span
                  className={cn(
                    "text-[13px]",
                    done ? "text-ink-muted line-through" : "font-medium text-ink",
                  )}
                >
                  {task.title}
                </span>
                {!done && task.priority !== "MEDIUM" ? (
                  <Badge tone={TASK_PRIORITY[task.priority].tone} square dot={false}>
                    {TASK_PRIORITY[task.priority].label}
                  </Badge>
                ) : null}
                {task.status === "BLOCKED" ? (
                  <Badge tone={TASK_STATUS.BLOCKED.tone} square dot={false}>
                    Bloqué
                  </Badge>
                ) : null}
              </div>

              {task.description ? (
                <p className="mt-0.5 line-clamp-2 text-xs text-ink-muted">{task.description}</p>
              ) : null}

              <div className="mt-1 flex flex-wrap items-center gap-x-3 gap-y-0.5 text-[11px] text-ink-muted">
                <span>{TASK_KIND[task.kind].label}</span>
                {showClient && task.client ? (
                  <Link
                    href={`/clients/${task.client.id}`}
                    className="hover:text-accent-text"
                  >
                    {task.client.company}
                  </Link>
                ) : null}
                {task.dueAt ? (
                  <span
                    data-tone={late ? "danger" : days !== null && days <= 3 ? "warning" : "muted"}
                    className="text-[color:var(--tone-fg)]"
                    title={formatDate(task.dueAt, "long")}
                  >
                    {done ? formatDate(task.dueAt, "short") : countdownLabel(days ?? 0)}
                  </span>
                ) : null}
              </div>
            </div>

            <form action={remove} className="opacity-0 transition-opacity group-hover:opacity-100">
              <Button variant="ghost" size="icon" type="submit" aria-label="Supprimer la tâche">
                <Trash2 className="h-3.5 w-3.5" />
              </Button>
            </form>
          </div>
        );
      })}
    </div>
  );
}
