"use client";

import type { PipelineStage } from "@prisma/client";
import { CheckCircle2, GripVertical, XCircle } from "lucide-react";
import Link from "next/link";
import { useOptimistic, useState, useTransition } from "react";

import { Avatar } from "@/components/ui/avatar";
import { useToast } from "@/components/ui/toast";
import { PIPELINE_BOARD_STAGES, PIPELINE_STAGE } from "@/lib/constants";
import { moveProspect } from "@/lib/actions/clients";
import { RelativeTime } from "@/components/ui/relative-time";
import { formatMoney } from "@/lib/format";
import { cn } from "@/lib/utils";

/**
 * Kanban de prospection.
 *
 * Glisser-déposer natif (HTML5) plutôt qu'une bibliothèque : le besoin se
 * limite à déplacer une carte d'une colonne à l'autre. `useOptimistic` déplace
 * la carte immédiatement, avant même la réponse du serveur — c'est ce qui rend
 * le tableau agréable à manipuler.
 */

export interface ProspectCard {
  id: string;
  company: string;
  contact: string | null;
  city: string | null;
  logoUrl: string | null;
  potentialValue: number;
  stage: PipelineStage;
  updatedAt: string;
}

export function PipelineBoard({ prospects }: { prospects: ProspectCard[] }) {
  const [, startTransition] = useTransition();
  const { notify } = useToast();
  const [dragOver, setDragOver] = useState<PipelineStage | null>(null);

  const [cards, moveCard] = useOptimistic(
    prospects,
    (current: ProspectCard[], move: { id: string; stage: PipelineStage }) =>
      current.map((card) => (card.id === move.id ? { ...card, stage: move.stage } : card)),
  );

  function handleDrop(stage: PipelineStage, id: string) {
    setDragOver(null);
    const card = cards.find((candidate) => candidate.id === id);
    if (!card || card.stage === stage) return;

    startTransition(async () => {
      moveCard({ id, stage });
      await moveProspect(id, stage);
      if (stage === "WON") {
        notify({
          title: `${card.company} converti en client`,
          description: "Le statut est passé à « Actif ». Créez son projet et son abonnement.",
          tone: "success",
        });
      }
    });
  }

  const columns: { stage: PipelineStage; label: string; tone: string }[] = [
    ...PIPELINE_BOARD_STAGES.map((stage) => ({
      stage,
      label: PIPELINE_STAGE[stage].label,
      tone: PIPELINE_STAGE[stage].tone,
    })),
  ];

  return (
    <div className="space-y-3">
      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-5">
        {columns.map((column) => {
          const columnCards = cards.filter((card) => card.stage === column.stage);
          const value = columnCards.reduce((sum, card) => sum + card.potentialValue, 0);

          return (
            <section
              key={column.stage}
              data-tone={column.tone}
              onDragOver={(event) => {
                event.preventDefault();
                setDragOver(column.stage);
              }}
              onDragLeave={() => setDragOver(null)}
              onDrop={(event) => {
                event.preventDefault();
                handleDrop(column.stage, event.dataTransfer.getData("text/plain"));
              }}
              className={cn(
                "flex min-h-[180px] flex-col rounded-lg border bg-canvas-subtle transition-colors duration-150",
                dragOver === column.stage
                  ? "border-[color:var(--tone-fg)] bg-[color:var(--tone-soft)]"
                  : "border-line",
              )}
            >
              <header className="flex items-center gap-2 border-b border-line px-3 py-2">
                <span className="dot" />
                <h2 className="min-w-0 flex-1 truncate text-xs font-semibold text-ink">
                  {column.label}
                </h2>
                <span className="text-[11px] text-ink-muted tabular-nums">
                  {columnCards.length}
                </span>
              </header>

              {value > 0 ? (
                <div className="border-b border-line px-3 py-1.5 text-[11px] text-ink-muted tabular-nums">
                  {formatMoney(value)} de potentiel
                </div>
              ) : null}

              <div className="flex-1 space-y-2 p-2">
                {columnCards.map((card) => (
                  <article
                    key={card.id}
                    draggable
                    onDragStart={(event) => event.dataTransfer.setData("text/plain", card.id)}
                    className="group cursor-grab rounded-md border border-line bg-surface p-2.5 shadow-[var(--shadow-xs)] transition-all duration-150 hover:border-line-strong hover:shadow-[var(--shadow-sm)] active:cursor-grabbing"
                  >
                    <div className="flex items-start gap-2">
                      <Avatar name={card.company} src={card.logoUrl} size="xs" />
                      <div className="min-w-0 flex-1">
                        <Link
                          href={`/clients/${card.id}`}
                          className="block truncate text-xs font-medium text-ink hover:text-accent-text"
                        >
                          {card.company}
                        </Link>
                        {card.contact || card.city ? (
                          <p className="truncate text-[11px] text-ink-muted">
                            {[card.contact, card.city].filter(Boolean).join(" · ")}
                          </p>
                        ) : null}
                      </div>
                      <GripVertical className="h-3.5 w-3.5 shrink-0 text-ink-muted opacity-0 transition-opacity group-hover:opacity-100" />
                    </div>

                    <div className="mt-2 flex items-center justify-between gap-2">
                      <RelativeTime value={card.updatedAt} className="text-[11px] text-ink-muted" />
                      {card.potentialValue > 0 ? (
                        <span className="text-[11px] font-medium text-ink tabular-nums">
                          {formatMoney(card.potentialValue, { compact: true })}
                        </span>
                      ) : null}
                    </div>
                  </article>
                ))}

                {columnCards.length === 0 ? (
                  <p className="px-1 py-4 text-center text-[11px] text-ink-muted">
                    Déposez une fiche ici
                  </p>
                ) : null}
              </div>
            </section>
          );
        })}
      </div>

      {/* Zones de sortie du pipeline, séparées du flux principal. */}
      <div className="grid gap-3 sm:grid-cols-2">
        {(["WON", "LOST"] as PipelineStage[]).map((stage) => (
          <div
            key={stage}
            data-tone={PIPELINE_STAGE[stage].tone}
            onDragOver={(event) => {
              event.preventDefault();
              setDragOver(stage);
            }}
            onDragLeave={() => setDragOver(null)}
            onDrop={(event) => {
              event.preventDefault();
              handleDrop(stage, event.dataTransfer.getData("text/plain"));
            }}
            className={cn(
              "flex items-center justify-center gap-2 rounded-lg border border-dashed px-4 py-4 text-xs font-medium transition-colors duration-150",
              dragOver === stage
                ? "border-[color:var(--tone-fg)] bg-[color:var(--tone-soft)] text-[color:var(--tone-fg)]"
                : "border-line text-ink-muted",
            )}
          >
            {stage === "WON" ? (
              <CheckCircle2 className="h-4 w-4" />
            ) : (
              <XCircle className="h-4 w-4" />
            )}
            {stage === "WON"
              ? "Déposer ici pour convertir en client actif"
              : "Déposer ici pour marquer comme perdu"}
          </div>
        ))}
      </div>
    </div>
  );
}
