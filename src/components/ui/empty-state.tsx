import type { ReactNode } from "react";

import { cn } from "@/lib/utils";

/**
 * État vide. Toujours accompagné d'une action : un écran vide sans porte de
 * sortie est un cul-de-sac.
 */
export function EmptyState({
  icon,
  title,
  description,
  action,
  className,
  compact,
}: {
  icon?: ReactNode;
  title: string;
  description?: string;
  action?: ReactNode;
  className?: string;
  compact?: boolean;
}) {
  return (
    <div
      className={cn(
        "flex flex-col items-center justify-center text-center",
        compact ? "gap-2 px-4 py-8" : "gap-3 px-6 py-14",
        className,
      )}
    >
      {icon ? (
        <div className="flex h-10 w-10 items-center justify-center rounded-lg border border-line bg-surface-inset text-ink-muted">
          {icon}
        </div>
      ) : null}
      <div className="space-y-1">
        <p className="text-sm font-medium text-ink">{title}</p>
        {description ? (
          <p className="mx-auto max-w-sm text-xs leading-relaxed text-ink-muted">{description}</p>
        ) : null}
      </div>
      {action ? <div className="mt-1">{action}</div> : null}
    </div>
  );
}
