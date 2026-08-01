import type { ComponentProps, ReactNode } from "react";

import { cn } from "@/lib/utils";

/**
 * Enveloppe de tableau. Le défilement horizontal est confiné au conteneur :
 * la page elle-même ne défile jamais latéralement, même avec 12 colonnes.
 */
export function TableWrapper({ className, ...props }: ComponentProps<"div">) {
  return (
    <div
      className={cn("card overflow-hidden", className)}
      {...props}
    />
  );
}

export function TableScroll({ className, ...props }: ComponentProps<"div">) {
  return <div className={cn("overflow-x-auto", className)} {...props} />;
}

export function Table({ className, ...props }: ComponentProps<"table">) {
  return <table className={cn("data-table", className)} {...props} />;
}

/** Cellule numérique : alignée à droite et en chiffres tabulaires. */
export function NumCell({ className, ...props }: ComponentProps<"td">) {
  return <td className={cn("text-right tabular-nums", className)} {...props} />;
}

export function NumHead({ className, ...props }: ComponentProps<"th">) {
  return <th className={cn("text-right", className)} {...props} />;
}

/** Cellule principale d'une ligne : titre + sous-titre sur deux niveaux. */
export function PrimaryCell({
  title,
  subtitle,
  leading,
}: {
  title: ReactNode;
  subtitle?: ReactNode;
  leading?: ReactNode;
}) {
  return (
    <div className="flex items-center gap-2.5">
      {leading}
      <div className="min-w-0">
        <div className="truncate font-medium text-ink">{title}</div>
        {subtitle ? <div className="truncate text-xs text-ink-muted">{subtitle}</div> : null}
      </div>
    </div>
  );
}

/** Nombre de résultats + libellé, affiché au-dessus des listes. */
export function ResultCount({ count, singular, plural }: { count: number; singular: string; plural: string }) {
  return (
    <span className="text-xs text-ink-muted tabular-nums">
      {count} {count > 1 ? plural : singular}
    </span>
  );
}
