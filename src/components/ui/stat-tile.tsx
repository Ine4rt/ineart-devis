import Link from "next/link";
import type { ReactNode } from "react";

import type { Tone } from "@/lib/constants";
import { cn } from "@/lib/utils";

/**
 * Tuile de statistique. Un chiffre est une information à lire d'un coup d'œil :
 * la valeur domine, le libellé la nomme, et un éventuel indice la contextualise.
 * Pas de graphique décoratif à l'intérieur — juste ce qui se lit.
 */
export function StatTile({
  label,
  value,
  hint,
  icon,
  tone = "neutral",
  href,
  trend,
  footer,
  className,
}: {
  label: string;
  value: ReactNode;
  hint?: ReactNode;
  icon?: ReactNode;
  tone?: Tone;
  href?: string;
  trend?: { value: number; label: string };
  footer?: ReactNode;
  className?: string;
}) {
  const content = (
    <>
      <div className="flex items-start justify-between gap-3">
        <span className="text-xs font-medium text-ink-secondary">{label}</span>
        {icon ? (
          <span data-tone={tone} className="text-[color:var(--tone-fg)] opacity-80">
            {icon}
          </span>
        ) : null}
      </div>

      <div className="mt-2 flex items-baseline gap-2">
        <span className="text-2xl font-semibold tracking-tight text-ink tabular-nums">{value}</span>
        {trend ? (
          <span
            data-tone={trend.value >= 0 ? "success" : "danger"}
            className="text-xs font-medium text-[color:var(--tone-fg)] tabular-nums"
          >
            {trend.value >= 0 ? "+" : ""}
            {trend.value}
            {trend.label}
          </span>
        ) : null}
      </div>

      {hint ? <p className="mt-1 text-xs text-ink-muted">{hint}</p> : null}
      {footer ? <div className="mt-3">{footer}</div> : null}
    </>
  );

  const classes = cn(
    "card p-4 transition-all duration-200",
    href && "hover:-translate-y-px hover:border-line-strong hover:shadow-[var(--shadow-sm)]",
    className,
  );

  if (href) {
    return (
      <Link href={href} className={cn(classes, "block")}>
        {content}
      </Link>
    );
  }

  return <div className={classes}>{content}</div>;
}
