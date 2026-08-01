import type { ReactNode } from "react";

import type { EnumMeta, Tone } from "@/lib/constants";
import { cn } from "@/lib/utils";

interface BadgeProps {
  children: ReactNode;
  tone?: Tone;
  /** Affiche une pastille de couleur avant le texte (listes denses). */
  dot?: boolean;
  square?: boolean;
  className?: string;
  title?: string;
}

export function Badge({ children, tone = "neutral", dot, square, className, title }: BadgeProps) {
  return (
    <span
      data-tone={tone}
      title={title}
      className={cn("badge", square && "badge-square", className)}
    >
      {dot ? <span className="dot" /> : null}
      {children}
    </span>
  );
}

/** Raccourci : rend un badge directement depuis une table d'enum (constants.ts). */
export function EnumBadge({
  meta,
  dot = true,
  className,
}: {
  meta: EnumMeta | undefined;
  dot?: boolean;
  className?: string;
}) {
  if (!meta) return <span className="text-ink-muted">—</span>;
  return (
    <Badge tone={meta.tone} dot={dot} title={meta.description} className={className}>
      {meta.label}
    </Badge>
  );
}

/** Pastille seule, pour les tableaux très denses où le texte prendrait trop de place. */
export function ToneDot({ tone = "neutral", title }: { tone?: Tone; title?: string }) {
  return <span data-tone={tone} className="dot" title={title} />;
}
