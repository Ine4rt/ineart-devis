"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";

import { cn } from "@/lib/utils";

export interface TabItem {
  href: string;
  label: string;
  count?: number;
  /** Correspondance exacte du chemin (onglet « Vue d'ensemble »). */
  exact?: boolean;
}

/**
 * Onglets basés sur l'URL plutôt que sur un état local : chaque onglet est une
 * vraie page, donc partageable, rechargeable et navigable au clavier.
 */
export function Tabs({ items, className }: { items: TabItem[]; className?: string }) {
  const pathname = usePathname();

  return (
    <nav
      className={cn(
        "-mb-px flex items-center gap-0.5 overflow-x-auto border-b border-line",
        className,
      )}
    >
      {items.map((item) => {
        const active = item.exact ? pathname === item.href : pathname.startsWith(item.href);
        return (
          <Link
            key={item.href}
            href={item.href}
            data-active={active}
            className={cn(
              "relative flex shrink-0 items-center gap-1.5 border-b-2 border-transparent px-3 py-2 text-[13px] font-medium text-ink-muted transition-colors",
              "hover:text-ink",
              active && "border-[var(--accent)] text-ink",
            )}
          >
            {item.label}
            {item.count !== undefined && item.count > 0 ? (
              <span
                className={cn(
                  "rounded-full px-1.5 py-px text-[10px] font-semibold tabular-nums",
                  active
                    ? "bg-accent-soft text-accent-text"
                    : "bg-surface-inset text-ink-muted",
                )}
              >
                {item.count}
              </span>
            ) : null}
          </Link>
        );
      })}
    </nav>
  );
}

/** Contrôle segmenté pour basculer entre deux ou trois vues d'un même écran. */
export function SegmentedLinks({
  items,
  className,
}: {
  items: { href: string; label: string; active: boolean }[];
  className?: string;
}) {
  return (
    <div className={cn("inline-flex rounded-md border border-line bg-surface-inset p-0.5", className)}>
      {items.map((item) => (
        <Link
          key={item.href}
          href={item.href}
          className={cn(
            "rounded-[5px] px-2.5 py-1 text-xs font-medium transition-all duration-150",
            item.active
              ? "bg-surface text-ink shadow-[var(--shadow-xs)]"
              : "text-ink-muted hover:text-ink",
          )}
        >
          {item.label}
        </Link>
      ))}
    </div>
  );
}
