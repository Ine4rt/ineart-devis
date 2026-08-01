"use client";

import Link from "next/link";
import { useEffect, useRef, useState, type ReactNode } from "react";

import { cn } from "@/lib/utils";

/**
 * Menu contextuel léger. Se ferme au clic extérieur, à Échap et à la sélection
 * d'un élément. Positionné en absolu par rapport au déclencheur (suffisant ici :
 * les menus sont courts et toujours dans le flux de la page).
 */
export function Dropdown({
  trigger,
  children,
  align = "end",
  className,
}: {
  trigger: ReactNode;
  children: ReactNode;
  align?: "start" | "end";
  className?: string;
}) {
  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;

    const onPointerDown = (event: MouseEvent) => {
      if (!containerRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setOpen(false);
    };

    document.addEventListener("mousedown", onPointerDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("mousedown", onPointerDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [open]);

  return (
    <div ref={containerRef} className="relative">
      <div onClick={() => setOpen((value) => !value)}>{trigger}</div>
      {open ? (
        <div
          onClick={() => setOpen(false)}
          className={cn(
            "absolute top-[calc(100%+4px)] z-40 min-w-[200px] animate-scale-in overflow-hidden rounded-lg border border-line bg-surface-raised p-1 shadow-[var(--shadow-lg)]",
            align === "end" ? "right-0 origin-top-right" : "left-0 origin-top-left",
            className,
          )}
          role="menu"
        >
          {children}
        </div>
      ) : null}
    </div>
  );
}

export function DropdownItem({
  children,
  icon,
  danger,
  className,
  ...props
}: React.ComponentProps<"button"> & { icon?: ReactNode; danger?: boolean }) {
  return (
    <button
      type="button"
      role="menuitem"
      className={cn(
        "flex w-full items-center gap-2.5 rounded-md px-2.5 py-1.5 text-left text-xs font-medium transition-colors",
        danger
          ? "text-[var(--tone-danger)] hover:bg-[var(--tone-danger-soft)]"
          : "text-ink-secondary hover:bg-surface-inset hover:text-ink",
        className,
      )}
      {...props}
    >
      {icon ? <span className="shrink-0 opacity-70">{icon}</span> : null}
      <span className="min-w-0 flex-1 truncate">{children}</span>
    </button>
  );
}

/** Même apparence que DropdownItem, mais c'est un vrai lien (pas de <a><button>). */
export function DropdownLink({
  children,
  icon,
  href,
  className,
}: {
  children: ReactNode;
  icon?: ReactNode;
  href: string;
  className?: string;
}) {
  return (
    <Link
      href={href}
      role="menuitem"
      className={cn(
        "flex w-full items-center gap-2.5 rounded-md px-2.5 py-1.5 text-left text-xs font-medium text-ink-secondary transition-colors hover:bg-surface-inset hover:text-ink",
        className,
      )}
    >
      {icon ? <span className="shrink-0 opacity-70">{icon}</span> : null}
      <span className="min-w-0 flex-1 truncate">{children}</span>
    </Link>
  );
}

export function DropdownSeparator() {
  return <div className="my-1 h-px bg-line" />;
}

export function DropdownLabel({ children }: { children: ReactNode }) {
  return <div className="px-2.5 py-1.5 section-label">{children}</div>;
}
