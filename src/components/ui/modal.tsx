"use client";

import { X } from "lucide-react";
import { useEffect, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";

import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

/**
 * Fenêtre modale et panneau latéral.
 *
 * Écrit à la main plutôt qu'importé d'une bibliothèque : on maîtrise ainsi le
 * timing des animations, la restitution du focus et le verrouillage du scroll,
 * qui sont exactement les détails qui séparent un outil agréable d'un outil
 * approximatif.
 */

export function Modal({
  open,
  onClose,
  title,
  description,
  children,
  footer,
  size = "md",
  variant = "dialog",
}: {
  open: boolean;
  onClose: () => void;
  title: ReactNode;
  description?: ReactNode;
  children: ReactNode;
  footer?: ReactNode;
  size?: "sm" | "md" | "lg" | "xl";
  /** `dialog` = centré ; `panel` = tiroir latéral pour les formulaires longs. */
  variant?: "dialog" | "panel";
}) {
  const [mounted, setMounted] = useState(false);
  const panelRef = useRef<HTMLDivElement>(null);
  const restoreFocusRef = useRef<HTMLElement | null>(null);

  useEffect(() => setMounted(true), []);

  useEffect(() => {
    if (!open) return;

    restoreFocusRef.current = document.activeElement as HTMLElement | null;

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        onClose();
        return;
      }
      if (event.key !== "Tab" || !panelRef.current) return;

      // Piège à focus : la tabulation ne doit pas sortir de la modale.
      const focusables = panelRef.current.querySelectorAll<HTMLElement>(
        'a[href], button:not([disabled]), textarea, input, select, [tabindex]:not([tabindex="-1"])',
      );
      if (focusables.length === 0) return;
      const first = focusables[0];
      const last = focusables[focusables.length - 1];

      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };

    document.addEventListener("keydown", onKeyDown);
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";

    // Laisse le temps au panneau d'être peint avant de déplacer le focus.
    const focusTimer = window.setTimeout(() => {
      panelRef.current
        ?.querySelector<HTMLElement>("[data-autofocus], input, textarea, select, button")
        ?.focus();
    }, 40);

    return () => {
      document.removeEventListener("keydown", onKeyDown);
      document.body.style.overflow = previousOverflow;
      window.clearTimeout(focusTimer);
      restoreFocusRef.current?.focus?.();
    };
  }, [open, onClose]);

  if (!mounted || !open) return null;

  const widths = { sm: "max-w-md", md: "max-w-xl", lg: "max-w-3xl", xl: "max-w-5xl" } as const;

  return createPortal(
    <div className="fixed inset-0 z-50 flex" role="dialog" aria-modal="true">
      <div
        className="absolute inset-0 animate-fade-in bg-[var(--overlay)] backdrop-blur-[2px]"
        onClick={onClose}
      />

      <div
        ref={panelRef}
        className={cn(
          "relative z-10 flex flex-col overflow-hidden border border-line bg-surface shadow-[var(--shadow-lg)]",
          variant === "panel"
            ? "ml-auto h-full w-full max-w-xl animate-slide-panel border-y-0 border-r-0"
            : cn("m-auto max-h-[88vh] w-[calc(100%-2rem)] animate-scale-in rounded-xl", widths[size]),
        )}
      >
        <div className="flex items-start justify-between gap-4 border-b border-line px-5 py-3.5">
          <div className="min-w-0">
            <h2 className="text-sm font-semibold text-ink">{title}</h2>
            {description ? (
              <p className="mt-0.5 text-xs text-ink-muted">{description}</p>
            ) : null}
          </div>
          <Button variant="ghost" size="icon" onClick={onClose} aria-label="Fermer">
            <X className="h-4 w-4" />
          </Button>
        </div>

        <div className="min-h-0 flex-1 overflow-y-auto">{children}</div>

        {footer ? (
          <div className="flex items-center justify-end gap-2 border-t border-line bg-surface-subtle px-5 py-3">
            {footer}
          </div>
        ) : null}
      </div>
    </div>,
    document.body,
  );
}

/**
 * Modale déclenchée par un bouton, avec état interne.
 * `children` reçoit `close` pour pouvoir fermer après soumission.
 */
export function ModalTrigger({
  trigger,
  children,
  ...modalProps
}: Omit<Parameters<typeof Modal>[0], "open" | "onClose" | "children"> & {
  trigger: (open: () => void) => ReactNode;
  children: (close: () => void) => ReactNode;
}) {
  const [open, setOpen] = useState(false);
  const close = () => setOpen(false);

  return (
    <>
      {trigger(() => setOpen(true))}
      <Modal {...modalProps} open={open} onClose={close}>
        {children(close)}
      </Modal>
    </>
  );
}
