"use client";

import { AlertTriangle, CheckCircle2, Info, X } from "lucide-react";
import {
  createContext,
  useCallback,
  useContext,
  useMemo,
  useState,
  type ReactNode,
} from "react";

import type { Tone } from "@/lib/constants";

/**
 * Notifications éphémères. Utilisées pour confirmer les actions qui ne
 * provoquent pas de changement d'écran visible (copie d'un identifiant,
 * marquage d'un paiement, suppression…).
 */

interface Toast {
  id: number;
  title: string;
  description?: string;
  tone: Tone;
}

const ToastContext = createContext<{
  notify: (toast: Omit<Toast, "id">) => void;
} | null>(null);

export function useToast() {
  const context = useContext(ToastContext);
  if (!context) throw new Error("useToast doit être utilisé à l'intérieur de <ToastProvider>");
  return context;
}

const ICONS: Partial<Record<Tone, ReactNode>> = {
  success: <CheckCircle2 className="h-4 w-4" />,
  danger: <AlertTriangle className="h-4 w-4" />,
  warning: <AlertTriangle className="h-4 w-4" />,
};

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([]);

  const dismiss = useCallback((id: number) => {
    setToasts((current) => current.filter((toast) => toast.id !== id));
  }, []);

  const notify = useCallback(
    (toast: Omit<Toast, "id">) => {
      const id = Date.now() + Math.random();
      setToasts((current) => [...current, { ...toast, id }]);
      window.setTimeout(() => dismiss(id), 4500);
    },
    [dismiss],
  );

  const value = useMemo(() => ({ notify }), [notify]);

  return (
    <ToastContext.Provider value={value}>
      {children}
      <div className="pointer-events-none fixed bottom-4 right-4 z-[60] flex w-[min(22rem,calc(100vw-2rem))] flex-col gap-2">
        {toasts.map((toast) => (
          <div
            key={toast.id}
            data-tone={toast.tone}
            role="status"
            className="pointer-events-auto flex animate-rise-in items-start gap-3 rounded-lg border border-line bg-surface-raised p-3 shadow-[var(--shadow-lg)]"
          >
            <span className="mt-px text-[color:var(--tone-fg)]">
              {ICONS[toast.tone] ?? <Info className="h-4 w-4" />}
            </span>
            <div className="min-w-0 flex-1">
              <p className="text-xs font-semibold text-ink">{toast.title}</p>
              {toast.description ? (
                <p className="mt-0.5 text-xs text-ink-muted">{toast.description}</p>
              ) : null}
            </div>
            <button
              type="button"
              onClick={() => dismiss(toast.id)}
              className="text-ink-muted transition-colors hover:text-ink"
              aria-label="Fermer la notification"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
}
