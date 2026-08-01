"use client";

import { Check, Copy } from "lucide-react";
import { useState } from "react";

import { useToast } from "@/components/ui/toast";
import { cn } from "@/lib/utils";

/** Copie une valeur dans le presse-papiers avec retour visuel immédiat. */
export function CopyButton({
  value,
  label = "Copier",
  className,
  silent,
}: {
  value: string;
  label?: string;
  className?: string;
  silent?: boolean;
}) {
  const [copied, setCopied] = useState(false);
  const { notify } = useToast();

  async function copy() {
    try {
      await navigator.clipboard.writeText(value);
      setCopied(true);
      if (!silent) notify({ title: "Copié", description: label, tone: "success" });
      window.setTimeout(() => setCopied(false), 1600);
    } catch {
      notify({
        title: "Copie impossible",
        description: "Le presse-papiers est indisponible dans ce contexte.",
        tone: "danger",
      });
    }
  }

  return (
    <button
      type="button"
      onClick={copy}
      aria-label={label}
      title={label}
      className={cn(
        "inline-flex h-6 w-6 shrink-0 items-center justify-center rounded-[var(--radius-xs)] text-ink-muted transition-colors hover:bg-surface-inset hover:text-ink",
        className,
      )}
    >
      {copied ? (
        <Check className="h-3.5 w-3.5 text-[var(--tone-success)]" />
      ) : (
        <Copy className="h-3.5 w-3.5" />
      )}
    </button>
  );
}
