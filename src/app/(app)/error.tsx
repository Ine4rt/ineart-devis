"use client";

import { RotateCcw, TriangleAlert } from "lucide-react";
import Link from "next/link";
import { useEffect } from "react";

import { Button } from "@/components/ui/button";

export default function AppError({
  error,
  reset,
}: {
  error: Error & { digest?: string };
  reset: () => void;
}) {
  useEffect(() => {
    console.error(error);
  }, [error]);

  return (
    <div className="flex min-h-[60vh] items-center justify-center">
      <div className="card max-w-md animate-rise-in p-6 text-center">
        <div
          data-tone="danger"
          className="mx-auto flex h-10 w-10 items-center justify-center rounded-lg border border-[color:var(--tone-line)] bg-[color:var(--tone-soft)] text-[color:var(--tone-fg)]"
        >
          <TriangleAlert className="h-4 w-4" />
        </div>

        <h1 className="mt-3 text-sm font-semibold text-ink">Une erreur est survenue</h1>
        <p className="mt-1.5 text-xs leading-relaxed text-ink-muted">
          {error.message || "L'écran n'a pas pu être chargé."}
          {error.digest ? (
            <span className="mt-1 block font-mono text-[11px] opacity-70">#{error.digest}</span>
          ) : null}
        </p>

        <div className="mt-4 flex items-center justify-center gap-2">
          <Button onClick={reset} variant="primary">
            <RotateCcw className="h-3.5 w-3.5" />
            Réessayer
          </Button>
          <Link href="/" className="btn btn-ghost h-8 px-3">
            Tableau de bord
          </Link>
        </div>
      </div>
    </div>
  );
}
