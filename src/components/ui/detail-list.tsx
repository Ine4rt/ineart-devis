import Link from "next/link";
import type { ReactNode } from "react";

import { CopyButton } from "@/components/ui/copy-button";
import { cn } from "@/lib/utils";

/**
 * Affichage clé / valeur des fiches. Une valeur absente affiche un tiret plutôt
 * que de disparaître : savoir qu'un champ n'est pas rempli est une information.
 */

export function DetailList({ className, ...props }: React.ComponentProps<"dl">) {
  return <dl className={cn("divide-y divide-line", className)} {...props} />;
}

export function DetailRow({
  label,
  children,
  copyValue,
  href,
  mono,
}: {
  label: string;
  children?: ReactNode;
  /** Ajoute un bouton de copie (IP, identifiant, numéro de TVA…). */
  copyValue?: string | null;
  href?: string | null;
  mono?: boolean;
}) {
  const isEmpty =
    children === null || children === undefined || children === "" || children === false;

  const content = isEmpty ? (
    <span className="text-ink-muted">—</span>
  ) : href ? (
    <Link
      href={href}
      target={href.startsWith("http") ? "_blank" : undefined}
      rel={href.startsWith("http") ? "noreferrer" : undefined}
      className="link-subtle"
    >
      {children}
    </Link>
  ) : (
    children
  );

  return (
    <div className="flex items-baseline gap-4 py-2">
      <dt className="w-36 shrink-0 text-xs text-ink-muted">{label}</dt>
      <dd
        className={cn(
          "flex min-w-0 flex-1 items-center gap-1 text-[13px] text-ink",
          mono && "font-mono text-xs",
        )}
      >
        <span className="min-w-0 break-words">{content}</span>
        {copyValue && !isEmpty ? <CopyButton value={copyValue} label={label} silent /> : null}
      </dd>
    </div>
  );
}

/** Groupe de lignes avec un intertitre discret. */
export function DetailGroup({ title, children }: { title: string; children: ReactNode }) {
  return (
    <div className="py-1">
      <p className="pb-1 pt-2 section-label">{title}</p>
      {children}
    </div>
  );
}
