import type { ComponentProps, ReactNode } from "react";

import { cn } from "@/lib/utils";

/**
 * Briques de formulaire. Volontairement non contrôlées : les écrans envoient
 * un `FormData` à une server action. Moins d'état client = moins de bugs et
 * des formulaires qui fonctionnent même pendant l'hydratation.
 */

export function Field({
  label,
  htmlFor,
  hint,
  error,
  required,
  children,
  className,
  span,
}: {
  label?: ReactNode;
  htmlFor?: string;
  hint?: ReactNode;
  error?: string | null;
  required?: boolean;
  children: ReactNode;
  className?: string;
  /** Largeur en colonnes dans une FormGrid. */
  span?: 1 | 2 | 3 | 4;
}) {
  return (
    <div
      className={cn("min-w-0 space-y-1.5", className)}
      data-span={span}
      style={span ? { gridColumn: `span ${span} / span ${span}` } : undefined}
    >
      {label ? (
        <label
          htmlFor={htmlFor}
          className="flex items-center gap-1 text-xs font-medium text-ink-secondary"
        >
          {label}
          {required ? <span className="text-[var(--tone-danger)]">*</span> : null}
        </label>
      ) : null}
      {children}
      {error ? (
        <p className="text-xs text-[var(--tone-danger)]">{error}</p>
      ) : hint ? (
        <p className="text-xs text-ink-muted">{hint}</p>
      ) : null}
    </div>
  );
}

export function Input({ className, ...props }: ComponentProps<"input">) {
  return <input className={cn("field", className)} {...props} />;
}

export function Textarea({ className, rows = 4, ...props }: ComponentProps<"textarea">) {
  return <textarea rows={rows} className={cn("field resize-y leading-relaxed", className)} {...props} />;
}

export function Select({ className, ...props }: ComponentProps<"select">) {
  return <select className={cn("field", className)} {...props} />;
}

/**
 * Select alimenté par une table d'enum de constants.ts — évite de recopier
 * les `<option>` sur chaque écran.
 */
export function EnumSelect<T extends string>({
  options,
  placeholder,
  ...props
}: ComponentProps<"select"> & {
  options: [T, { label: string }][];
  placeholder?: string;
}) {
  return (
    <Select {...props}>
      {placeholder ? <option value="">{placeholder}</option> : null}
      {options.map(([value, meta]) => (
        <option key={value} value={value}>
          {meta.label}
        </option>
      ))}
    </Select>
  );
}

export function Checkbox({
  label,
  description,
  className,
  ...props
}: ComponentProps<"input"> & { label: ReactNode; description?: string }) {
  return (
    <label
      className={cn(
        "flex cursor-pointer items-start gap-2.5 rounded-md border border-line bg-surface px-3 py-2.5 transition-colors hover:border-line-strong",
        className,
      )}
    >
      <input
        type="checkbox"
        className="mt-0.5 h-3.5 w-3.5 shrink-0 accent-[var(--accent)]"
        {...props}
      />
      <span className="min-w-0">
        <span className="block text-xs font-medium text-ink">{label}</span>
        {description ? (
          <span className="mt-0.5 block text-xs text-ink-muted">{description}</span>
        ) : null}
      </span>
    </label>
  );
}

/** Grille de formulaire responsive (1 colonne en mobile, N en bureau). */
export function FormGrid({
  columns = 2,
  className,
  ...props
}: ComponentProps<"div"> & { columns?: 1 | 2 | 3 | 4 }) {
  const map = {
    1: "sm:grid-cols-1",
    2: "sm:grid-cols-2",
    3: "sm:grid-cols-2 lg:grid-cols-3",
    4: "sm:grid-cols-2 lg:grid-cols-4",
  } as const;
  return <div className={cn("grid grid-cols-1 gap-4", map[columns], className)} {...props} />;
}

/** Bloc thématique d'un formulaire long (Identité, Adresse, Facturation…). */
export function FormSection({
  title,
  description,
  children,
  icon,
}: {
  title: string;
  description?: string;
  children: ReactNode;
  icon?: ReactNode;
}) {
  return (
    <section className="grid gap-5 border-b border-line px-4 py-5 last:border-b-0 lg:grid-cols-[220px_1fr] lg:gap-8 lg:px-5">
      <div className="lg:pt-0.5">
        <h3 className="flex items-center gap-2 text-sm font-semibold text-ink">
          {icon ? <span className="text-ink-muted">{icon}</span> : null}
          {title}
        </h3>
        {description ? (
          <p className="mt-1 text-xs leading-relaxed text-ink-muted">{description}</p>
        ) : null}
      </div>
      <div className="min-w-0">{children}</div>
    </section>
  );
}

/** Barre d'actions collante en pied de formulaire long. */
export function FormActions({ children }: { children: ReactNode }) {
  return (
    <div className="sticky bottom-0 z-10 flex items-center justify-end gap-2 border-t border-line bg-surface/85 px-4 py-3 backdrop-blur-sm">
      {children}
    </div>
  );
}

/** Liste de suggestions réutilisable pour les champs texte (registrars, technos…). */
export function DataList({ id, options }: { id: string; options: readonly string[] }) {
  return (
    <datalist id={id}>
      {options.map((option) => (
        <option key={option} value={option} />
      ))}
    </datalist>
  );
}
