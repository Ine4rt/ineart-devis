import type { EnumMeta } from "@/lib/constants";

/**
 * Construit les options d'un filtre à partir d'une table d'enum de constants.ts.
 *
 * Volontairement hors de `list-toolbar.tsx` : ce fichier porte la directive
 * « use client », et une fonction qui y serait exportée ne pourrait pas être
 * appelée depuis un composant serveur — or ce sont les pages serveur qui
 * composent les filtres.
 */
export interface FilterOption {
  value: string;
  label: string;
}

export function filterOptions<T extends string>(entries: [T, EnumMeta][]): FilterOption[] {
  return entries.map(([value, meta]) => ({ value, label: meta.label }));
}
