import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

/** Fusionne des classes Tailwind en résolvant les conflits (`p-2` + `p-4` → `p-4`). */
export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

/** Découpe une valeur JSON stockée en texte, avec repli sûr. */
export function parseJsonArray(value: string | null | undefined): string[] {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed.filter((v): v is string => typeof v === "string") : [];
  } catch {
    return [];
  }
}

/** Transforme "Next.js, Tailwind , Prisma" en tableau nettoyé. */
export function splitList(value: string | null | undefined): string[] {
  if (!value) return [];
  return value
    .split(/[,\n]/)
    .map((v) => v.trim())
    .filter(Boolean);
}

/** Initiales affichées dans les avatars (2 lettres max). */
export function initials(...parts: (string | null | undefined)[]): string {
  const source = parts.filter(Boolean).join(" ").trim();
  if (!source) return "?";
  const words = source.split(/\s+/).slice(0, 2);
  return words.map((w) => w[0]?.toUpperCase() ?? "").join("") || "?";
}

/**
 * Couleur d'avatar déterministe dérivée d'une chaîne : le même client garde
 * toujours la même teinte, sans avoir à la stocker.
 */
export function hueFromString(value: string): number {
  let hash = 0;
  for (let i = 0; i < value.length; i += 1) {
    hash = (hash << 5) - hash + value.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash) % 360;
}

/** Retire les accents et met en minuscules — utilisé par la recherche globale. */
export function normalize(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

export function slugify(value: string): string {
  return normalize(value)
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

/** Génère la prochaine référence séquentielle (CLI-0007) à partir de la dernière. */
export function nextReference(prefix: string, last: string | null | undefined): string {
  const current = last ? Number.parseInt(last.split("-")[1] ?? "0", 10) : 0;
  const next = Number.isFinite(current) ? current + 1 : 1;
  return `${prefix}-${String(next).padStart(4, "0")}`;
}
