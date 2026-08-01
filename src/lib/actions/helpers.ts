import "server-only";

import { db } from "@/lib/db";
import { nextReference } from "@/lib/utils";

/**
 * Lecture de FormData.
 *
 * Les formulaires HTML ne renvoient que des chaînes : ces helpers centralisent
 * la conversion (chaîne vide → null, virgule décimale, case cochée, valeur
 * d'enum inconnue → repli). Sans eux, chaque action réécrirait les mêmes
 * dix lignes fragiles.
 */

export function str(formData: FormData, key: string): string {
  return (formData.get(key) ?? "").toString().trim();
}

/** Chaîne optionnelle : une valeur vide devient `null`, pas `""`. */
export function optStr(formData: FormData, key: string): string | null {
  const value = str(formData, key);
  return value.length > 0 ? value : null;
}

export function date(formData: FormData, key: string): Date | null {
  const value = str(formData, key);
  if (!value) return null;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function requiredDate(formData: FormData, key: string, fallback: Date = new Date()): Date {
  return date(formData, key) ?? fallback;
}

/** Nombre décimal tolérant à la virgule française et aux espaces. */
export function dec(formData: FormData, key: string, fallback = 0): number {
  const raw = str(formData, key).replace(/\s/g, "").replace(",", ".");
  if (!raw) return fallback;
  const parsed = Number.parseFloat(raw);
  return Number.isFinite(parsed) ? parsed : fallback;
}

export function int(formData: FormData, key: string, fallback = 0): number {
  const parsed = Number.parseInt(str(formData, key), 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

/** Une case non cochée n'est pas envoyée par le navigateur : absence = false. */
export function bool(formData: FormData, key: string): boolean {
  const value = formData.get(key);
  return value === "on" || value === "true" || value === "1";
}

/** Valide une valeur d'enum ; toute valeur inattendue retombe sur `fallback`. */
export function enumOf<T extends string>(
  formData: FormData,
  key: string,
  allowed: readonly T[],
  fallback: T,
): T {
  const value = str(formData, key) as T;
  return allowed.includes(value) ? value : fallback;
}

export function optEnumOf<T extends string>(
  formData: FormData,
  key: string,
  allowed: readonly T[],
): T | null {
  const value = str(formData, key) as T;
  return allowed.includes(value) ? value : null;
}

/** Liste saisie librement (« Next.js, Tailwind ») sérialisée en JSON. */
export function jsonList(formData: FormData, key: string): string {
  const items = str(formData, key)
    .split(/[,\n]/)
    .map((item) => item.trim())
    .filter(Boolean);
  return JSON.stringify(items);
}

// ---------------------------------------------------------------------------
// Journal d'activité
// ---------------------------------------------------------------------------

/**
 * Trace une action dans l'historique. Volontairement silencieux en cas d'échec :
 * un problème d'écriture du journal ne doit jamais faire échouer l'opération
 * métier qui vient de réussir.
 */
export async function logActivity(entry: {
  entityType: string;
  entityId: string;
  clientId?: string | null;
  action: string;
  summary: string;
}): Promise<void> {
  try {
    await db.activity.create({
      data: {
        entityType: entry.entityType,
        entityId: entry.entityId,
        clientId: entry.clientId ?? null,
        action: entry.action,
        summary: entry.summary,
      },
    });
  } catch {
    /* le journal est un confort, jamais un point de rupture */
  }
}

// ---------------------------------------------------------------------------
// Références séquentielles
// ---------------------------------------------------------------------------

export async function nextClientReference(): Promise<string> {
  const last = await db.client.findFirst({
    orderBy: { reference: "desc" },
    select: { reference: true },
  });
  return nextReference("CLI", last?.reference);
}

export async function nextQuoteNumber(): Promise<string> {
  const year = new Date().getFullYear();
  const prefix = `DEV-${year}`;
  const last = await db.quote.findFirst({
    where: { number: { startsWith: prefix } },
    orderBy: { number: "desc" },
    select: { number: true },
  });
  const current = last ? Number.parseInt(last.number.split("-")[2] ?? "0", 10) : 0;
  const next = Number.isFinite(current) ? current + 1 : 1;
  return `${prefix}-${String(next).padStart(3, "0")}`;
}
