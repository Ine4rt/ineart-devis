import type { Prisma } from "@prisma/client";

/**
 * Formatage centralisé. Toute valeur affichée (argent, date, taille de fichier)
 * passe par ici : une seule locale, un seul comportement, zéro divergence
 * d'un écran à l'autre.
 */

export const LOCALE = "fr-BE";

/** Convertit un Decimal Prisma (ou tout scalaire numérique) en number sûr. */
export function toNumber(value: Prisma.Decimal | number | string | null | undefined): number {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return value;
  const parsed = Number(value.toString());
  return Number.isFinite(parsed) ? parsed : 0;
}

const currencyFormatters = new Map<string, Intl.NumberFormat>();

function currencyFormatter(currency: string, compact: boolean) {
  const key = `${currency}:${compact}`;
  let formatter = currencyFormatters.get(key);
  if (!formatter) {
    formatter = new Intl.NumberFormat(LOCALE, {
      style: "currency",
      currency,
      notation: compact ? "compact" : "standard",
      maximumFractionDigits: compact ? 1 : 2,
      minimumFractionDigits: compact ? 0 : 2,
    });
    currencyFormatters.set(key, formatter);
  }
  return formatter;
}

export function formatMoney(
  value: Prisma.Decimal | number | string | null | undefined,
  options: { currency?: string; compact?: boolean } = {},
): string {
  const { currency = "EUR", compact = false } = options;
  return currencyFormatter(currency, compact).format(toNumber(value));
}

export function formatNumber(value: number, maximumFractionDigits = 0): string {
  return new Intl.NumberFormat(LOCALE, { maximumFractionDigits }).format(value);
}

export function formatPercent(value: number, maximumFractionDigits = 0): string {
  return new Intl.NumberFormat(LOCALE, {
    style: "percent",
    maximumFractionDigits,
  }).format(value);
}

export function formatDate(
  value: Date | string | null | undefined,
  style: "short" | "medium" | "long" = "medium",
): string {
  if (!value) return "—";
  const date = typeof value === "string" ? new Date(value) : value;
  if (Number.isNaN(date.getTime())) return "—";

  const options: Intl.DateTimeFormatOptions =
    style === "short"
      ? { day: "2-digit", month: "2-digit", year: "2-digit" }
      : style === "long"
        ? { day: "numeric", month: "long", year: "numeric" }
        : { day: "2-digit", month: "short", year: "numeric" };

  return new Intl.DateTimeFormat(LOCALE, options).format(date);
}

export function formatDateTime(value: Date | string | null | undefined): string {
  if (!value) return "—";
  const date = typeof value === "string" ? new Date(value) : value;
  if (Number.isNaN(date.getTime())) return "—";
  return new Intl.DateTimeFormat(LOCALE, {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

/** « il y a 3 jours », « dans 2 mois » — pour les timelines et les échéances. */
export function formatRelative(value: Date | string | null | undefined): string {
  if (!value) return "—";
  const date = typeof value === "string" ? new Date(value) : value;
  if (Number.isNaN(date.getTime())) return "—";

  const diffMs = date.getTime() - Date.now();
  const diffDays = Math.round(diffMs / 86_400_000);
  const rtf = new Intl.RelativeTimeFormat(LOCALE, { numeric: "auto" });

  const absDays = Math.abs(diffDays);
  if (absDays < 1) {
    const diffHours = Math.round(diffMs / 3_600_000);
    if (Math.abs(diffHours) < 1) return rtf.format(Math.round(diffMs / 60_000), "minute");
    return rtf.format(diffHours, "hour");
  }
  if (absDays < 31) return rtf.format(diffDays, "day");
  if (absDays < 365) return rtf.format(Math.round(diffDays / 30), "month");
  return rtf.format(Math.round(diffDays / 365), "year");
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} o`;
  const units = ["Ko", "Mo", "Go"];
  let value = bytes / 1024;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[unitIndex]}`;
}

/** Valeur pour un `<input type="date">` (yyyy-MM-dd) sans décalage de fuseau. */
export function toDateInputValue(value: Date | string | null | undefined): string {
  if (!value) return "";
  const date = typeof value === "string" ? new Date(value) : value;
  if (Number.isNaN(date.getTime())) return "";
  const offset = date.getTimezoneOffset() * 60_000;
  return new Date(date.getTime() - offset).toISOString().slice(0, 10);
}
