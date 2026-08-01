import { BillingCycle } from "@prisma/client";

/**
 * Moteur d'échéances : convertit une date de renouvellement en « palier
 * d'alerte » (90 / 60 / 30 / 15 / 7 jours, aujourd'hui, en retard).
 *
 * Toute la logique de rappel de l'application dérive d'ici — dashboard,
 * abonnements, domaines et hébergements partagent le même vocabulaire visuel.
 */

export const ALERT_THRESHOLDS = [90, 60, 30, 15, 7] as const;

export type UrgencyLevel = "overdue" | "today" | "d7" | "d15" | "d30" | "d60" | "d90" | "later";

export interface UrgencyMeta {
  level: UrgencyLevel;
  label: string;
  /** Tonalité partagée avec les Badge/Card pour rester cohérent partout. */
  tone: "danger" | "critical" | "warning" | "caution" | "info" | "neutral";
  /** Poids de tri : plus c'est petit, plus c'est urgent. */
  weight: number;
}

const URGENCY: Record<UrgencyLevel, Omit<UrgencyMeta, "level">> = {
  overdue: { label: "En retard", tone: "danger", weight: 0 },
  today: { label: "Aujourd'hui", tone: "danger", weight: 1 },
  d7: { label: "≤ 7 jours", tone: "critical", weight: 2 },
  d15: { label: "≤ 15 jours", tone: "warning", weight: 3 },
  d30: { label: "≤ 30 jours", tone: "warning", weight: 4 },
  d60: { label: "≤ 60 jours", tone: "caution", weight: 5 },
  d90: { label: "≤ 90 jours", tone: "info", weight: 6 },
  later: { label: "Plus tard", tone: "neutral", weight: 7 },
};

/** Nombre de jours calendaires entre aujourd'hui et `date` (négatif = passé). */
export function daysUntil(date: Date | string | null | undefined, from: Date = new Date()): number {
  if (!date) return Number.POSITIVE_INFINITY;
  const target = typeof date === "string" ? new Date(date) : date;
  if (Number.isNaN(target.getTime())) return Number.POSITIVE_INFINITY;

  const startOfTarget = Date.UTC(target.getFullYear(), target.getMonth(), target.getDate());
  const startOfFrom = Date.UTC(from.getFullYear(), from.getMonth(), from.getDate());
  return Math.round((startOfTarget - startOfFrom) / 86_400_000);
}

export function urgencyFromDays(days: number): UrgencyMeta {
  const level: UrgencyLevel =
    days < 0
      ? "overdue"
      : days === 0
        ? "today"
        : days <= 7
          ? "d7"
          : days <= 15
            ? "d15"
            : days <= 30
              ? "d30"
              : days <= 60
                ? "d60"
                : days <= 90
                  ? "d90"
                  : "later";
  return { level, ...URGENCY[level] };
}

export function urgencyFor(date: Date | string | null | undefined): UrgencyMeta {
  return urgencyFromDays(daysUntil(date));
}

/** Libellé humain : « dans 12 jours », « en retard de 3 jours ». */
export function countdownLabel(days: number): string {
  if (!Number.isFinite(days)) return "Aucune échéance";
  if (days < 0) return `En retard de ${Math.abs(days)} j`;
  if (days === 0) return "Aujourd'hui";
  if (days === 1) return "Demain";
  if (days < 60) return `Dans ${days} j`;
  const months = Math.round(days / 30);
  return `Dans ${months} mois`;
}

// ---------------------------------------------------------------------------
// Cycles de facturation
// ---------------------------------------------------------------------------

/** Nombre de mois couverts par un cycle — base des conversions MRR/ARR. */
export const CYCLE_MONTHS: Record<BillingCycle, number> = {
  MONTHLY: 1,
  QUARTERLY: 3,
  YEARLY: 12,
  BIENNIAL: 24,
};

/** Ramène n'importe quel montant cyclique à un équivalent mensuel. */
export function toMonthly(amount: number, cycle: BillingCycle): number {
  return amount / CYCLE_MONTHS[cycle];
}

/** Ramène n'importe quel montant cyclique à un équivalent annuel. */
export function toYearly(amount: number, cycle: BillingCycle): number {
  return (amount * 12) / CYCLE_MONTHS[cycle];
}

/** Avance une date d'un cycle complet (utilisé au renouvellement). */
export function addCycle(date: Date, cycle: BillingCycle): Date {
  const next = new Date(date);
  next.setMonth(next.getMonth() + CYCLE_MONTHS[cycle]);
  return next;
}
