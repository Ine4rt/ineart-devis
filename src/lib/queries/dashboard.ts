import "server-only";

import type { SubscriptionType } from "@prisma/client";

import { SUBSCRIPTION_TYPE, LIVE_PROJECT_STATUSES } from "@/lib/constants";
import { db } from "@/lib/db";
import { toNumber } from "@/lib/format";
import { toMonthly, toYearly } from "@/lib/renewals";
import { getAnnualInfrastructureCost, getRenewalFeed } from "@/lib/queries/renewal-feed";

/**
 * Agrégats du tableau de bord.
 *
 * Tout est calculé côté serveur en une passe : la page reste un composant
 * serveur sans état de chargement, et les chiffres affichés sont toujours
 * cohérents entre eux (même instant de référence).
 */

export interface DashboardAlert {
  id: string;
  title: string;
  detail: string;
  tone: "danger" | "warning" | "info";
  href: string;
}

export async function getDashboardData() {
  const now = new Date();

  const [
    clientCounts,
    liveProjects,
    projectsInProgress,
    domainCount,
    hostingCount,
    subscriptions,
    unpaidPeriods,
    paidThisYear,
    openTasks,
    lateTasks,
    renewals,
    infrastructureCost,
    recentActivity,
    pipeline,
  ] = await Promise.all([
    db.client.groupBy({ by: ["status"], _count: { _all: true } }),
    db.project.count({ where: { status: { in: LIVE_PROJECT_STATUSES } } }),
    db.project.count({ where: { status: { in: ["DISCOVERY", "DESIGN", "DEVELOPMENT", "REVIEW"] } } }),
    db.domain.count(),
    db.hosting.count(),
    db.subscription.findMany({
      include: { client: { select: { id: true, company: true } } },
    }),
    db.subscriptionPeriod.findMany({
      where: { status: { in: ["DUE", "OVERDUE"] } },
      include: {
        subscription: {
          select: { name: true, currency: true, client: { select: { id: true, company: true } } },
        },
      },
      orderBy: { dueAt: "asc" },
    }),
    db.subscriptionPeriod.findMany({
      where: { status: "PAID", paidAt: { gte: new Date(now.getFullYear(), 0, 1) } },
      select: { amount: true, paidAt: true },
    }),
    db.task.findMany({
      where: { status: { not: "DONE" } },
      include: { client: { select: { id: true, company: true } } },
      orderBy: [{ dueAt: "asc" }],
      take: 8,
    }),
    db.task.count({ where: { status: { not: "DONE" }, dueAt: { lt: now } } }),
    getRenewalFeed(120),
    getAnnualInfrastructureCost(),
    db.activity.findMany({
      orderBy: { createdAt: "desc" },
      take: 8,
      include: { client: { select: { id: true, company: true } } },
    }),
    db.client.findMany({
      where: { status: "PROSPECT" },
      select: { id: true, pipelineStage: true, potentialValue: true },
    }),
  ]);

  const countByStatus = Object.fromEntries(
    clientCounts.map((row) => [row.status, row._count._all]),
  ) as Record<string, number>;

  const activeSubscriptions = subscriptions.filter((s) => s.status === "ACTIVE");

  // --- Revenus récurrents -------------------------------------------------
  const mrr = activeSubscriptions.reduce(
    (sum, subscription) => sum + toMonthly(toNumber(subscription.amount), subscription.billingCycle),
    0,
  );
  const arr = activeSubscriptions.reduce(
    (sum, subscription) => sum + toYearly(toNumber(subscription.amount), subscription.billingCycle),
    0,
  );

  const revenueByType = new Map<SubscriptionType, number>();
  for (const subscription of activeSubscriptions) {
    const yearly = toYearly(toNumber(subscription.amount), subscription.billingCycle);
    revenueByType.set(subscription.type, (revenueByType.get(subscription.type) ?? 0) + yearly);
  }

  const revenueMix = [...revenueByType.entries()]
    .map(([type, value]) => ({ label: SUBSCRIPTION_TYPE[type].label, value }))
    .sort((a, b) => b.value - a.value);

  // --- Encaissements ------------------------------------------------------
  const unpaidTotal = unpaidPeriods.reduce((sum, period) => sum + toNumber(period.amount), 0);
  const overduePeriods = unpaidPeriods.filter((period) => period.dueAt < now);
  const overdueTotal = overduePeriods.reduce((sum, period) => sum + toNumber(period.amount), 0);

  // Encaissements mois par mois sur l'année civile en cours.
  const monthlyCollected = Array.from({ length: 12 }, (_, month) => ({
    label: new Intl.DateTimeFormat("fr-BE", { month: "narrow" }).format(
      new Date(now.getFullYear(), month, 1),
    ),
    value: 0,
  }));
  for (const period of paidThisYear) {
    if (!period.paidAt) continue;
    monthlyCollected[period.paidAt.getMonth()].value += toNumber(period.amount);
  }

  // --- Trajectoire de l'ARR sur 12 mois glissants -------------------------
  // Reconstruite depuis les dates de souscription/résiliation : pas besoin de
  // stocker un historique, la donnée existante suffit.
  const arrTrend = Array.from({ length: 12 }, (_, index) => {
    const monthEnd = new Date(now.getFullYear(), now.getMonth() - (11 - index) + 1, 0);
    const value = subscriptions.reduce((sum, subscription) => {
      const started = subscription.startedAt <= monthEnd;
      const ended = subscription.cancelledAt !== null && subscription.cancelledAt <= monthEnd;
      if (!started || ended) return sum;
      return sum + toYearly(toNumber(subscription.amount), subscription.billingCycle);
    }, 0);
    return {
      label: new Intl.DateTimeFormat("fr-BE", { month: "short" }).format(monthEnd),
      value,
    };
  });

  // --- Alertes ------------------------------------------------------------
  const alerts: DashboardAlert[] = [];

  if (overduePeriods.length > 0) {
    alerts.push({
      id: "overdue-payments",
      title: `${overduePeriods.length} paiement${overduePeriods.length > 1 ? "s" : ""} en retard`,
      detail: `${overdueTotal.toFixed(0)} € à recouvrer auprès de ${
        new Set(overduePeriods.map((p) => p.subscription.client.id)).size
      } client(s).`,
      tone: "danger",
      href: "/paiements?statut=OVERDUE",
    });
  }

  const criticalRenewals = renewals.filter((item) => item.days <= 15);
  if (criticalRenewals.length > 0) {
    alerts.push({
      id: "critical-renewals",
      title: `${criticalRenewals.length} échéance${criticalRenewals.length > 1 ? "s" : ""} sous 15 jours`,
      detail: "Domaines, hébergements ou abonnements à renouveler très prochainement.",
      tone: "danger",
      href: "/echeancier",
    });
  }

  const manualRenewals = renewals.filter((item) => !item.autoRenew && item.days <= 30);
  if (manualRenewals.length > 0) {
    alerts.push({
      id: "manual-renewals",
      title: `${manualRenewals.length} renouvellement${manualRenewals.length > 1 ? "s" : ""} sans reconduction automatique`,
      detail: "Sans action manuelle de votre part, ces éléments expireront.",
      tone: "warning",
      href: "/echeancier?auto=manuel",
    });
  }

  if (lateTasks > 0) {
    alerts.push({
      id: "late-tasks",
      title: `${lateTasks} tâche${lateTasks > 1 ? "s" : ""} en retard`,
      detail: "Des échéances sont dépassées dans votre liste de travail.",
      tone: "warning",
      href: "/taches?filtre=retard",
    });
  }

  // Angle mort classique : un site en ligne qui ne génère aucun revenu récurrent.
  const uncoveredProjects = await db.project.count({
    where: { status: { in: LIVE_PROJECT_STATUSES }, subscriptions: { none: {} } },
  });
  if (uncoveredProjects > 0) {
    alerts.push({
      id: "uncovered-projects",
      title: `${uncoveredProjects} site${uncoveredProjects > 1 ? "s" : ""} en ligne sans abonnement`,
      detail: "Ces sites sont maintenus sans contrepartie récurrente facturée.",
      tone: "info",
      href: "/projets?couverture=sans-abonnement",
    });
  }

  const pipelineValue = pipeline.reduce((sum, prospect) => sum + toNumber(prospect.potentialValue), 0);

  return {
    now,
    clients: {
      total: clientCounts.reduce((sum, row) => sum + row._count._all, 0),
      active: countByStatus.ACTIVE ?? 0,
      prospects: countByStatus.PROSPECT ?? 0,
      paused: countByStatus.PAUSED ?? 0,
    },
    projects: { live: liveProjects, inProgress: projectsInProgress },
    infrastructure: { domains: domainCount, hostings: hostingCount, annualCost: infrastructureCost },
    revenue: {
      mrr,
      arr,
      margin: arr - infrastructureCost,
      marginRate: arr > 0 ? (arr - infrastructureCost) / arr : 0,
      mix: revenueMix,
      trend: arrTrend,
      monthlyCollected,
      collectedThisYear: paidThisYear.reduce((sum, period) => sum + toNumber(period.amount), 0),
    },
    billing: {
      activeSubscriptions: activeSubscriptions.length,
      unpaidCount: unpaidPeriods.length,
      unpaidTotal,
      overdueCount: overduePeriods.length,
      overdueTotal,
      upcomingUnpaid: unpaidPeriods.slice(0, 6),
    },
    renewals,
    tasks: { open: openTasks, lateCount: lateTasks },
    alerts,
    activity: recentActivity,
    pipeline: { count: pipeline.length, value: pipelineValue },
  };
}
