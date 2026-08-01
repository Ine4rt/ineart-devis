"use server";

import type { BillingCycle, PeriodStatus, SubscriptionStatus, SubscriptionType } from "@prisma/client";
import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";

import { requireUser } from "@/lib/auth";
import { BILLING_CYCLE, PERIOD_STATUS, SUBSCRIPTION_STATUS, SUBSCRIPTION_TYPE } from "@/lib/constants";
import { db } from "@/lib/db";
import { bool, date, dec, enumOf, logActivity, optStr, requiredDate, str } from "@/lib/actions/helpers";
import { addCycle } from "@/lib/renewals";

const CYCLES = Object.keys(BILLING_CYCLE) as BillingCycle[];
const STATUSES = Object.keys(SUBSCRIPTION_STATUS) as SubscriptionStatus[];
const TYPES = Object.keys(SUBSCRIPTION_TYPE) as SubscriptionType[];
const PERIOD_STATUSES = Object.keys(PERIOD_STATUS) as PeriodStatus[];

/**
 * Abonnements et périodes de facturation.
 *
 * Le modèle sépare volontairement le *contrat* (Subscription) de ses
 * *échéances* (SubscriptionPeriod). C'est ce qui permet de répondre séparément
 * à « quel est le prochain renouvellement ? » et à « qui n'a pas payé la période
 * de l'an dernier ? » — deux questions différentes qu'un seul champ « payé »
 * ne saurait couvrir.
 */

export async function createSubscription(formData: FormData) {
  await requireUser();

  const clientId = str(formData, "clientId");
  const name = str(formData, "name");
  if (!clientId || !name) throw new Error("Client et intitulé sont obligatoires.");

  const startedAt = requiredDate(formData, "startedAt");
  const billingCycle = enumOf(formData, "billingCycle", CYCLES, "YEARLY");
  const amount = dec(formData, "amount");

  // La période courante démarre à la souscription et couvre un cycle complet.
  const currentPeriodStart = startedAt;
  const currentPeriodEnd = addCycle(currentPeriodStart, billingCycle);

  const subscription = await db.subscription.create({
    data: {
      clientId,
      projectId: optStr(formData, "projectId"),
      name,
      type: enumOf(formData, "type", TYPES, "FULL_CARE"),
      amount,
      currency: str(formData, "currency") || "EUR",
      billingCycle,
      status: enumOf(formData, "status", STATUSES, "ACTIVE"),
      autoRenew: bool(formData, "autoRenew"),
      startedAt,
      currentPeriodStart,
      currentPeriodEnd,
      nextRenewalAt: currentPeriodEnd,
      paymentMethod: optStr(formData, "paymentMethod"),
      notes: optStr(formData, "notes"),
      // Première échéance créée d'emblée : un abonnement sans période à payer
      // serait invisible dans le suivi des encaissements.
      periods: {
        create: {
          periodStart: currentPeriodStart,
          periodEnd: currentPeriodEnd,
          dueAt: date(formData, "firstDueAt") ?? currentPeriodStart,
          amount,
          status: bool(formData, "firstPeriodPaid") ? "PAID" : "DUE",
          paidAt: bool(formData, "firstPeriodPaid") ? new Date() : null,
        },
      },
    },
  });

  await logActivity({
    entityType: "subscription",
    entityId: subscription.id,
    clientId,
    action: "created",
    summary: `Abonnement « ${name} » souscrit par`,
  });

  revalidatePath("/abonnements");
  redirect(`/abonnements/${subscription.id}`);
}

export async function updateSubscription(subscriptionId: string, formData: FormData) {
  await requireUser();

  const status = enumOf(formData, "status", STATUSES, "ACTIVE");

  await db.subscription.update({
    where: { id: subscriptionId },
    data: {
      name: str(formData, "name"),
      type: enumOf(formData, "type", TYPES, "FULL_CARE"),
      amount: dec(formData, "amount"),
      currency: str(formData, "currency") || "EUR",
      billingCycle: enumOf(formData, "billingCycle", CYCLES, "YEARLY"),
      status,
      autoRenew: bool(formData, "autoRenew"),
      projectId: optStr(formData, "projectId"),
      startedAt: requiredDate(formData, "startedAt"),
      currentPeriodStart: requiredDate(formData, "currentPeriodStart"),
      currentPeriodEnd: requiredDate(formData, "currentPeriodEnd"),
      nextRenewalAt: requiredDate(formData, "nextRenewalAt"),
      cancelledAt: status === "CANCELLED" ? new Date() : null,
      paymentMethod: optStr(formData, "paymentMethod"),
      notes: optStr(formData, "notes"),
    },
  });

  revalidatePath(`/abonnements/${subscriptionId}`);
  revalidatePath("/abonnements");
  redirect(`/abonnements/${subscriptionId}`);
}

export async function deleteSubscription(subscriptionId: string) {
  await requireUser();
  await db.subscription.delete({ where: { id: subscriptionId } });
  revalidatePath("/abonnements");
  redirect("/abonnements");
}

/**
 * Reconduit l'abonnement d'un cycle : avance la période courante et ouvre la
 * prochaine échéance à payer. C'est l'action que l'on déclenche le jour du
 * renouvellement, en un clic depuis l'échéancier.
 */
export async function renewSubscription(subscriptionId: string) {
  await requireUser();

  const subscription = await db.subscription.findUnique({ where: { id: subscriptionId } });
  if (!subscription) return;

  const periodStart = subscription.currentPeriodEnd;
  const periodEnd = addCycle(periodStart, subscription.billingCycle);

  await db.$transaction([
    db.subscription.update({
      where: { id: subscriptionId },
      data: {
        currentPeriodStart: periodStart,
        currentPeriodEnd: periodEnd,
        nextRenewalAt: periodEnd,
        status: subscription.status === "CANCELLED" ? "CANCELLED" : "ACTIVE",
      },
    }),
    db.subscriptionPeriod.create({
      data: {
        subscriptionId,
        periodStart,
        periodEnd,
        dueAt: periodStart,
        amount: subscription.amount,
        status: "DUE",
      },
    }),
  ]);

  await logActivity({
    entityType: "subscription",
    entityId: subscriptionId,
    clientId: subscription.clientId,
    action: "renewed",
    summary: `Abonnement « ${subscription.name} » reconduit pour`,
  });

  revalidatePath(`/abonnements/${subscriptionId}`);
  revalidatePath("/paiements");
  revalidatePath("/echeancier");
}

// ---------------------------------------------------------------------------
// Périodes / encaissements
// ---------------------------------------------------------------------------

export async function markPeriodPaid(periodId: string, formData?: FormData) {
  await requireUser();

  const period = await db.subscriptionPeriod.update({
    where: { id: periodId },
    data: {
      status: "PAID",
      paidAt: new Date(),
      paymentMethod: formData ? optStr(formData, "paymentMethod") : undefined,
      reference: formData ? optStr(formData, "reference") : undefined,
    },
    include: { subscription: { select: { clientId: true, name: true } } },
  });

  // Un client redevient « à jour » dès lors qu'il ne reste aucun impayé.
  const remaining = await db.subscriptionPeriod.count({
    where: { subscriptionId: period.subscriptionId, status: { in: ["DUE", "OVERDUE"] } },
  });
  if (remaining === 0) {
    await db.subscription.updateMany({
      where: { id: period.subscriptionId, status: "PAST_DUE" },
      data: { status: "ACTIVE" },
    });
  }

  await logActivity({
    entityType: "subscription_period",
    entityId: periodId,
    clientId: period.subscription.clientId,
    action: "paid",
    summary: `Paiement enregistré pour « ${period.subscription.name} » —`,
  });

  revalidatePath("/paiements");
  revalidatePath(`/abonnements/${period.subscriptionId}`);
}

export async function setPeriodStatus(periodId: string, status: PeriodStatus) {
  await requireUser();
  if (!PERIOD_STATUSES.includes(status)) return;

  await db.subscriptionPeriod.update({
    where: { id: periodId },
    data: { status, paidAt: status === "PAID" ? new Date() : null },
  });

  revalidatePath("/paiements");
}

export async function createPeriod(subscriptionId: string, formData: FormData) {
  await requireUser();

  const subscription = await db.subscription.findUnique({ where: { id: subscriptionId } });
  if (!subscription) return;

  const periodStart = requiredDate(formData, "periodStart");

  await db.subscriptionPeriod.create({
    data: {
      subscriptionId,
      periodStart,
      periodEnd: date(formData, "periodEnd") ?? addCycle(periodStart, subscription.billingCycle),
      dueAt: date(formData, "dueAt") ?? periodStart,
      amount: dec(formData, "amount", Number(subscription.amount)),
      status: enumOf(formData, "status", PERIOD_STATUSES, "DUE"),
      notes: optStr(formData, "notes"),
    },
  });

  revalidatePath(`/abonnements/${subscriptionId}`);
  revalidatePath("/paiements");
}

export async function deletePeriod(periodId: string, subscriptionId: string) {
  await requireUser();
  await db.subscriptionPeriod.delete({ where: { id: periodId } });
  revalidatePath(`/abonnements/${subscriptionId}`);
  revalidatePath("/paiements");
}

/**
 * Bascule en « en retard » les échéances dont la date est dépassée.
 *
 * Appelée à l'ouverture des écrans concernés plutôt que par une tâche planifiée :
 * l'outil n'a pas de processus de fond, et le résultat est identique — le statut
 * est toujours juste au moment où on le regarde.
 */
export async function refreshOverdueStatuses(): Promise<void> {
  const now = new Date();

  await db.subscriptionPeriod.updateMany({
    where: { status: "DUE", dueAt: { lt: now } },
    data: { status: "OVERDUE" },
  });

  const overdueSubscriptionIds = await db.subscriptionPeriod.findMany({
    where: { status: "OVERDUE" },
    select: { subscriptionId: true },
    distinct: ["subscriptionId"],
  });

  if (overdueSubscriptionIds.length > 0) {
    await db.subscription.updateMany({
      where: {
        id: { in: overdueSubscriptionIds.map((row) => row.subscriptionId) },
        status: "ACTIVE",
      },
      data: { status: "PAST_DUE" },
    });
  }
}
