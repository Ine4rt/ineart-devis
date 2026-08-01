import "server-only";

import { db } from "@/lib/db";
import { toNumber } from "@/lib/format";
import { daysUntil, toYearly, urgencyFromDays, type UrgencyMeta } from "@/lib/renewals";

/**
 * Flux unifié des échéances.
 *
 * Abonnements, domaines, hébergements et certificats SSL expirent tous, mais
 * vivent dans quatre tables. Les fusionner ici en une seule liste triée par
 * urgence évite de reproduire quatre fois la même logique de rappel — et
 * surtout permet de répondre à la seule question qui compte le matin :
 * « qu'est-ce qui arrive à échéance ? ».
 */

export type RenewalKind = "subscription" | "domain" | "hosting" | "ssl";

export interface RenewalItem {
  id: string;
  kind: RenewalKind;
  label: string;
  sublabel: string;
  date: Date;
  days: number;
  urgency: UrgencyMeta;
  /** Montant concerné : facturé au client (abonnement) ou payé par moi (coût). */
  amount: number;
  currency: string;
  /** true = revenu entrant ; false = dépense sortante. */
  isRevenue: boolean;
  autoRenew: boolean;
  href: string;
  clientId: string | null;
  clientName: string | null;
}

export const RENEWAL_KIND_LABEL: Record<RenewalKind, string> = {
  subscription: "Abonnement",
  domain: "Domaine",
  hosting: "Hébergement",
  ssl: "Certificat SSL",
};

/**
 * @param horizonDays fenêtre en avant. Les éléments déjà en retard sont toujours
 *   inclus, quel que soit l'horizon : un oubli ne doit jamais disparaître.
 */
export async function getRenewalFeed(horizonDays = 120): Promise<RenewalItem[]> {
  const horizon = new Date();
  horizon.setDate(horizon.getDate() + horizonDays);

  const [subscriptions, domains, hostings] = await Promise.all([
    db.subscription.findMany({
      where: { status: { in: ["ACTIVE", "PENDING", "PAST_DUE"] }, nextRenewalAt: { lte: horizon } },
      include: { client: { select: { id: true, company: true } } },
    }),
    db.domain.findMany({
      where: { expiresAt: { not: null, lte: horizon } },
      include: { client: { select: { id: true, company: true } } },
    }),
    db.hosting.findMany({
      where: { renewsAt: { not: null, lte: horizon } },
      include: { client: { select: { id: true, company: true } } },
    }),
  ]);

  const items: RenewalItem[] = [];

  for (const subscription of subscriptions) {
    const days = daysUntil(subscription.nextRenewalAt);
    items.push({
      id: `sub-${subscription.id}`,
      kind: "subscription",
      label: subscription.name,
      sublabel: subscription.client.company,
      date: subscription.nextRenewalAt,
      days,
      urgency: urgencyFromDays(days),
      amount: toNumber(subscription.amount),
      currency: subscription.currency,
      isRevenue: true,
      autoRenew: subscription.autoRenew,
      href: `/abonnements/${subscription.id}`,
      clientId: subscription.client.id,
      clientName: subscription.client.company,
    });
  }

  for (const domain of domains) {
    if (!domain.expiresAt) continue;
    const days = daysUntil(domain.expiresAt);
    items.push({
      id: `dom-${domain.id}`,
      kind: "domain",
      label: domain.name,
      sublabel: domain.registrar ?? "Registrar non renseigné",
      date: domain.expiresAt,
      days,
      urgency: urgencyFromDays(days),
      amount: toNumber(domain.renewalPrice),
      currency: "EUR",
      isRevenue: false,
      autoRenew: domain.autoRenew,
      href: `/domaines/${domain.id}`,
      clientId: domain.client?.id ?? null,
      clientName: domain.client?.company ?? null,
    });

    // Le SSL est suivi séparément : il expire souvent avant le domaine.
    if (domain.sslExpiresAt && domain.sslExpiresAt <= horizon) {
      const sslDays = daysUntil(domain.sslExpiresAt);
      items.push({
        id: `ssl-${domain.id}`,
        kind: "ssl",
        label: `SSL — ${domain.name}`,
        sublabel: domain.sslProvider ?? "Émetteur non renseigné",
        date: domain.sslExpiresAt,
        days: sslDays,
        urgency: urgencyFromDays(sslDays),
        amount: 0,
        currency: "EUR",
        isRevenue: false,
        autoRenew: domain.sslAutoRenew,
        href: `/domaines/${domain.id}`,
        clientId: domain.client?.id ?? null,
        clientName: domain.client?.company ?? null,
      });
    }
  }

  for (const hosting of hostings) {
    if (!hosting.renewsAt) continue;
    const days = daysUntil(hosting.renewsAt);
    items.push({
      id: `host-${hosting.id}`,
      kind: "hosting",
      label: [hosting.provider, hosting.plan].filter(Boolean).join(" — "),
      sublabel: hosting.client?.company ?? hosting.label ?? "Sans client",
      date: hosting.renewsAt,
      days,
      urgency: urgencyFromDays(days),
      amount: toNumber(hosting.price),
      currency: "EUR",
      isRevenue: false,
      autoRenew: hosting.autoRenew,
      href: `/hebergements/${hosting.id}`,
      clientId: hosting.client?.id ?? null,
      clientName: hosting.client?.company ?? null,
    });
  }

  return items.sort((a, b) => a.days - b.days);
}

/** Coût annuel total de l'infrastructure (domaines + hébergements). */
export async function getAnnualInfrastructureCost(): Promise<number> {
  const [domains, hostings] = await Promise.all([
    db.domain.findMany({ select: { renewalPrice: true } }),
    db.hosting.findMany({ select: { price: true, billingCycle: true } }),
  ]);

  const domainCost = domains.reduce((sum, domain) => sum + toNumber(domain.renewalPrice), 0);
  const hostingCost = hostings.reduce(
    (sum, hosting) => sum + toYearly(toNumber(hosting.price), hosting.billingCycle),
    0,
  );

  return domainCost + hostingCost;
}
