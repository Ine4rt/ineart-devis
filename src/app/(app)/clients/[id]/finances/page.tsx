import { Check, FileText, Plus, Waypoints } from "lucide-react";
import Link from "next/link";

import { EnumBadge } from "@/components/ui/badge";
import { Button, ButtonLink } from "@/components/ui/button";
import { Card, CardHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { StatTile } from "@/components/ui/stat-tile";
import {
  BILLING_CYCLE,
  PERIOD_STATUS,
  QUOTE_STATUS,
  SUBSCRIPTION_STATUS,
} from "@/lib/constants";
import { markPeriodPaid, refreshOverdueStatuses } from "@/lib/actions/subscriptions";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { computeQuoteTotals } from "@/lib/quotes";
import { countdownLabel, daysUntil, toYearly } from "@/lib/renewals";

export const dynamic = "force-dynamic";

/**
 * Onglet finances : abonnements, échéances et devis du client.
 * Répond en un écran à « où en est-on financièrement avec ce client ».
 */
export default async function ClientFinancesPage({
  params,
}: {
  params: Promise<{ id: string }>;
}) {
  await refreshOverdueStatuses();
  const { id } = await params;

  const [subscriptions, periods, quotes] = await Promise.all([
    db.subscription.findMany({
      where: { clientId: id },
      orderBy: { nextRenewalAt: "asc" },
    }),
    db.subscriptionPeriod.findMany({
      where: { subscription: { clientId: id } },
      orderBy: { dueAt: "desc" },
      include: { subscription: { select: { id: true, name: true, currency: true } } },
    }),
    db.quote.findMany({
      where: { clientId: id },
      orderBy: { issuedAt: "desc" },
      include: { items: { select: { quantity: true, unitPrice: true } } },
    }),
  ]);

  const active = subscriptions.filter((subscription) => subscription.status === "ACTIVE");
  const arr = active.reduce(
    (sum, subscription) => sum + toYearly(toNumber(subscription.amount), subscription.billingCycle),
    0,
  );
  const paid = periods.filter((period) => period.status === "PAID");
  const unpaid = periods.filter((period) => ["DUE", "OVERDUE"].includes(period.status));
  const collected = paid.reduce((sum, period) => sum + toNumber(period.amount), 0);
  const outstanding = unpaid.reduce((sum, period) => sum + toNumber(period.amount), 0);

  return (
    <div className="space-y-4">
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile label="Revenu annuel" value={formatMoney(arr)} tone="success" />
        <StatTile label="Total encaissé" value={formatMoney(collected)} hint={`${paid.length} paiement(s)`} />
        <StatTile
          label="Reste à encaisser"
          value={formatMoney(outstanding)}
          tone={outstanding > 0 ? "danger" : "neutral"}
          hint={`${unpaid.length} échéance(s)`}
        />
        <StatTile label="Abonnements actifs" value={active.length} />
      </div>

      <Card>
        <CardHeader
          title="Abonnements"
          icon={<Waypoints className="h-4 w-4" />}
          action={
            <ButtonLink href={`/abonnements/nouveau?client=${id}`} size="sm" variant="primary">
              <Plus className="h-3.5 w-3.5" />
              Nouvel abonnement
            </ButtonLink>
          }
        />
        <div className="divide-y divide-line">
          {subscriptions.length > 0 ? (
            subscriptions.map((subscription) => (
              <Link
                key={subscription.id}
                href={`/abonnements/${subscription.id}`}
                className="flex flex-wrap items-center gap-3 px-4 py-2.5 transition-colors hover:bg-surface-inset"
              >
                <div className="min-w-[180px] flex-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="text-[13px] font-medium text-ink">{subscription.name}</span>
                    <EnumBadge meta={SUBSCRIPTION_STATUS[subscription.status]} />
                  </div>
                  <p className="mt-0.5 text-xs text-ink-muted">
                    Période {formatDate(subscription.currentPeriodStart, "short")} →{" "}
                    {formatDate(subscription.currentPeriodEnd, "short")}
                  </p>
                </div>
                <div className="text-right">
                  <p className="text-[13px] font-medium text-ink tabular-nums">
                    {formatMoney(subscription.amount, { currency: subscription.currency })}
                  </p>
                  <p className="text-[11px] text-ink-muted">
                    {BILLING_CYCLE[subscription.billingCycle].label}
                  </p>
                </div>
                <div className="w-28 text-right text-xs text-ink-muted">
                  {countdownLabel(daysUntil(subscription.nextRenewalAt))}
                </div>
              </Link>
            ))
          ) : (
            <EmptyState
              compact
              title="Aucun abonnement"
              description="Ce client ne génère aucun revenu récurrent."
              action={
                <ButtonLink href={`/abonnements/nouveau?client=${id}`} size="sm" variant="primary">
                  Créer un abonnement
                </ButtonLink>
              }
            />
          )}
        </div>
      </Card>

      <Card>
        <CardHeader title="Historique de paiement" description="Toutes les périodes facturées." />
        <div className="divide-y divide-line">
          {periods.length > 0 ? (
            periods.map((period) => {
              const markPaid = markPeriodPaid.bind(null, period.id);
              const isOpen = ["DUE", "OVERDUE"].includes(period.status);
              return (
                <div key={period.id} className="flex flex-wrap items-center gap-3 px-4 py-2.5">
                  <div className="min-w-[180px] flex-1">
                    <p className="text-[13px] text-ink">
                      {formatDate(period.periodStart, "short")}
                      <span className="mx-1.5 text-ink-muted">→</span>
                      {formatDate(period.periodEnd, "short")}
                    </p>
                    <p className="text-xs text-ink-muted">{period.subscription.name}</p>
                  </div>
                  <EnumBadge meta={PERIOD_STATUS[period.status]} />
                  <span className="w-24 text-right text-[13px] font-medium text-ink tabular-nums">
                    {formatMoney(period.amount, { currency: period.subscription.currency })}
                  </span>
                  {isOpen ? (
                    <form action={markPaid}>
                      <Button type="submit" size="sm">
                        <Check className="h-3.5 w-3.5" />
                        Payé
                      </Button>
                    </form>
                  ) : (
                    <span className="w-[76px] text-right text-[11px] text-ink-muted">
                      {formatDate(period.paidAt, "short")}
                    </span>
                  )}
                </div>
              );
            })
          ) : (
            <EmptyState compact title="Aucune échéance enregistrée" />
          )}
        </div>
      </Card>

      <Card>
        <CardHeader
          title="Devis"
          icon={<FileText className="h-4 w-4" />}
          action={
            <ButtonLink href={`/devis/nouveau?client=${id}`} size="sm">
              <Plus className="h-3.5 w-3.5" />
              Nouveau devis
            </ButtonLink>
          }
        />
        <div className="divide-y divide-line">
          {quotes.length > 0 ? (
            quotes.map((quote) => {
              const totals = computeQuoteTotals(
                quote.items.map((item) => ({
                  quantity: toNumber(item.quantity),
                  unitPrice: toNumber(item.unitPrice),
                })),
                toNumber(quote.discountPct),
                toNumber(quote.taxRate),
              );
              return (
                <Link
                  key={quote.id}
                  href={`/devis/${quote.id}`}
                  className="flex flex-wrap items-center gap-3 px-4 py-2.5 transition-colors hover:bg-surface-inset"
                >
                  <div className="min-w-[180px] flex-1">
                    <p className="text-[13px] font-medium text-ink">{quote.number}</p>
                    <p className="truncate text-xs text-ink-muted">{quote.title}</p>
                  </div>
                  <EnumBadge meta={QUOTE_STATUS[quote.status]} />
                  <span className="w-24 text-right text-[13px] font-medium text-ink tabular-nums">
                    {formatMoney(totals.total, { currency: quote.currency })}
                  </span>
                  <span className="w-20 text-right text-[11px] text-ink-muted">
                    {formatDate(quote.issuedAt, "short")}
                  </span>
                </Link>
              );
            })
          ) : (
            <EmptyState compact title="Aucun devis" description="Aucune proposition envoyée à ce client." />
          )}
        </div>
      </Card>
    </div>
  );
}
