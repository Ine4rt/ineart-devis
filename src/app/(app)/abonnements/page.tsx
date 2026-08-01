import type { Prisma, SubscriptionStatus, SubscriptionType } from "@prisma/client";
import { Plus, Waypoints } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { AutoRenewBadge } from "@/components/domain/renewal-list";
import { ListToolbar } from "@/components/domain/list-toolbar";
import { filterOptions } from "@/lib/filters";
import { EnumBadge } from "@/components/ui/badge";
import { ButtonLink } from "@/components/ui/button";
import { PageHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { StatTile } from "@/components/ui/stat-tile";
import {
  NumCell,
  NumHead,
  PrimaryCell,
  ResultCount,
  Table,
  TableScroll,
  TableWrapper,
} from "@/components/ui/table";
import {
  BILLING_CYCLE,
  SUBSCRIPTION_STATUS,
  SUBSCRIPTION_STATUS_LIST,
  SUBSCRIPTION_TYPE,
  SUBSCRIPTION_TYPE_LIST,
} from "@/lib/constants";
import { refreshOverdueStatuses } from "@/lib/actions/subscriptions";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { countdownLabel, daysUntil, toMonthly, toYearly, urgencyFor } from "@/lib/renewals";

export const metadata: Metadata = { title: "Abonnements" };
export const dynamic = "force-dynamic";

export default async function SubscriptionsPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; statut?: string; type?: string; echeance?: string }>;
}) {
  // Les statuts d'impayé sont recalculés à l'ouverture : ce que l'écran montre
  // est vrai à l'instant où on le regarde.
  await refreshOverdueStatuses();

  const params = await searchParams;
  const query = params.q?.trim();

  const horizon = new Date();
  if (params.echeance) horizon.setDate(horizon.getDate() + Number(params.echeance));

  const where: Prisma.SubscriptionWhereInput = {
    ...(params.statut && params.statut in SUBSCRIPTION_STATUS
      ? { status: params.statut as SubscriptionStatus }
      : {}),
    ...(params.type && params.type in SUBSCRIPTION_TYPE
      ? { type: params.type as SubscriptionType }
      : {}),
    ...(params.echeance ? { nextRenewalAt: { lte: horizon } } : {}),
    ...(query
      ? {
          OR: [{ name: { contains: query } }, { client: { company: { contains: query } } }],
        }
      : {}),
  };

  const subscriptions = await db.subscription.findMany({
    where,
    orderBy: [{ nextRenewalAt: "asc" }],
    include: {
      client: { select: { id: true, company: true } },
      _count: { select: { periods: true } },
      periods: {
        where: { status: { in: ["DUE", "OVERDUE"] } },
        select: { id: true },
      },
    },
  });

  const active = subscriptions.filter((subscription) => subscription.status === "ACTIVE");
  const mrr = active.reduce(
    (sum, subscription) => sum + toMonthly(toNumber(subscription.amount), subscription.billingCycle),
    0,
  );
  const arr = active.reduce(
    (sum, subscription) => sum + toYearly(toNumber(subscription.amount), subscription.billingCycle),
    0,
  );
  const withUnpaid = subscriptions.filter((subscription) => subscription.periods.length > 0).length;

  return (
    <div className="space-y-4">
      <PageHeader
        title="Abonnements"
        description="Les contrats récurrents qui financent l'hébergement, le domaine et la maintenance."
        actions={
          <ButtonLink href="/abonnements/nouveau" variant="primary">
            <Plus className="h-3.5 w-3.5" />
            Nouvel abonnement
          </ButtonLink>
        }
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile label="Abonnements actifs" value={active.length} />
        <StatTile label="Revenus annuels" value={formatMoney(arr)} tone="success" />
        <StatTile label="Revenus mensuels estimés" value={formatMoney(mrr)} />
        <StatTile
          label="Avec impayé"
          value={withUnpaid}
          tone={withUnpaid > 0 ? "danger" : "neutral"}
          href="/paiements"
        />
      </div>

      <ListToolbar
        searchPlaceholder="Intitulé, client…"
        filters={[
          { name: "statut", label: "Statut", options: filterOptions(SUBSCRIPTION_STATUS_LIST) },
          { name: "type", label: "Type", options: filterOptions(SUBSCRIPTION_TYPE_LIST) },
          {
            name: "echeance",
            label: "Renouvellement dans",
            options: [
              { value: "7", label: "7 jours" },
              { value: "30", label: "30 jours" },
              { value: "90", label: "90 jours" },
            ],
          },
        ]}
      >
        <ResultCount count={subscriptions.length} singular="abonnement" plural="abonnements" />
      </ListToolbar>

      {subscriptions.length === 0 ? (
        <TableWrapper>
          <EmptyState
            icon={<Waypoints className="h-4 w-4" />}
            title="Aucun abonnement"
            description="C'est ici que se suit le revenu récurrent : qui paie quoi, pour quelle période."
            action={
              <ButtonLink href="/abonnements/nouveau" variant="primary">
                Créer un abonnement
              </ButtonLink>
            }
          />
        </TableWrapper>
      ) : (
        <TableWrapper>
          <TableScroll>
            <Table>
              <thead>
                <tr>
                  <th>Abonnement</th>
                  <th>Client</th>
                  <th>Statut</th>
                  <th>Type</th>
                  <th>Période couverte</th>
                  <th>Renouvellement</th>
                  <th>Reconduction</th>
                  <NumHead>Montant</NumHead>
                  <NumHead>Annualisé</NumHead>
                </tr>
              </thead>
              <tbody>
                {subscriptions.map((subscription) => {
                  const urgency = urgencyFor(subscription.nextRenewalAt);
                  const days = daysUntil(subscription.nextRenewalAt);

                  return (
                    <tr key={subscription.id}>
                      <td>
                        <Link href={`/abonnements/${subscription.id}`} className="block">
                          <PrimaryCell
                            title={subscription.name}
                            subtitle={
                              subscription.periods.length > 0
                                ? `${subscription.periods.length} échéance(s) impayée(s)`
                                : undefined
                            }
                          />
                        </Link>
                      </td>
                      <td>
                        <Link
                          href={`/clients/${subscription.client.id}`}
                          className="text-ink-secondary hover:text-accent-text"
                        >
                          {subscription.client.company}
                        </Link>
                      </td>
                      <td>
                        <EnumBadge meta={SUBSCRIPTION_STATUS[subscription.status]} />
                      </td>
                      <td className="text-ink-secondary">
                        {SUBSCRIPTION_TYPE[subscription.type].label}
                      </td>
                      <td className="whitespace-nowrap text-xs text-ink-secondary">
                        {formatDate(subscription.currentPeriodStart, "short")}
                        <span className="mx-1 text-ink-muted">→</span>
                        {formatDate(subscription.currentPeriodEnd, "short")}
                      </td>
                      <td>
                        <div className="text-ink">
                          {formatDate(subscription.nextRenewalAt, "short")}
                        </div>
                        <div
                          data-tone={urgency.tone}
                          className="text-[11px] text-[color:var(--tone-fg)]"
                        >
                          {countdownLabel(days)}
                        </div>
                      </td>
                      <td>
                        <AutoRenewBadge autoRenew={subscription.autoRenew} />
                      </td>
                      <NumCell className="text-ink-secondary">
                        {formatMoney(subscription.amount, { currency: subscription.currency })}
                        <span className="ml-1 text-[11px] text-ink-muted">
                          /{BILLING_CYCLE[subscription.billingCycle].label.toLowerCase()}
                        </span>
                      </NumCell>
                      <NumCell className="font-medium text-ink">
                        {formatMoney(
                          toYearly(toNumber(subscription.amount), subscription.billingCycle),
                          { currency: subscription.currency },
                        )}
                      </NumCell>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={8} className="text-xs font-medium text-ink-muted">
                    Revenus annuels récurrents (abonnements actifs)
                  </td>
                  <NumCell className="font-semibold text-ink">{formatMoney(arr)}</NumCell>
                </tr>
              </tfoot>
            </Table>
          </TableScroll>
        </TableWrapper>
      )}
    </div>
  );
}
