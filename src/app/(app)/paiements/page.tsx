import type { PeriodStatus, Prisma } from "@prisma/client";
import { Banknote, Check } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { ListToolbar } from "@/components/domain/list-toolbar";
import { filterOptions } from "@/lib/filters";
import { EnumBadge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { PageHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { SegmentedLinks } from "@/components/ui/tabs";
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
import { PERIOD_STATUS, PERIOD_STATUS_LIST } from "@/lib/constants";
import { markPeriodPaid, refreshOverdueStatuses } from "@/lib/actions/subscriptions";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { countdownLabel, daysUntil } from "@/lib/renewals";

export const metadata: Metadata = { title: "Paiements" };
export const dynamic = "force-dynamic";

/**
 * Registre des encaissements.
 *
 * Répond en un écran aux trois questions du suivi financier : qui a payé, qui
 * doit payer, et pour quelle période. Les vues rapides en haut évitent d'avoir
 * à composer des filtres à la main pour les cas les plus fréquents.
 */
export default async function PaymentsPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; statut?: string; vue?: string }>;
}) {
  await refreshOverdueStatuses();

  const params = await searchParams;
  const query = params.q?.trim();
  const view = params.vue ?? "a-encaisser";

  const viewFilter: Prisma.SubscriptionPeriodWhereInput =
    view === "payes"
      ? { status: "PAID" }
      : view === "retard"
        ? { status: "OVERDUE" }
        : view === "tout"
          ? {}
          : { status: { in: ["DUE", "OVERDUE"] } };

  const where: Prisma.SubscriptionPeriodWhereInput = {
    ...viewFilter,
    ...(params.statut && params.statut in PERIOD_STATUS
      ? { status: params.statut as PeriodStatus }
      : {}),
    ...(query
      ? {
          OR: [
            { subscription: { name: { contains: query } } },
            { subscription: { client: { company: { contains: query } } } },
            { reference: { contains: query } },
          ],
        }
      : {}),
  };

  const [periods, totals] = await Promise.all([
    db.subscriptionPeriod.findMany({
      where,
      orderBy: [{ dueAt: "asc" }],
      include: {
        subscription: {
          select: {
            id: true,
            name: true,
            currency: true,
            client: { select: { id: true, company: true } },
          },
        },
      },
    }),
    db.subscriptionPeriod.groupBy({ by: ["status"], _sum: { amount: true }, _count: { _all: true } }),
  ]);

  const sumFor = (status: PeriodStatus) =>
    toNumber(totals.find((row) => row.status === status)?._sum.amount ?? 0);
  const countFor = (status: PeriodStatus) =>
    totals.find((row) => row.status === status)?._count._all ?? 0;

  const selectionTotal = periods.reduce((sum, period) => sum + toNumber(period.amount), 0);

  const views = [
    { href: "/paiements?vue=a-encaisser", label: "À encaisser", key: "a-encaisser" },
    { href: "/paiements?vue=retard", label: "En retard", key: "retard" },
    { href: "/paiements?vue=payes", label: "Payés", key: "payes" },
    { href: "/paiements?vue=tout", label: "Tout", key: "tout" },
  ];

  return (
    <div className="space-y-4">
      <PageHeader
        title="Paiements"
        description="Qui a payé, qui doit payer, et pour quelle période exactement."
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile
          label="En retard"
          value={formatMoney(sumFor("OVERDUE"))}
          hint={`${countFor("OVERDUE")} échéance(s)`}
          tone={countFor("OVERDUE") > 0 ? "danger" : "neutral"}
          href="/paiements?vue=retard"
        />
        <StatTile
          label="À échoir"
          value={formatMoney(sumFor("DUE"))}
          hint={`${countFor("DUE")} échéance(s)`}
          tone="warning"
        />
        <StatTile
          label="Encaissé (total)"
          value={formatMoney(sumFor("PAID"))}
          hint={`${countFor("PAID")} paiement(s)`}
          tone="success"
        />
        <StatTile
          label="Taux de recouvrement"
          value={
            sumFor("PAID") + sumFor("DUE") + sumFor("OVERDUE") > 0
              ? `${Math.round(
                  (sumFor("PAID") / (sumFor("PAID") + sumFor("DUE") + sumFor("OVERDUE"))) * 100,
                )}%`
              : "—"
          }
          hint="Part réglée sur l'ensemble facturé"
        />
      </div>

      <div className="flex flex-wrap items-center justify-between gap-3">
        <SegmentedLinks
          items={views.map((item) => ({
            href: item.href,
            label: item.label,
            active: view === item.key,
          }))}
        />
      </div>

      <ListToolbar
        searchPlaceholder="Client, abonnement, référence…"
        filters={[{ name: "statut", label: "Statut", options: filterOptions(PERIOD_STATUS_LIST) }]}
      >
        <ResultCount count={periods.length} singular="échéance" plural="échéances" />
      </ListToolbar>

      {periods.length === 0 ? (
        <TableWrapper>
          <EmptyState
            icon={<Banknote className="h-4 w-4" />}
            title={view === "a-encaisser" ? "Rien à encaisser" : "Aucune échéance"}
            description={
              view === "a-encaisser"
                ? "Tous vos clients sont à jour de leurs paiements."
                : "Aucune échéance ne correspond à cette vue."
            }
          />
        </TableWrapper>
      ) : (
        <TableWrapper>
          <TableScroll>
            <Table>
              <thead>
                <tr>
                  <th>Client</th>
                  <th>Abonnement</th>
                  <th>Période couverte</th>
                  <th>Échéance</th>
                  <th>Statut</th>
                  <th>Payé le</th>
                  <NumHead>Montant</NumHead>
                  <th />
                </tr>
              </thead>
              <tbody>
                {periods.map((period) => {
                  const markPaid = markPeriodPaid.bind(null, period.id);
                  const days = daysUntil(period.dueAt);
                  const isOpen = ["DUE", "OVERDUE"].includes(period.status);

                  return (
                    <tr key={period.id}>
                      <td>
                        <Link
                          href={`/clients/${period.subscription.client.id}`}
                          className="block"
                        >
                          <PrimaryCell title={period.subscription.client.company} />
                        </Link>
                      </td>
                      <td>
                        <Link
                          href={`/abonnements/${period.subscription.id}`}
                          className="text-ink-secondary hover:text-accent-text"
                        >
                          {period.subscription.name}
                        </Link>
                      </td>
                      <td className="whitespace-nowrap text-xs text-ink-secondary">
                        {formatDate(period.periodStart, "short")}
                        <span className="mx-1 text-ink-muted">→</span>
                        {formatDate(period.periodEnd, "short")}
                      </td>
                      <td>
                        <div className="whitespace-nowrap text-ink">
                          {formatDate(period.dueAt, "short")}
                        </div>
                        {isOpen ? (
                          <div
                            data-tone={days < 0 ? "danger" : days <= 7 ? "warning" : "muted"}
                            className="text-[11px] text-[color:var(--tone-fg)]"
                          >
                            {countdownLabel(days)}
                          </div>
                        ) : null}
                      </td>
                      <td>
                        <EnumBadge meta={PERIOD_STATUS[period.status]} />
                      </td>
                      <td className="whitespace-nowrap text-ink-muted">
                        {formatDate(period.paidAt, "short")}
                      </td>
                      <NumCell className="font-medium text-ink">
                        {formatMoney(period.amount, { currency: period.subscription.currency })}
                      </NumCell>
                      <td className="text-right">
                        {isOpen ? (
                          <form action={markPaid}>
                            <Button type="submit" size="sm" variant="default">
                              <Check className="h-3.5 w-3.5" />
                              Payé
                            </Button>
                          </form>
                        ) : null}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={6} className="text-xs font-medium text-ink-muted">
                    Total de la sélection
                  </td>
                  <NumCell className="font-semibold text-ink">
                    {formatMoney(selectionTotal)}
                  </NumCell>
                  <td />
                </tr>
              </tfoot>
            </Table>
          </TableScroll>
        </TableWrapper>
      )}
    </div>
  );
}
