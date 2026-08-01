import { TrendingUp, Wallet } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { AreaChart, BarChart, DonutChart } from "@/components/charts/chart-kit";
import { Badge } from "@/components/ui/badge";
import { Card, CardBody, CardHeader, PageHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { StatTile } from "@/components/ui/stat-tile";
import {
  NumCell,
  NumHead,
  PrimaryCell,
  Table,
  TableScroll,
  TableWrapper,
} from "@/components/ui/table";
import { formatMoney, formatPercent, toNumber } from "@/lib/format";
import { getDashboardData } from "@/lib/queries/dashboard";
import { db } from "@/lib/db";
import { toYearly } from "@/lib/renewals";
import { cn } from "@/lib/utils";

export const metadata: Metadata = { title: "Revenus" };
export const dynamic = "force-dynamic";

/**
 * Analyse financière.
 *
 * L'apport principal par rapport à l'écran « Abonnements » : la **marge réelle
 * par client**. Un contrat à 480 €/an dont le domaine et l'hébergement coûtent
 * 180 € ne rapporte pas 480 € — et c'est cette différence qui doit guider les
 * décisions de tarification.
 */
export default async function RevenuePage() {
  const [data, clients] = await Promise.all([
    getDashboardData(),
    db.client.findMany({
      where: { subscriptions: { some: {} } },
      include: {
        subscriptions: {
          where: { status: { in: ["ACTIVE", "PAST_DUE"] } },
          select: { amount: true, billingCycle: true, currency: true },
        },
        domains: { select: { renewalPrice: true } },
        hostings: { select: { price: true, billingCycle: true } },
      },
    }),
  ]);

  const profitability = clients
    .map((client) => {
      const revenue = client.subscriptions.reduce(
        (sum, subscription) =>
          sum + toYearly(toNumber(subscription.amount), subscription.billingCycle),
        0,
      );
      const domainCost = client.domains.reduce(
        (sum, domain) => sum + toNumber(domain.renewalPrice),
        0,
      );
      const hostingCost = client.hostings.reduce(
        (sum, hosting) => sum + toYearly(toNumber(hosting.price), hosting.billingCycle),
        0,
      );
      const cost = domainCost + hostingCost;

      return {
        id: client.id,
        company: client.company,
        revenue,
        cost,
        margin: revenue - cost,
        marginRate: revenue > 0 ? (revenue - cost) / revenue : 0,
      };
    })
    .filter((row) => row.revenue > 0 || row.cost > 0)
    .sort((a, b) => b.margin - a.margin);

  const totalRevenue = profitability.reduce((sum, row) => sum + row.revenue, 0);
  const totalCost = profitability.reduce((sum, row) => sum + row.cost, 0);
  const averageContract =
    data.billing.activeSubscriptions > 0 ? data.revenue.arr / data.billing.activeSubscriptions : 0;

  // Concentration du chiffre d'affaires : un indicateur de risque souvent ignoré.
  const topShare =
    totalRevenue > 0
      ? [...profitability].sort((a, b) => b.revenue - a.revenue)[0].revenue / totalRevenue
      : 0;

  return (
    <div className="space-y-4">
      <PageHeader
        title="Revenus & rentabilité"
        description="Ce que vous facturez, ce que vous dépensez, et ce qui reste réellement."
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile
          label="Revenus annuels récurrents"
          value={formatMoney(data.revenue.arr)}
          hint={`${formatMoney(data.revenue.mrr)} / mois`}
          icon={<TrendingUp className="h-4 w-4" />}
          tone="success"
        />
        <StatTile
          label="Coûts d'infrastructure"
          value={formatMoney(data.infrastructure.annualCost)}
          hint={`${data.infrastructure.domains} domaine(s) · ${data.infrastructure.hostings} hébergement(s)`}
          tone="warning"
        />
        <StatTile
          label="Marge nette annuelle"
          value={formatMoney(data.revenue.margin)}
          hint={`${formatPercent(data.revenue.marginRate)} des revenus`}
          icon={<Wallet className="h-4 w-4" />}
          tone="accent"
        />
        <StatTile
          label="Contrat moyen"
          value={formatMoney(averageContract)}
          hint={
            topShare > 0.3
              ? `⚠ ${formatPercent(topShare)} du CA sur un seul client`
              : "Répartition saine du chiffre d'affaires"
          }
          tone={topShare > 0.3 ? "caution" : "neutral"}
        />
      </div>

      <div className="grid gap-4 lg:grid-cols-3">
        <Card className="lg:col-span-2">
          <CardHeader
            title="Évolution des revenus récurrents"
            description="ARR reconstitué mois par mois depuis les dates de souscription"
          />
          <CardBody>
            {data.revenue.arr > 0 ? (
              <AreaChart
                data={data.revenue.trend}
                format="money"
                height={200}
              />
            ) : (
              <EmptyState compact title="Aucun revenu récurrent" />
            )}
          </CardBody>
        </Card>

        <Card>
          <CardHeader title="Répartition par type" />
          <CardBody>
            {data.revenue.mix.length > 0 ? (
              <DonutChart
                data={data.revenue.mix}
                format="money-compact"
                centerLabel="ARR"
                centerValue={formatMoney(data.revenue.arr, { compact: true })}
                size={130}
              />
            ) : (
              <EmptyState compact title="Pas de données" />
            )}
          </CardBody>
        </Card>
      </div>

      <Card>
        <CardHeader
          title="Encaissements par mois"
          description={`${formatMoney(data.revenue.collectedThisYear)} encaissés cette année`}
        />
        <CardBody>
          <BarChart
            data={data.revenue.monthlyCollected}
            format="money"
            highlightIndex={data.now.getMonth()}
            height={170}
          />
        </CardBody>
      </Card>

      <Card className="overflow-hidden">
        <CardHeader
          title="Rentabilité par client"
          description="Revenus facturés moins les coûts de domaine et d'hébergement réellement payés."
        />
        {profitability.length === 0 ? (
          <EmptyState
            compact
            title="Aucune donnée de rentabilité"
            description="Renseignez les prix de vos domaines et hébergements pour voir apparaître vos marges."
          />
        ) : (
          <TableScroll>
            <Table>
              <thead>
                <tr>
                  <th>Client</th>
                  <NumHead>Revenus / an</NumHead>
                  <NumHead>Coûts / an</NumHead>
                  <NumHead>Marge</NumHead>
                  <th className="w-32">Taux de marge</th>
                </tr>
              </thead>
              <tbody>
                {profitability.map((row) => (
                  <tr key={row.id}>
                    <td>
                      <Link href={`/clients/${row.id}`} className="block">
                        <PrimaryCell title={row.company} />
                      </Link>
                    </td>
                    <NumCell className="text-ink-secondary">{formatMoney(row.revenue)}</NumCell>
                    <NumCell className="text-ink-muted">−{formatMoney(row.cost)}</NumCell>
                    <NumCell
                      className={cn(
                        "font-semibold",
                        row.margin >= 0 ? "text-ink" : "text-[var(--tone-danger)]",
                      )}
                    >
                      {formatMoney(row.margin)}
                    </NumCell>
                    <td>
                      <Badge
                        tone={
                          row.marginRate >= 0.6
                            ? "success"
                            : row.marginRate >= 0.3
                              ? "caution"
                              : "danger"
                        }
                      >
                        {formatPercent(row.marginRate)}
                      </Badge>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td className="text-xs font-medium text-ink-muted">Total</td>
                  <NumCell className="font-semibold text-ink">{formatMoney(totalRevenue)}</NumCell>
                  <NumCell className="font-semibold text-ink-muted">
                    −{formatMoney(totalCost)}
                  </NumCell>
                  <NumCell className="font-semibold text-ink">
                    {formatMoney(totalRevenue - totalCost)}
                  </NumCell>
                  <td className="text-xs text-ink-muted">
                    {totalRevenue > 0
                      ? formatPercent((totalRevenue - totalCost) / totalRevenue)
                      : "—"}
                  </td>
                </tr>
              </tfoot>
            </Table>
          </TableScroll>
        )}
      </Card>
    </div>
  );
}
