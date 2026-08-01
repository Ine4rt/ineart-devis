import type { Prisma } from "@prisma/client";
import { Plus, Server } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { AutoRenewBadge } from "@/components/domain/renewal-list";
import { ListToolbar } from "@/components/domain/list-toolbar";
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
import { BILLING_CYCLE, COMMON_HOSTS } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { countdownLabel, daysUntil, toYearly, urgencyFor } from "@/lib/renewals";

export const metadata: Metadata = { title: "Hébergements" };
export const dynamic = "force-dynamic";

export default async function HostingsPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; fournisseur?: string; auto?: string }>;
}) {
  const params = await searchParams;
  const query = params.q?.trim();

  const where: Prisma.HostingWhereInput = {
    ...(params.fournisseur ? { provider: params.fournisseur } : {}),
    ...(params.auto === "manuel" ? { autoRenew: false } : {}),
    ...(query
      ? {
          OR: [
            { provider: { contains: query } },
            { plan: { contains: query } },
            { serverIp: { contains: query } },
            { client: { company: { contains: query } } },
          ],
        }
      : {}),
  };

  const hostings = await db.hosting.findMany({
    where,
    orderBy: [{ renewsAt: "asc" }],
    include: { client: { select: { id: true, company: true } } },
  });

  const annualCost = hostings.reduce(
    (sum, hosting) => sum + toYearly(toNumber(hosting.price), hosting.billingCycle),
    0,
  );
  const expiringSoon = hostings.filter((hosting) => {
    const days = daysUntil(hosting.renewsAt);
    return Number.isFinite(days) && days <= 30;
  }).length;

  return (
    <div className="space-y-4">
      <PageHeader
        title="Hébergements"
        description="Fournisseurs, offres, échéances et informations techniques."
        actions={
          <ButtonLink href="/hebergements/nouveau" variant="primary">
            <Plus className="h-3.5 w-3.5" />
            Nouvel hébergement
          </ButtonLink>
        }
      />

      <div className="grid gap-3 sm:grid-cols-3">
        <StatTile label="Hébergements" value={hostings.length} />
        <StatTile
          label="Coût annuel"
          value={formatMoney(annualCost)}
          hint="Tous cycles ramenés à l'année"
        />
        <StatTile
          label="Échéances sous 30 jours"
          value={expiringSoon}
          tone={expiringSoon > 0 ? "warning" : "neutral"}
        />
      </div>

      <ListToolbar
        searchPlaceholder="Fournisseur, offre, IP, client…"
        filters={[
          {
            name: "fournisseur",
            label: "Fournisseur",
            options: COMMON_HOSTS.map((value) => ({ value, label: value })),
          },
          {
            name: "auto",
            label: "Reconduction",
            options: [{ value: "manuel", label: "Manuelle uniquement" }],
          },
        ]}
      >
        <ResultCount count={hostings.length} singular="hébergement" plural="hébergements" />
      </ListToolbar>

      {hostings.length === 0 ? (
        <TableWrapper>
          <EmptyState
            icon={<Server className="h-4 w-4" />}
            title="Aucun hébergement"
            description="Ajoutez vos serveurs et formules pour suivre coûts et échéances."
            action={
              <ButtonLink href="/hebergements/nouveau" variant="primary">
                Ajouter un hébergement
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
                  <th>Hébergement</th>
                  <th>Client</th>
                  <th>Échéance</th>
                  <th>Reconduction</th>
                  <th>Disque</th>
                  <NumHead>Prix</NumHead>
                  <NumHead>Équivalent annuel</NumHead>
                </tr>
              </thead>
              <tbody>
                {hostings.map((hosting) => {
                  const urgency = urgencyFor(hosting.renewsAt);
                  const days = daysUntil(hosting.renewsAt);

                  return (
                    <tr key={hosting.id}>
                      <td>
                        <Link href={`/hebergements/${hosting.id}`} className="block">
                          <PrimaryCell
                            title={hosting.provider}
                            subtitle={hosting.plan ?? hosting.label ?? undefined}
                          />
                        </Link>
                      </td>
                      <td>
                        {hosting.client ? (
                          <Link
                            href={`/clients/${hosting.client.id}`}
                            className="text-ink-secondary hover:text-accent-text"
                          >
                            {hosting.client.company}
                          </Link>
                        ) : (
                          <span className="text-ink-muted">Non rattaché</span>
                        )}
                      </td>
                      <td>
                        {hosting.renewsAt ? (
                          <div>
                            <div className="text-ink">{formatDate(hosting.renewsAt, "short")}</div>
                            <div
                              data-tone={urgency.tone}
                              className="text-[11px] text-[color:var(--tone-fg)]"
                            >
                              {countdownLabel(days)}
                            </div>
                          </div>
                        ) : (
                          <span className="text-ink-muted">—</span>
                        )}
                      </td>
                      <td>
                        <AutoRenewBadge autoRenew={hosting.autoRenew} />
                      </td>
                      <td className="text-ink-secondary">
                        {hosting.diskSpaceGb ? `${hosting.diskSpaceGb} Go` : "—"}
                      </td>
                      <NumCell className="text-ink-secondary">
                        {formatMoney(hosting.price)}
                        <span className="ml-1 text-[11px] text-ink-muted">
                          /{BILLING_CYCLE[hosting.billingCycle].label.toLowerCase()}
                        </span>
                      </NumCell>
                      <NumCell className="font-medium text-ink">
                        {formatMoney(toYearly(toNumber(hosting.price), hosting.billingCycle))}
                      </NumCell>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={6} className="text-xs font-medium text-ink-muted">
                    Coût annuel total
                  </td>
                  <NumCell className="font-semibold text-ink">{formatMoney(annualCost)}</NumCell>
                </tr>
              </tfoot>
            </Table>
          </TableScroll>
        </TableWrapper>
      )}
    </div>
  );
}
