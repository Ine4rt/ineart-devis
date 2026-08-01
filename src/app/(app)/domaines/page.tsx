import type { Prisma } from "@prisma/client";
import { Globe, Plus } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { AutoRenewBadge } from "@/components/domain/renewal-list";
import { ListToolbar } from "@/components/domain/list-toolbar";
import { Badge } from "@/components/ui/badge";
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
import { COMMON_REGISTRARS } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { countdownLabel, daysUntil, urgencyFor } from "@/lib/renewals";

export const metadata: Metadata = { title: "Domaines" };
export const dynamic = "force-dynamic";

export default async function DomainsPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; registrar?: string; echeance?: string; auto?: string }>;
}) {
  const params = await searchParams;
  const query = params.q?.trim();

  const horizon = new Date();
  if (params.echeance) horizon.setDate(horizon.getDate() + Number(params.echeance));

  const where: Prisma.DomainWhereInput = {
    ...(params.registrar ? { registrar: params.registrar } : {}),
    ...(params.auto === "manuel" ? { autoRenew: false } : {}),
    ...(params.echeance ? { expiresAt: { not: null, lte: horizon } } : {}),
    ...(query
      ? {
          OR: [
            { name: { contains: query } },
            { registrar: { contains: query } },
            { client: { company: { contains: query } } },
          ],
        }
      : {}),
  };

  const domains = await db.domain.findMany({
    where,
    orderBy: [{ expiresAt: "asc" }],
    include: {
      client: { select: { id: true, company: true } },
      _count: { select: { subdomains: true } },
    },
  });

  const annualCost = domains.reduce((sum, domain) => sum + toNumber(domain.renewalPrice), 0);
  const expiringSoon = domains.filter((domain) => {
    const days = daysUntil(domain.expiresAt);
    return Number.isFinite(days) && days <= 30;
  }).length;
  const manualRenewal = domains.filter((domain) => !domain.autoRenew).length;

  return (
    <div className="space-y-4">
      <PageHeader
        title="Domaines"
        description="Registrars, expirations, DNS et certificats — tout ce qui doit rester valide."
        actions={
          <ButtonLink href="/domaines/nouveau" variant="primary">
            <Plus className="h-3.5 w-3.5" />
            Nouveau domaine
          </ButtonLink>
        }
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile label="Domaines suivis" value={domains.length} />
        <StatTile
          label="Coût annuel"
          value={formatMoney(annualCost)}
          hint="Ce que vous payez aux registrars"
        />
        <StatTile
          label="Expirent sous 30 jours"
          value={expiringSoon}
          tone={expiringSoon > 0 ? "warning" : "neutral"}
        />
        <StatTile
          label="Sans reconduction auto"
          value={manualRenewal}
          tone={manualRenewal > 0 ? "caution" : "neutral"}
          hint="Nécessitent une action manuelle"
        />
      </div>

      <ListToolbar
        searchPlaceholder="Domaine, registrar, client…"
        filters={[
          {
            name: "registrar",
            label: "Registrar",
            options: COMMON_REGISTRARS.map((value) => ({ value, label: value })),
          },
          {
            name: "echeance",
            label: "Expire dans",
            options: [
              { value: "7", label: "7 jours" },
              { value: "30", label: "30 jours" },
              { value: "90", label: "90 jours" },
            ],
          },
          {
            name: "auto",
            label: "Reconduction",
            options: [{ value: "manuel", label: "Manuelle uniquement" }],
          },
        ]}
      >
        <ResultCount count={domains.length} singular="domaine" plural="domaines" />
      </ListToolbar>

      {domains.length === 0 ? (
        <TableWrapper>
          <EmptyState
            icon={<Globe className="h-4 w-4" />}
            title="Aucun domaine"
            description="Enregistrez vos noms de domaine pour être alerté avant chaque expiration."
            action={
              <ButtonLink href="/domaines/nouveau" variant="primary">
                Ajouter un domaine
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
                  <th>Domaine</th>
                  <th>Client</th>
                  <th>Registrar</th>
                  <th>Expiration</th>
                  <th>Reconduction</th>
                  <th>SSL</th>
                  <NumHead>Prix / an</NumHead>
                </tr>
              </thead>
              <tbody>
                {domains.map((domain) => {
                  const urgency = urgencyFor(domain.expiresAt);
                  const days = daysUntil(domain.expiresAt);
                  const sslDays = daysUntil(domain.sslExpiresAt);

                  return (
                    <tr key={domain.id}>
                      <td>
                        <Link href={`/domaines/${domain.id}`} className="block">
                          <PrimaryCell
                            title={domain.name}
                            subtitle={
                              domain._count.subdomains > 0
                                ? `${domain._count.subdomains} sous-domaine(s)`
                                : undefined
                            }
                          />
                        </Link>
                      </td>
                      <td>
                        {domain.client ? (
                          <Link
                            href={`/clients/${domain.client.id}`}
                            className="text-ink-secondary hover:text-accent-text"
                          >
                            {domain.client.company}
                          </Link>
                        ) : (
                          <span className="text-ink-muted">Non rattaché</span>
                        )}
                      </td>
                      <td className="text-ink-secondary">{domain.registrar ?? "—"}</td>
                      <td>
                        {domain.expiresAt ? (
                          <div>
                            <div className="text-ink">{formatDate(domain.expiresAt, "short")}</div>
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
                        <AutoRenewBadge autoRenew={domain.autoRenew} />
                      </td>
                      <td>
                        {domain.sslExpiresAt ? (
                          <Badge tone={sslDays < 15 ? "danger" : sslDays < 30 ? "warning" : "success"}>
                            {countdownLabel(sslDays)}
                          </Badge>
                        ) : (
                          <span className="text-ink-muted">—</span>
                        )}
                      </td>
                      <NumCell className="text-ink-secondary">
                        {formatMoney(domain.renewalPrice)}
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
