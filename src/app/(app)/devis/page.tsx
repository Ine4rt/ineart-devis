import type { Prisma, QuoteStatus } from "@prisma/client";
import { FileText, Plus } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

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
import { QUOTE_STATUS, QUOTE_STATUS_LIST } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { computeQuoteTotals } from "@/lib/quotes";
import { daysUntil } from "@/lib/renewals";

export const metadata: Metadata = { title: "Devis" };
export const dynamic = "force-dynamic";

export default async function QuotesPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; statut?: string }>;
}) {
  const params = await searchParams;
  const query = params.q?.trim();

  const where: Prisma.QuoteWhereInput = {
    ...(params.statut && params.statut in QUOTE_STATUS
      ? { status: params.statut as QuoteStatus }
      : {}),
    ...(query
      ? {
          OR: [
            { number: { contains: query } },
            { title: { contains: query } },
            { client: { company: { contains: query } } },
          ],
        }
      : {}),
  };

  const quotes = await db.quote.findMany({
    where,
    orderBy: { issuedAt: "desc" },
    include: {
      client: { select: { id: true, company: true } },
      items: { select: { quantity: true, unitPrice: true } },
    },
  });

  const rows = quotes.map((quote) => ({
    ...quote,
    totals: computeQuoteTotals(
      quote.items.map((item) => ({
        quantity: toNumber(item.quantity),
        unitPrice: toNumber(item.unitPrice),
      })),
      toNumber(quote.discountPct),
      toNumber(quote.taxRate),
    ),
  }));

  const sent = rows.filter((row) => row.status === "SENT");
  const accepted = rows.filter((row) => row.status === "ACCEPTED");
  const decided = rows.filter((row) => ["ACCEPTED", "REJECTED"].includes(row.status));

  const pendingValue = sent.reduce((sum, row) => sum + row.totals.total, 0);
  const wonValue = accepted.reduce((sum, row) => sum + row.totals.total, 0);
  const winRate =
    decided.length > 0 ? Math.round((accepted.length / decided.length) * 100) : null;

  return (
    <div className="space-y-4">
      <PageHeader
        title="Devis"
        description="Vos propositions commerciales, de l'envoi à la signature."
        actions={
          <ButtonLink href="/devis/nouveau" variant="primary">
            <Plus className="h-3.5 w-3.5" />
            Nouveau devis
          </ButtonLink>
        }
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile label="Devis émis" value={rows.length} />
        <StatTile
          label="En attente de réponse"
          value={formatMoney(pendingValue)}
          hint={`${sent.length} devis envoyé(s)`}
          tone="info"
        />
        <StatTile
          label="Signés"
          value={formatMoney(wonValue)}
          hint={`${accepted.length} devis accepté(s)`}
          tone="success"
        />
        <StatTile
          label="Taux d'acceptation"
          value={winRate !== null ? `${winRate}%` : "—"}
          hint={winRate !== null ? `sur ${decided.length} devis tranchés` : "Aucun devis tranché"}
        />
      </div>

      <ListToolbar
        searchPlaceholder="Numéro, objet, client…"
        filters={[{ name: "statut", label: "Statut", options: filterOptions(QUOTE_STATUS_LIST) }]}
      >
        <ResultCount count={rows.length} singular="devis" plural="devis" />
      </ListToolbar>

      {rows.length === 0 ? (
        <TableWrapper>
          <EmptyState
            icon={<FileText className="h-4 w-4" />}
            title="Aucun devis"
            description="Créez une proposition chiffrée : prestation de création, puis abonnement annuel."
            action={
              <ButtonLink href="/devis/nouveau" variant="primary">
                Créer un devis
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
                  <th>Devis</th>
                  <th>Client</th>
                  <th>Statut</th>
                  <th>Émis le</th>
                  <th>Validité</th>
                  <NumHead>Total HT</NumHead>
                  <NumHead>Total TTC</NumHead>
                </tr>
              </thead>
              <tbody>
                {rows.map((quote) => {
                  const validityDays = quote.validUntil ? daysUntil(quote.validUntil) : null;
                  const expiringSoon =
                    quote.status === "SENT" && validityDays !== null && validityDays <= 7;

                  return (
                    <tr key={quote.id}>
                      <td>
                        <Link href={`/devis/${quote.id}`} className="block">
                          <PrimaryCell
                            title={quote.number}
                            subtitle={quote.title}
                          />
                        </Link>
                      </td>
                      <td>
                        <Link
                          href={`/clients/${quote.client.id}`}
                          className="text-ink-secondary hover:text-accent-text"
                        >
                          {quote.client.company}
                        </Link>
                      </td>
                      <td>
                        <EnumBadge meta={QUOTE_STATUS[quote.status]} />
                      </td>
                      <td className="whitespace-nowrap text-ink-muted">
                        {formatDate(quote.issuedAt, "short")}
                      </td>
                      <td>
                        {quote.validUntil ? (
                          <span
                            data-tone={expiringSoon ? "warning" : "muted"}
                            className="whitespace-nowrap text-[color:var(--tone-fg)]"
                          >
                            {formatDate(quote.validUntil, "short")}
                          </span>
                        ) : (
                          <span className="text-ink-muted">—</span>
                        )}
                      </td>
                      <NumCell className="text-ink-secondary">
                        {formatMoney(quote.totals.taxable, { currency: quote.currency })}
                      </NumCell>
                      <NumCell className="font-medium text-ink">
                        {formatMoney(quote.totals.total, { currency: quote.currency })}
                      </NumCell>
                    </tr>
                  );
                })}
              </tbody>
            </Table>
          </TableScroll>
        </TableWrapper>
      )}
    </div>
  );
}
