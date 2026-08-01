import type { ClientStatus, Prisma } from "@prisma/client";
import { Building2, Plus } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { ListToolbar } from "@/components/domain/list-toolbar";
import { filterOptions } from "@/lib/filters";
import { EnumBadge } from "@/components/ui/badge";
import { ButtonLink } from "@/components/ui/button";
import { PageHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { Avatar } from "@/components/ui/avatar";
import {
  NumCell,
  NumHead,
  PrimaryCell,
  ResultCount,
  Table,
  TableScroll,
  TableWrapper,
} from "@/components/ui/table";
import { CLIENT_STATUS, CLIENT_STATUS_LIST } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { toYearly } from "@/lib/renewals";

export const metadata: Metadata = { title: "Clients" };
export const dynamic = "force-dynamic";

const SORTS = {
  recent: { createdAt: "desc" },
  nom: { company: "asc" },
} as const;

export default async function ClientsPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; statut?: string; tri?: string }>;
}) {
  const params = await searchParams;
  const query = params.q?.trim();
  const status = params.statut as ClientStatus | undefined;
  const sort = params.tri === "nom" ? "nom" : "recent";

  const where: Prisma.ClientWhereInput = {
    ...(status && status in CLIENT_STATUS ? { status } : {}),
    ...(query
      ? {
          OR: [
            { company: { contains: query } },
            { firstName: { contains: query } },
            { lastName: { contains: query } },
            { email: { contains: query } },
            { city: { contains: query } },
            { vatNumber: { contains: query } },
            { reference: { contains: query } },
          ],
        }
      : {}),
  };

  const clients = await db.client.findMany({
    where,
    orderBy: SORTS[sort],
    include: {
      _count: { select: { projects: true, domains: true } },
      subscriptions: {
        where: { status: "ACTIVE" },
        select: { amount: true, billingCycle: true, currency: true },
      },
    },
  });

  // Le revenu annuel par client se calcule ici plutôt qu'en base : les cycles
  // de facturation sont hétérogènes et la conversion vit dans lib/renewals.
  const rows = clients.map((client) => ({
    ...client,
    arr: client.subscriptions.reduce(
      (sum, subscription) => sum + toYearly(toNumber(subscription.amount), subscription.billingCycle),
      0,
    ),
  }));

  const totalArr = rows.reduce((sum, row) => sum + row.arr, 0);

  return (
    <div className="space-y-4">
      <PageHeader
        title="Clients"
        description="Toutes vos entreprises, prospects compris. Cliquez sur une ligne pour ouvrir la fiche complète."
        actions={
          <ButtonLink href="/clients/nouveau" variant="primary">
            <Plus className="h-3.5 w-3.5" />
            Nouveau client
          </ButtonLink>
        }
      />

      <ListToolbar
        searchPlaceholder="Société, contact, e-mail, ville, TVA…"
        filters={[
          { name: "statut", label: "Statut", options: filterOptions(CLIENT_STATUS_LIST) },
          {
            name: "tri",
            label: "Tri",
            options: [
              { value: "recent", label: "Plus récents" },
              { value: "nom", label: "Nom (A→Z)" },
            ],
          },
        ]}
      >
        <ResultCount count={rows.length} singular="client" plural="clients" />
      </ListToolbar>

      {rows.length === 0 ? (
        <TableWrapper>
          <EmptyState
            icon={<Building2 className="h-4 w-4" />}
            title={query || status ? "Aucun résultat" : "Aucun client pour l'instant"}
            description={
              query || status
                ? "Ajustez votre recherche ou réinitialisez les filtres."
                : "Créez votre première fiche pour commencer à suivre projets, domaines et abonnements."
            }
            action={
              !query && !status ? (
                <ButtonLink href="/clients/nouveau" variant="primary">
                  Créer un client
                </ButtonLink>
              ) : null
            }
          />
        </TableWrapper>
      ) : (
        <TableWrapper>
          <TableScroll>
            <Table>
              <thead>
                <tr>
                  <th>Société</th>
                  <th>Statut</th>
                  <th>Localité</th>
                  <NumHead>Projets</NumHead>
                  <NumHead>Domaines</NumHead>
                  <NumHead>Revenu annuel</NumHead>
                  <th>Créé le</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((client) => (
                  <tr key={client.id} className="cursor-pointer">
                    <td>
                      <Link href={`/clients/${client.id}`} className="block">
                        <PrimaryCell
                          leading={
                            <Avatar name={client.company} src={client.logoUrl} size="sm" />
                          }
                          title={client.company}
                          subtitle={
                            [client.firstName, client.lastName].filter(Boolean).join(" ") ||
                            client.email ||
                            client.reference
                          }
                        />
                      </Link>
                    </td>
                    <td>
                      <EnumBadge meta={CLIENT_STATUS[client.status]} />
                    </td>
                    <td className="text-ink-secondary">{client.city ?? "—"}</td>
                    <NumCell className="text-ink-secondary">{client._count.projects}</NumCell>
                    <NumCell className="text-ink-secondary">{client._count.domains}</NumCell>
                    <NumCell className={client.arr > 0 ? "font-medium text-ink" : "text-ink-muted"}>
                      {client.arr > 0 ? formatMoney(client.arr) : "—"}
                    </NumCell>
                    <td className="whitespace-nowrap text-ink-muted">
                      {formatDate(client.createdAt, "short")}
                    </td>
                  </tr>
                ))}
              </tbody>
              {totalArr > 0 ? (
                <tfoot>
                  <tr>
                    <td colSpan={5} className="text-xs font-medium text-ink-muted">
                      Total sur la sélection
                    </td>
                    <NumCell className="font-semibold text-ink">{formatMoney(totalArr)}</NumCell>
                    <td />
                  </tr>
                </tfoot>
              ) : null}
            </Table>
          </TableScroll>
        </TableWrapper>
      )}
    </div>
  );
}
