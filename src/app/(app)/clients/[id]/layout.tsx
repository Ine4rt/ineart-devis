import { ExternalLink, Mail, Pencil, Phone } from "lucide-react";
import { notFound } from "next/navigation";

import { EnumBadge } from "@/components/ui/badge";
import { ButtonLink } from "@/components/ui/button";
import { Avatar } from "@/components/ui/avatar";
import { CopyButton } from "@/components/ui/copy-button";
import { Tabs } from "@/components/ui/tabs";
import { CLIENT_STATUS, PIPELINE_STAGE } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatMoney, toNumber } from "@/lib/format";
import { toYearly } from "@/lib/renewals";

/**
 * En-tête persistant de la fiche client.
 *
 * Placé dans un layout : il reste affiché d'un onglet à l'autre et n'est pas
 * re-rendu à chaque navigation. Les onglets sont de vraies routes, donc chaque
 * vue d'un client a son URL.
 */
export default async function ClientLayout({
  children,
  params,
}: {
  children: React.ReactNode;
  params: Promise<{ id: string }>;
}) {
  const { id } = await params;

  const client = await db.client.findUnique({
    where: { id },
    include: {
      _count: {
        select: { projects: true, domains: true, hostings: true, documents: true, quotes: true },
      },
      subscriptions: {
        where: { status: "ACTIVE" },
        select: { amount: true, billingCycle: true },
      },
    },
  });

  if (!client) notFound();

  const arr = client.subscriptions.reduce(
    (sum, subscription) => sum + toYearly(toNumber(subscription.amount), subscription.billingCycle),
    0,
  );

  const fullName = [client.firstName, client.lastName].filter(Boolean).join(" ");
  const base = `/clients/${client.id}`;

  return (
    <div className="space-y-4">
      <header className="flex flex-wrap items-start justify-between gap-4">
        <div className="flex min-w-0 items-start gap-3.5">
          <Avatar name={client.company} src={client.logoUrl} size="xl" />
          <div className="min-w-0">
            <div className="flex flex-wrap items-center gap-2">
              <h1 className="truncate text-xl font-semibold text-ink">{client.company}</h1>
              <EnumBadge meta={CLIENT_STATUS[client.status]} />
              {client.status === "PROSPECT" && client.pipelineStage ? (
                <EnumBadge meta={PIPELINE_STAGE[client.pipelineStage]} dot={false} />
              ) : null}
            </div>

            <p className="mt-1 text-sm text-ink-secondary">
              {fullName || "Aucun interlocuteur renseigné"}
              <span className="mx-1.5 text-ink-muted">·</span>
              <span className="font-mono text-xs text-ink-muted">{client.reference}</span>
            </p>

            <div className="mt-2 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs">
              {client.email ? (
                <span className="flex items-center gap-1 text-ink-secondary">
                  <Mail className="h-3 w-3 text-ink-muted" />
                  <a href={`mailto:${client.email}`} className="link-subtle">
                    {client.email}
                  </a>
                  <CopyButton value={client.email} label="E-mail" silent />
                </span>
              ) : null}
              {client.phone ? (
                <span className="flex items-center gap-1 text-ink-secondary">
                  <Phone className="h-3 w-3 text-ink-muted" />
                  <a href={`tel:${client.phone}`} className="link-subtle">
                    {client.phone}
                  </a>
                  <CopyButton value={client.phone} label="Téléphone" silent />
                </span>
              ) : null}
              {client.website ? (
                <a
                  href={client.website.startsWith("http") ? client.website : `https://${client.website}`}
                  target="_blank"
                  rel="noreferrer"
                  className="flex items-center gap-1 text-ink-secondary link-subtle"
                >
                  <ExternalLink className="h-3 w-3 text-ink-muted" />
                  {client.website.replace(/^https?:\/\//, "")}
                </a>
              ) : null}
            </div>
          </div>
        </div>

        <div className="flex items-center gap-3">
          {arr > 0 ? (
            <div className="rounded-lg border border-line bg-surface px-3.5 py-2 text-right">
              <div className="text-[10px] uppercase tracking-wide text-ink-muted">Revenu annuel</div>
              <div className="text-base font-semibold text-ink tabular-nums">
                {formatMoney(arr)}
              </div>
            </div>
          ) : null}
          <ButtonLink href={`${base}/modifier`}>
            <Pencil className="h-3.5 w-3.5" />
            Modifier
          </ButtonLink>
        </div>
      </header>

      <Tabs
        items={[
          { href: base, label: "Vue d'ensemble", exact: true },
          { href: `${base}/projets`, label: "Projets", count: client._count.projects },
          {
            href: `${base}/technique`,
            label: "Technique",
            count: client._count.domains + client._count.hostings,
          },
          { href: `${base}/finances`, label: "Finances" },
          { href: `${base}/documents`, label: "Documents", count: client._count.documents },
          { href: `${base}/suivi`, label: "Suivi" },
        ]}
      />

      <div className="pt-1">{children}</div>
    </div>
  );
}

export async function generateMetadata({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const client = await db.client.findUnique({ where: { id }, select: { company: true } });
  return { title: client?.company ?? "Client" };
}

export const dynamic = "force-dynamic";
