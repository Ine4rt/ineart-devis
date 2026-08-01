import { Globe, Pencil, Plus, RefreshCw, ShieldCheck, Trash2 } from "lucide-react";
import Link from "next/link";
import { notFound } from "next/navigation";

import { AutoRenewBadge } from "@/components/domain/renewal-list";
import { CredentialVault } from "@/components/domain/credential-vault";
import { Badge } from "@/components/ui/badge";
import { Button, ButtonLink } from "@/components/ui/button";
import { Card, CardBody, CardHeader } from "@/components/ui/card";
import { DetailList, DetailRow } from "@/components/ui/detail-list";
import { EmptyState } from "@/components/ui/empty-state";
import { Field, Input } from "@/components/ui/form";
import { SubmitButton } from "@/components/ui/submit-button";
import { createSubdomain, deleteSubdomain, renewDomain } from "@/lib/actions/domains";
import { db } from "@/lib/db";
import { formatDate, formatMoney } from "@/lib/format";
import { countdownLabel, daysUntil, urgencyFor } from "@/lib/renewals";

export const dynamic = "force-dynamic";

export default async function DomainPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;

  const domain = await db.domain.findUnique({
    where: { id },
    include: {
      client: { select: { id: true, company: true } },
      project: { select: { id: true, name: true } },
      subdomains: { orderBy: { name: "asc" } },
      credentials: { orderBy: { createdAt: "desc" } },
    },
  });

  if (!domain) notFound();

  const urgency = urgencyFor(domain.expiresAt);
  const days = daysUntil(domain.expiresAt);
  const sslDays = daysUntil(domain.sslExpiresAt);

  const addSubdomain = createSubdomain.bind(null, domain.id);
  const renew = renewDomain.bind(null, domain.id);

  return (
    <div className="space-y-4">
      <header className="flex flex-wrap items-start justify-between gap-4">
        <div className="min-w-0">
          {domain.client ? (
            <p className="section-label">
              <Link href={`/clients/${domain.client.id}`} className="hover:text-accent-text">
                {domain.client.company}
              </Link>
            </p>
          ) : (
            <p className="section-label">Non rattaché</p>
          )}
          <div className="mt-1 flex flex-wrap items-center gap-2">
            <h1 className="font-mono text-xl font-semibold text-ink">{domain.name}</h1>
            <AutoRenewBadge autoRenew={domain.autoRenew} />
          </div>
        </div>

        <div className="flex items-center gap-2">
          <form action={renew}>
            <SubmitButton variant="default" pendingLabel="Renouvellement…">
              <RefreshCw className="h-3.5 w-3.5" />
              Marquer renouvelé (+1 an)
            </SubmitButton>
          </form>
          <ButtonLink href={`/domaines/${domain.id}/modifier`}>
            <Pencil className="h-3.5 w-3.5" />
            Modifier
          </ButtonLink>
        </div>
      </header>

      <div className="grid gap-3 sm:grid-cols-3">
        <div data-tone={urgency.tone} className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Expiration du domaine</p>
          <p className="mt-1.5 text-lg font-semibold text-ink tabular-nums">
            {formatDate(domain.expiresAt)}
          </p>
          <p className="mt-0.5 text-xs font-medium text-[color:var(--tone-fg)]">
            {domain.expiresAt ? countdownLabel(days) : "Date non renseignée"}
          </p>
        </div>

        <div
          data-tone={
            !domain.sslExpiresAt ? "muted" : sslDays < 15 ? "danger" : sslDays < 30 ? "warning" : "success"
          }
          className="card p-4"
        >
          <p className="flex items-center gap-1.5 text-xs font-medium text-ink-secondary">
            <ShieldCheck className="h-3.5 w-3.5" />
            Certificat SSL
          </p>
          <p className="mt-1.5 text-lg font-semibold text-ink tabular-nums">
            {formatDate(domain.sslExpiresAt)}
          </p>
          <p className="mt-0.5 text-xs font-medium text-[color:var(--tone-fg)]">
            {domain.sslExpiresAt ? countdownLabel(sslDays) : "Non suivi"}
          </p>
        </div>

        <div className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Coût de renouvellement</p>
          <p className="mt-1.5 text-lg font-semibold text-ink tabular-nums">
            {formatMoney(domain.renewalPrice)}
          </p>
          <p className="mt-0.5 text-xs text-ink-muted">par an, chez {domain.registrar ?? "—"}</p>
        </div>
      </div>

      <div className="grid gap-4 lg:grid-cols-2">
        <Card>
          <CardHeader title="Informations" icon={<Globe className="h-4 w-4" />} />
          <CardBody className="pt-0">
            <DetailList>
              <DetailRow label="Registrar">{domain.registrar}</DetailRow>
              <DetailRow label="Acheté le">{formatDate(domain.purchasedAt)}</DetailRow>
              <DetailRow label="Fournisseur DNS">{domain.dnsProvider}</DetailRow>
              <DetailRow label="Serveurs de noms" mono>
                {domain.nameservers ? (
                  <span className="whitespace-pre-wrap">{domain.nameservers}</span>
                ) : null}
              </DetailRow>
              <DetailRow label="Émetteur SSL">{domain.sslProvider}</DetailRow>
              <DetailRow label="Renouvellement SSL">
                {domain.sslAutoRenew ? "Automatique" : "Manuel"}
              </DetailRow>
              <DetailRow label="Projet" href={domain.project ? `/projets/${domain.project.id}` : null}>
                {domain.project?.name}
              </DetailRow>
              <DetailRow label="Remarques">
                {domain.notes ? <span className="whitespace-pre-wrap">{domain.notes}</span> : null}
              </DetailRow>
            </DetailList>
          </CardBody>
        </Card>

        <div className="space-y-4">
          <Card>
            <CardHeader title="Sous-domaines" description="Redirections et services rattachés." />
            <CardBody className="border-b border-line">
              <form action={addSubdomain} className="flex flex-wrap items-end gap-2">
                <Field label="Sous-domaine" htmlFor="sub-name" className="min-w-[130px] flex-1">
                  <Input id="sub-name" name="name" required placeholder="blog" />
                </Field>
                <Field label="Cible" htmlFor="sub-target" className="min-w-[130px] flex-1">
                  <Input id="sub-target" name="target" placeholder="192.0.2.10 ou CNAME" />
                </Field>
                <SubmitButton variant="default">
                  <Plus className="h-3.5 w-3.5" />
                  Ajouter
                </SubmitButton>
              </form>
            </CardBody>

            <div className="divide-y divide-line">
              {domain.subdomains.length > 0 ? (
                domain.subdomains.map((subdomain) => {
                  const remove = deleteSubdomain.bind(null, subdomain.id, domain.id);
                  return (
                    <div key={subdomain.id} className="flex items-center gap-3 px-4 py-2.5">
                      <div className="min-w-0 flex-1">
                        <p className="truncate font-mono text-xs text-ink">
                          {subdomain.name}.{domain.name}
                        </p>
                        {subdomain.target ? (
                          <p className="truncate text-xs text-ink-muted">→ {subdomain.target}</p>
                        ) : null}
                      </div>
                      <form action={remove}>
                        <Button variant="ghost" size="icon" type="submit" aria-label="Supprimer">
                          <Trash2 className="h-3.5 w-3.5" />
                        </Button>
                      </form>
                    </div>
                  );
                })
              ) : (
                <EmptyState compact title="Aucun sous-domaine" />
              )}
            </div>
          </Card>

          <CredentialVault
            credentials={domain.credentials.map((credential) => ({
              id: credential.id,
              label: credential.label,
              kind: credential.kind,
              username: credential.username,
              url: credential.url,
              notes: credential.notes,
              hasSecret: Boolean(credential.secretEnc),
            }))}
            scope={{ domainId: domain.id, clientId: domain.clientId ?? undefined }}
            revalidatePath={`/domaines/${domain.id}`}
          />
        </div>
      </div>

      {!domain.autoRenew && Number.isFinite(days) && days <= 60 ? (
        <Badge tone="warning">
          Reconduction manuelle — pensez à renouveler avant le {formatDate(domain.expiresAt)}
        </Badge>
      ) : null}
    </div>
  );
}

export async function generateMetadata({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const domain = await db.domain.findUnique({ where: { id }, select: { name: true } });
  return { title: domain?.name ?? "Domaine" };
}
