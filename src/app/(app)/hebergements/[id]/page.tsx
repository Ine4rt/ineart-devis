import { Pencil, RefreshCw, Server, Wrench } from "lucide-react";
import Link from "next/link";
import { notFound } from "next/navigation";

import { AutoRenewBadge } from "@/components/domain/renewal-list";
import { CredentialVault } from "@/components/domain/credential-vault";
import { ButtonLink } from "@/components/ui/button";
import { Card, CardBody, CardHeader } from "@/components/ui/card";
import { DetailList, DetailRow } from "@/components/ui/detail-list";
import { SubmitButton } from "@/components/ui/submit-button";
import { BILLING_CYCLE } from "@/lib/constants";
import { renewHosting } from "@/lib/actions/hostings";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { countdownLabel, daysUntil, toYearly, urgencyFor } from "@/lib/renewals";

export const dynamic = "force-dynamic";

export default async function HostingPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;

  const hosting = await db.hosting.findUnique({
    where: { id },
    include: {
      client: { select: { id: true, company: true } },
      project: { select: { id: true, name: true } },
      credentials: { orderBy: { createdAt: "desc" } },
    },
  });

  if (!hosting) notFound();

  const urgency = urgencyFor(hosting.renewsAt);
  const days = daysUntil(hosting.renewsAt);
  const yearly = toYearly(toNumber(hosting.price), hosting.billingCycle);
  const renew = renewHosting.bind(null, hosting.id);

  return (
    <div className="space-y-4">
      <header className="flex flex-wrap items-start justify-between gap-4">
        <div className="min-w-0">
          <p className="section-label">
            {hosting.client ? (
              <Link href={`/clients/${hosting.client.id}`} className="hover:text-accent-text">
                {hosting.client.company}
              </Link>
            ) : (
              "Non rattaché"
            )}
          </p>
          <div className="mt-1 flex flex-wrap items-center gap-2">
            <h1 className="text-xl font-semibold text-ink">{hosting.provider}</h1>
            {hosting.plan ? (
              <span className="text-sm text-ink-secondary">{hosting.plan}</span>
            ) : null}
            <AutoRenewBadge autoRenew={hosting.autoRenew} />
          </div>
        </div>

        <div className="flex items-center gap-2">
          <form action={renew}>
            <SubmitButton variant="default" pendingLabel="Renouvellement…">
              <RefreshCw className="h-3.5 w-3.5" />
              Marquer renouvelé
            </SubmitButton>
          </form>
          <ButtonLink href={`/hebergements/${hosting.id}/modifier`}>
            <Pencil className="h-3.5 w-3.5" />
            Modifier
          </ButtonLink>
        </div>
      </header>

      <div className="grid gap-3 sm:grid-cols-3">
        <div data-tone={urgency.tone} className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Prochaine échéance</p>
          <p className="mt-1.5 text-lg font-semibold text-ink tabular-nums">
            {formatDate(hosting.renewsAt)}
          </p>
          <p className="mt-0.5 text-xs font-medium text-[color:var(--tone-fg)]">
            {hosting.renewsAt ? countdownLabel(days) : "Date non renseignée"}
          </p>
        </div>

        <div className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Prix par cycle</p>
          <p className="mt-1.5 text-lg font-semibold text-ink tabular-nums">
            {formatMoney(hosting.price)}
          </p>
          <p className="mt-0.5 text-xs text-ink-muted">
            {BILLING_CYCLE[hosting.billingCycle].label}
          </p>
        </div>

        <div className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Équivalent annuel</p>
          <p className="mt-1.5 text-lg font-semibold text-ink tabular-nums">
            {formatMoney(yearly)}
          </p>
          <p className="mt-0.5 text-xs text-ink-muted">entre dans le calcul de marge</p>
        </div>
      </div>

      <div className="grid gap-4 lg:grid-cols-2">
        <Card>
          <CardHeader title="Offre" icon={<Server className="h-4 w-4" />} />
          <CardBody className="pt-0">
            <DetailList>
              <DetailRow label="Fournisseur">{hosting.provider}</DetailRow>
              <DetailRow label="Formule">{hosting.plan}</DetailRow>
              <DetailRow label="Libellé interne">{hosting.label}</DetailRow>
              <DetailRow
                label="Projet"
                href={hosting.project ? `/projets/${hosting.project.id}` : null}
              >
                {hosting.project?.name}
              </DetailRow>
              <DetailRow
                label="Panneau d'administration"
                href={hosting.controlPanelUrl}
              >
                {hosting.controlPanelUrl}
              </DetailRow>
              <DetailRow label="Remarques">
                {hosting.notes ? <span className="whitespace-pre-wrap">{hosting.notes}</span> : null}
              </DetailRow>
            </DetailList>
          </CardBody>
        </Card>

        <Card>
          <CardHeader title="Technique" icon={<Wrench className="h-4 w-4" />} />
          <CardBody className="pt-0">
            <DetailList>
              <DetailRow label="Adresse IP" copyValue={hosting.serverIp} mono>
                {hosting.serverIp}
              </DetailRow>
              <DetailRow label="Espace disque">
                {hosting.diskSpaceGb ? `${hosting.diskSpaceGb} Go` : null}
              </DetailRow>
              <DetailRow label="Bande passante">{hosting.bandwidth}</DetailRow>
              <DetailRow label="Région">{hosting.region}</DetailRow>
              <DetailRow label="Version PHP / runtime">{hosting.phpVersion}</DetailRow>
              <DetailRow label="Notes techniques">
                {hosting.technicalNotes ? (
                  <span className="whitespace-pre-wrap">{hosting.technicalNotes}</span>
                ) : null}
              </DetailRow>
            </DetailList>
          </CardBody>
        </Card>
      </div>

      <CredentialVault
        credentials={hosting.credentials.map((credential) => ({
          id: credential.id,
          label: credential.label,
          kind: credential.kind,
          username: credential.username,
          url: credential.url,
          notes: credential.notes,
          hasSecret: Boolean(credential.secretEnc),
        }))}
        scope={{ hostingId: hosting.id, clientId: hosting.clientId ?? undefined }}
        revalidatePath={`/hebergements/${hosting.id}`}
      />
    </div>
  );
}

export async function generateMetadata({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const hosting = await db.hosting.findUnique({ where: { id }, select: { provider: true } });
  return { title: hosting?.provider ?? "Hébergement" };
}
