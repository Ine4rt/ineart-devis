import { Globe, Plus, Server } from "lucide-react";
import Link from "next/link";

import { AutoRenewBadge } from "@/components/domain/renewal-list";
import { CredentialVault } from "@/components/domain/credential-vault";
import { ButtonLink } from "@/components/ui/button";
import { Card, CardHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { BILLING_CYCLE } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatDate, formatMoney } from "@/lib/format";
import { countdownLabel, daysUntil, urgencyFor } from "@/lib/renewals";

export const dynamic = "force-dynamic";

/**
 * Onglet technique : tout ce qu'il faut pour intervenir sur l'infrastructure
 * d'un client — domaines, hébergements et accès — au même endroit.
 */
export default async function ClientTechnicalPage({
  params,
}: {
  params: Promise<{ id: string }>;
}) {
  const { id } = await params;

  const [domains, hostings, credentials] = await Promise.all([
    db.domain.findMany({ where: { clientId: id }, orderBy: { expiresAt: "asc" } }),
    db.hosting.findMany({ where: { clientId: id }, orderBy: { renewsAt: "asc" } }),
    // Le coffre du client regroupe aussi les accès rattachés à ses domaines et
    // hébergements : on ne veut pas chercher un mot de passe à trois endroits.
    db.credential.findMany({
      where: {
        OR: [
          { clientId: id },
          { domain: { clientId: id } },
          { hosting: { clientId: id } },
          { project: { clientId: id } },
        ],
      },
      orderBy: { createdAt: "desc" },
    }),
  ]);

  return (
    <div className="space-y-4">
      <div className="grid gap-4 lg:grid-cols-2">
        <Card>
          <CardHeader
            title="Domaines"
            icon={<Globe className="h-4 w-4" />}
            action={
              <ButtonLink href={`/domaines/nouveau?client=${id}`} size="sm">
                <Plus className="h-3.5 w-3.5" />
                Ajouter
              </ButtonLink>
            }
          />
          <div className="divide-y divide-line">
            {domains.length > 0 ? (
              domains.map((domain) => {
                const urgency = urgencyFor(domain.expiresAt);
                return (
                  <Link
                    key={domain.id}
                    href={`/domaines/${domain.id}`}
                    className="flex items-center gap-3 px-4 py-2.5 transition-colors hover:bg-surface-inset"
                  >
                    <div className="min-w-0 flex-1">
                      <p className="truncate font-mono text-[13px] text-ink">{domain.name}</p>
                      <p className="truncate text-xs text-ink-muted">
                        {domain.registrar ?? "Registrar inconnu"} ·{" "}
                        {formatMoney(domain.renewalPrice)}/an
                      </p>
                    </div>
                    <div className="shrink-0 text-right">
                      <p className="text-xs text-ink">{formatDate(domain.expiresAt, "short")}</p>
                      <p
                        data-tone={urgency.tone}
                        className="text-[11px] text-[color:var(--tone-fg)]"
                      >
                        {domain.expiresAt ? countdownLabel(daysUntil(domain.expiresAt)) : "—"}
                      </p>
                    </div>
                  </Link>
                );
              })
            ) : (
              <EmptyState
                compact
                title="Aucun domaine"
                action={
                  <ButtonLink href={`/domaines/nouveau?client=${id}`} size="sm" variant="primary">
                    Ajouter un domaine
                  </ButtonLink>
                }
              />
            )}
          </div>
        </Card>

        <Card>
          <CardHeader
            title="Hébergements"
            icon={<Server className="h-4 w-4" />}
            action={
              <ButtonLink href={`/hebergements/nouveau?client=${id}`} size="sm">
                <Plus className="h-3.5 w-3.5" />
                Ajouter
              </ButtonLink>
            }
          />
          <div className="divide-y divide-line">
            {hostings.length > 0 ? (
              hostings.map((hosting) => (
                <Link
                  key={hosting.id}
                  href={`/hebergements/${hosting.id}`}
                  className="flex items-center gap-3 px-4 py-2.5 transition-colors hover:bg-surface-inset"
                >
                  <div className="min-w-0 flex-1">
                    <p className="truncate text-[13px] text-ink">
                      {hosting.provider}
                      {hosting.plan ? ` — ${hosting.plan}` : ""}
                    </p>
                    <p className="truncate text-xs text-ink-muted">
                      {formatMoney(hosting.price)} /{" "}
                      {BILLING_CYCLE[hosting.billingCycle].label.toLowerCase()}
                      {hosting.serverIp ? ` · ${hosting.serverIp}` : ""}
                    </p>
                  </div>
                  <div className="shrink-0 text-right">
                    <p className="text-xs text-ink">{formatDate(hosting.renewsAt, "short")}</p>
                    <AutoRenewBadge autoRenew={hosting.autoRenew} />
                  </div>
                </Link>
              ))
            ) : (
              <EmptyState
                compact
                title="Aucun hébergement"
                action={
                  <ButtonLink href={`/hebergements/nouveau?client=${id}`} size="sm" variant="primary">
                    Ajouter un hébergement
                  </ButtonLink>
                }
              />
            )}
          </div>
        </Card>
      </div>

      <CredentialVault
        credentials={credentials.map((credential) => ({
          id: credential.id,
          label: credential.label,
          kind: credential.kind,
          username: credential.username,
          url: credential.url,
          notes: credential.notes,
          hasSecret: Boolean(credential.secretEnc),
        }))}
        scope={{ clientId: id }}
        revalidatePath={`/clients/${id}/technique`}
      />
    </div>
  );
}
