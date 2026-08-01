import { Contact, Plus, StickyNote, Trash2 } from "lucide-react";
import Link from "next/link";
import { notFound } from "next/navigation";

import { RenewalRow } from "@/components/domain/renewal-list";
import { Badge } from "@/components/ui/badge";
import { Button, ButtonLink } from "@/components/ui/button";
import { Card, CardBody, CardHeader, CardLink } from "@/components/ui/card";
import { DetailList, DetailRow } from "@/components/ui/detail-list";
import { EmptyState } from "@/components/ui/empty-state";
import { Checkbox, Field, Input } from "@/components/ui/form";
import { StatTile } from "@/components/ui/stat-tile";
import { SubmitButton } from "@/components/ui/submit-button";
import { createContact, deleteContact } from "@/lib/actions/clients";
import { db } from "@/lib/db";
import { formatDate, formatMoney, formatRelative, toNumber } from "@/lib/format";
import { getRenewalFeed } from "@/lib/queries/renewal-feed";
import { toYearly } from "@/lib/renewals";

export const dynamic = "force-dynamic";

export default async function ClientOverviewPage({
  params,
}: {
  params: Promise<{ id: string }>;
}) {
  const { id } = await params;

  const [client, renewals] = await Promise.all([
    db.client.findUnique({
      where: { id },
      include: {
        contacts: { orderBy: [{ isPrimary: "desc" }, { firstName: "asc" }] },
        projects: { orderBy: { updatedAt: "desc" }, take: 4 },
        subscriptions: { where: { status: { in: ["ACTIVE", "PAST_DUE"] } } },
        domains: { select: { renewalPrice: true } },
        hostings: { select: { price: true, billingCycle: true } },
        activities: { orderBy: { createdAt: "desc" }, take: 8 },
        notesList: { where: { pinned: true }, orderBy: { updatedAt: "desc" }, take: 3 },
      },
    }),
    getRenewalFeed(180),
  ]);

  if (!client) notFound();

  const arr = client.subscriptions.reduce(
    (sum, subscription) => sum + toYearly(toNumber(subscription.amount), subscription.billingCycle),
    0,
  );
  const cost =
    client.domains.reduce((sum, domain) => sum + toNumber(domain.renewalPrice), 0) +
    client.hostings.reduce(
      (sum, hosting) => sum + toYearly(toNumber(hosting.price), hosting.billingCycle),
      0,
    );

  const clientRenewals = renewals.filter((item) => item.clientId === client.id).slice(0, 6);
  const addContact = createContact.bind(null, client.id);

  return (
    <div className="grid gap-4 lg:grid-cols-3">
      <div className="space-y-4 lg:col-span-2">
        <div className="grid gap-3 sm:grid-cols-3">
          <StatTile label="Revenu annuel" value={formatMoney(arr)} tone="success" />
          <StatTile label="Coûts annuels" value={formatMoney(cost)} tone="warning" />
          <StatTile
            label="Marge"
            value={formatMoney(arr - cost)}
            hint={arr > 0 ? `${Math.round(((arr - cost) / arr) * 100)}% du revenu` : undefined}
            tone="accent"
          />
        </div>

        <Card>
          <CardHeader title="Coordonnées" />
          <CardBody className="pt-0">
            <DetailList>
              <DetailRow label="Société">{client.company}</DetailRow>
              <DetailRow label="Interlocuteur">
                {[client.firstName, client.lastName].filter(Boolean).join(" ") || null}
              </DetailRow>
              <DetailRow label="E-mail" copyValue={client.email} href={client.email ? `mailto:${client.email}` : null}>
                {client.email}
              </DetailRow>
              <DetailRow label="Téléphone" copyValue={client.phone}>
                {client.phone}
              </DetailRow>
              <DetailRow label="Mobile" copyValue={client.mobile}>
                {client.mobile}
              </DetailRow>
              <DetailRow label="Adresse">
                {client.addressLine1 ? (
                  <span className="whitespace-pre-line">
                    {[
                      client.addressLine1,
                      client.addressLine2,
                      [client.postalCode, client.city].filter(Boolean).join(" "),
                      client.country,
                    ]
                      .filter(Boolean)
                      .join("\n")}
                  </span>
                ) : null}
              </DetailRow>
              <DetailRow label="Numéro de TVA" copyValue={client.vatNumber} mono>
                {client.vatNumber}
              </DetailRow>
              <DetailRow label="Numéro d'entreprise" copyValue={client.companyNumber} mono>
                {client.companyNumber}
              </DetailRow>
              <DetailRow label="Secteur">{client.industry}</DetailRow>
              <DetailRow label="Origine">{client.source}</DetailRow>
              <DetailRow label={client.wonAt ? "Client depuis" : "Fiche créée le"}>
                {formatDate(client.wonAt ?? client.createdAt)}
              </DetailRow>
            </DetailList>
          </CardBody>
        </Card>

        <Card>
          <CardHeader
            title="Interlocuteurs"
            description="Les autres personnes à contacter chez ce client."
            icon={<Contact className="h-4 w-4" />}
          />

          <CardBody className="border-b border-line">
            <form action={addContact} className="flex flex-wrap items-end gap-2">
              <Field label="Prénom" htmlFor="c-firstName" className="min-w-[110px] flex-1">
                <Input id="c-firstName" name="firstName" required />
              </Field>
              <Field label="Nom" htmlFor="c-lastName" className="min-w-[110px] flex-1">
                <Input id="c-lastName" name="lastName" />
              </Field>
              <Field label="Fonction" htmlFor="c-role" className="min-w-[110px] flex-1">
                <Input id="c-role" name="role" placeholder="Gérant" />
              </Field>
              <Field label="E-mail" htmlFor="c-email" className="min-w-[150px] flex-1">
                <Input id="c-email" name="email" type="email" />
              </Field>
              <SubmitButton variant="default">
                <Plus className="h-3.5 w-3.5" />
                Ajouter
              </SubmitButton>
            </form>
          </CardBody>

          <div className="divide-y divide-line">
            {client.contacts.length > 0 ? (
              client.contacts.map((contact) => {
                const remove = deleteContact.bind(null, contact.id, client.id);
                return (
                  <div key={contact.id} className="group flex items-center gap-3 px-4 py-2.5">
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-2">
                        <span className="text-[13px] font-medium text-ink">
                          {[contact.firstName, contact.lastName].filter(Boolean).join(" ")}
                        </span>
                        {contact.isPrimary ? (
                          <Badge tone="accent" square dot={false}>
                            Principal
                          </Badge>
                        ) : null}
                      </div>
                      <p className="truncate text-xs text-ink-muted">
                        {[contact.role, contact.email, contact.phone].filter(Boolean).join(" · ")}
                      </p>
                    </div>
                    <form action={remove} className="opacity-0 transition-opacity group-hover:opacity-100">
                      <Button variant="ghost" size="icon" type="submit" aria-label="Supprimer">
                        <Trash2 className="h-3.5 w-3.5" />
                      </Button>
                    </form>
                  </div>
                );
              })
            ) : (
              <EmptyState compact title="Aucun interlocuteur supplémentaire" />
            )}
          </div>
        </Card>

        {client.notes ? (
          <Card>
            <CardHeader title="Notes privées" icon={<StickyNote className="h-4 w-4" />} />
            <CardBody>
              <p className="whitespace-pre-wrap text-sm leading-relaxed text-ink-secondary">
                {client.notes}
              </p>
            </CardBody>
          </Card>
        ) : null}
      </div>

      <div className="space-y-4">
        <Card>
          <CardHeader
            title="Projets"
            action={<CardLink href={`/clients/${client.id}/projets`}>Tous</CardLink>}
          />
          <div className="divide-y divide-line">
            {client.projects.length > 0 ? (
              client.projects.map((project) => (
                <Link
                  key={project.id}
                  href={`/projets/${project.id}`}
                  className="block px-4 py-2.5 transition-colors hover:bg-surface-inset"
                >
                  <p className="truncate text-[13px] font-medium text-ink">{project.name}</p>
                  <p className="truncate text-xs text-ink-muted">
                    {project.url?.replace(/^https?:\/\//, "") ?? "Pas d'URL"}
                  </p>
                </Link>
              ))
            ) : (
              <EmptyState
                compact
                title="Aucun projet"
                action={
                  <ButtonLink href={`/projets/nouveau?client=${client.id}`} size="sm" variant="primary">
                    Créer un projet
                  </ButtonLink>
                }
              />
            )}
          </div>
        </Card>

        <Card>
          <CardHeader title="Prochaines échéances" />
          <div className="divide-y divide-line">
            {clientRenewals.length > 0 ? (
              clientRenewals.map((item) => <RenewalRow key={item.id} item={item} />)
            ) : (
              <EmptyState compact title="Aucune échéance" description="Rien à renouveler sous 6 mois." />
            )}
          </div>
        </Card>

        {client.notesList.length > 0 ? (
          <Card>
            <CardHeader
              title="Notes épinglées"
              action={<CardLink href={`/clients/${client.id}/suivi`}>Toutes</CardLink>}
            />
            <div className="divide-y divide-line">
              {client.notesList.map((note) => (
                <div key={note.id} className="px-4 py-2.5">
                  {note.title ? (
                    <p className="text-xs font-semibold text-ink">{note.title}</p>
                  ) : null}
                  <p className="mt-0.5 line-clamp-3 text-xs text-ink-secondary">{note.body}</p>
                </div>
              ))}
            </div>
          </Card>
        ) : null}

        <Card>
          <CardHeader title="Historique" />
          <div className="divide-y divide-line">
            {client.activities.length > 0 ? (
              client.activities.map((activity) => (
                <div key={activity.id} className="px-4 py-2">
                  <p className="text-xs text-ink">{activity.summary.replace(/\s*:?\s*$/, "")}</p>
                  <p className="mt-0.5 text-[11px] text-ink-muted">
                    {formatRelative(activity.createdAt)}
                  </p>
                </div>
              ))
            ) : (
              <EmptyState compact title="Aucune activité" />
            )}
          </div>
        </Card>
      </div>
    </div>
  );
}
