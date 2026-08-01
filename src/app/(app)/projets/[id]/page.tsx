import { ExternalLink, GitBranch, History, Pencil, Plus } from "lucide-react";
import Link from "next/link";
import { notFound } from "next/navigation";

import { AutoRenewBadge } from "@/components/domain/renewal-list";
import { Badge, EnumBadge } from "@/components/ui/badge";
import { ButtonLink } from "@/components/ui/button";
import { Card, CardBody, CardHeader } from "@/components/ui/card";
import { DetailList, DetailRow } from "@/components/ui/detail-list";
import { EmptyState } from "@/components/ui/empty-state";
import { Field, Input } from "@/components/ui/form";
import { Progress } from "@/components/ui/progress";
import { SubmitButton } from "@/components/ui/submit-button";
import { PROJECT_STATUS, PROJECT_TYPE, SUBSCRIPTION_STATUS } from "@/lib/constants";
import { createMilestone } from "@/lib/actions/clients";
import { db } from "@/lib/db";
import { formatDate, formatMoney } from "@/lib/format";
import { parseJsonArray } from "@/lib/utils";

export const dynamic = "force-dynamic";

export default async function ProjectPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;

  const project = await db.project.findUnique({
    where: { id },
    include: {
      client: { select: { id: true, company: true } },
      domains: true,
      hostings: true,
      subscriptions: true,
      milestones: { orderBy: { happenedAt: "desc" } },
      documents: { orderBy: { createdAt: "desc" }, take: 5 },
    },
  });

  if (!project) notFound();

  const stack = parseJsonArray(project.techStack);
  const addMilestone = createMilestone.bind(null, project.id);

  return (
    <div className="space-y-4">
      <header className="flex flex-wrap items-start justify-between gap-4">
        <div className="min-w-0">
          <p className="section-label">
            <Link href={`/clients/${project.client.id}`} className="hover:text-accent-text">
              {project.client.company}
            </Link>
          </p>
          <div className="mt-1 flex flex-wrap items-center gap-2">
            <h1 className="text-xl font-semibold text-ink">{project.name}</h1>
            <EnumBadge meta={PROJECT_STATUS[project.status]} />
          </div>
          {project.url ? (
            <a
              href={project.url.startsWith("http") ? project.url : `https://${project.url}`}
              target="_blank"
              rel="noreferrer"
              className="mt-1.5 inline-flex items-center gap-1 text-sm link-subtle"
            >
              <ExternalLink className="h-3 w-3" />
              {project.url.replace(/^https?:\/\//, "")}
            </a>
          ) : null}
        </div>

        <ButtonLink href={`/projets/${project.id}/modifier`}>
          <Pencil className="h-3.5 w-3.5" />
          Modifier
        </ButtonLink>
      </header>

      <div className="grid gap-4 lg:grid-cols-3">
        <div className="space-y-4 lg:col-span-2">
          <Card>
            <CardHeader
              title="Avancement"
              action={
                <span className="text-sm font-semibold text-ink tabular-nums">
                  {project.progress}%
                </span>
              }
            />
            <CardBody className="space-y-4">
              <Progress
                value={project.progress}
                tone={project.progress === 100 ? "success" : "accent"}
              />
              {project.description ? (
                <p className="text-sm leading-relaxed text-ink-secondary">{project.description}</p>
              ) : null}
            </CardBody>
          </Card>

          <Card>
            <CardHeader title="Caractéristiques" />
            <CardBody className="pt-0">
              <DetailList>
                <DetailRow label="Type">{PROJECT_TYPE[project.type].label}</DetailRow>
                <DetailRow label="Conception">
                  {project.isCustom ? "Développement sur mesure" : (project.cms ?? "CMS")}
                </DetailRow>
                <DetailRow label="Technologies">
                  {stack.length > 0 ? (
                    <span className="flex flex-wrap gap-1">
                      {stack.map((tech) => (
                        <Badge key={tech} tone="muted" dot={false} square>
                          {tech}
                        </Badge>
                      ))}
                    </span>
                  ) : null}
                </DetailRow>
                <DetailRow label="Préproduction" href={project.stagingUrl}>
                  {project.stagingUrl}
                </DetailRow>
                <DetailRow label="Dépôt Git" href={project.repositoryUrl}>
                  {project.repositoryUrl ? (
                    <span className="inline-flex items-center gap-1">
                      <GitBranch className="h-3 w-3" />
                      {project.repositoryUrl.replace(/^https?:\/\//, "")}
                    </span>
                  ) : null}
                </DetailRow>
                <DetailRow label="Démarré le">{formatDate(project.startedAt)}</DetailRow>
                <DetailRow label="Mis en ligne le">{formatDate(project.launchedAt)}</DetailRow>
                <DetailRow label="Remarques">
                  {project.notes ? (
                    <span className="whitespace-pre-wrap">{project.notes}</span>
                  ) : null}
                </DetailRow>
              </DetailList>
            </CardBody>
          </Card>

          <Card>
            <CardHeader
              title="Historique des modifications"
              description="Chaque évolution notable du site, datée."
              icon={<History className="h-4 w-4" />}
            />
            <CardBody className="border-b border-line">
              <form action={addMilestone} className="flex flex-wrap items-end gap-2">
                <Field label="Événement" htmlFor="milestone-title" className="min-w-[200px] flex-1">
                  <Input
                    id="milestone-title"
                    name="title"
                    required
                    placeholder="Refonte de la page d'accueil"
                  />
                </Field>
                <Field label="Date" htmlFor="milestone-date">
                  <Input id="milestone-date" name="happenedAt" type="date" />
                </Field>
                <SubmitButton variant="default">
                  <Plus className="h-3.5 w-3.5" />
                  Ajouter
                </SubmitButton>
              </form>
            </CardBody>

            <div className="divide-y divide-line">
              {project.milestones.length > 0 ? (
                project.milestones.map((milestone) => (
                  <div key={milestone.id} className="flex items-start gap-3 px-4 py-2.5">
                    <span className="mt-1.5 h-1.5 w-1.5 shrink-0 rounded-full bg-[var(--line-strong)]" />
                    <div className="min-w-0 flex-1">
                      <p className="text-[13px] text-ink">{milestone.title}</p>
                      {milestone.detail ? (
                        <p className="mt-0.5 text-xs text-ink-muted">{milestone.detail}</p>
                      ) : null}
                    </div>
                    <span className="shrink-0 text-xs text-ink-muted tabular-nums">
                      {formatDate(milestone.happenedAt, "short")}
                    </span>
                  </div>
                ))
              ) : (
                <EmptyState compact title="Aucun jalon" description="L'historique se construit ici." />
              )}
            </div>
          </Card>
        </div>

        <div className="space-y-4">
          <Card>
            <CardHeader title="Domaines" />
            <div className="divide-y divide-line">
              {project.domains.length > 0 ? (
                project.domains.map((domain) => (
                  <Link
                    key={domain.id}
                    href={`/domaines/${domain.id}`}
                    className="flex items-center justify-between gap-2 px-4 py-2.5 transition-colors hover:bg-surface-inset"
                  >
                    <span className="min-w-0 truncate text-[13px] text-ink">{domain.name}</span>
                    <AutoRenewBadge autoRenew={domain.autoRenew} />
                  </Link>
                ))
              ) : (
                <EmptyState
                  compact
                  title="Aucun domaine"
                  action={
                    <ButtonLink href={`/domaines/nouveau?client=${project.clientId}`} size="sm">
                      Rattacher un domaine
                    </ButtonLink>
                  }
                />
              )}
            </div>
          </Card>

          <Card>
            <CardHeader title="Hébergement" />
            <div className="divide-y divide-line">
              {project.hostings.length > 0 ? (
                project.hostings.map((hosting) => (
                  <Link
                    key={hosting.id}
                    href={`/hebergements/${hosting.id}`}
                    className="flex items-center justify-between gap-2 px-4 py-2.5 transition-colors hover:bg-surface-inset"
                  >
                    <span className="min-w-0 truncate text-[13px] text-ink">
                      {hosting.provider}
                      {hosting.plan ? ` — ${hosting.plan}` : ""}
                    </span>
                    <span className="shrink-0 text-xs text-ink-muted tabular-nums">
                      {formatMoney(hosting.price)}
                    </span>
                  </Link>
                ))
              ) : (
                <EmptyState
                  compact
                  title="Aucun hébergement"
                  action={
                    <ButtonLink href={`/hebergements/nouveau?client=${project.clientId}`} size="sm">
                      Rattacher un hébergement
                    </ButtonLink>
                  }
                />
              )}
            </div>
          </Card>

          <Card>
            <CardHeader title="Abonnements" />
            <div className="divide-y divide-line">
              {project.subscriptions.length > 0 ? (
                project.subscriptions.map((subscription) => (
                  <Link
                    key={subscription.id}
                    href={`/abonnements/${subscription.id}`}
                    className="block px-4 py-2.5 transition-colors hover:bg-surface-inset"
                  >
                    <div className="flex items-center justify-between gap-2">
                      <span className="min-w-0 truncate text-[13px] text-ink">
                        {subscription.name}
                      </span>
                      <span className="shrink-0 text-xs font-medium text-ink tabular-nums">
                        {formatMoney(subscription.amount, { currency: subscription.currency })}
                      </span>
                    </div>
                    <div className="mt-1">
                      <EnumBadge meta={SUBSCRIPTION_STATUS[subscription.status]} />
                    </div>
                  </Link>
                ))
              ) : (
                <EmptyState
                  compact
                  title="Aucun abonnement"
                  description="Ce site n'est couvert par aucun contrat récurrent."
                  action={
                    <ButtonLink
                      href={`/abonnements/nouveau?client=${project.clientId}`}
                      size="sm"
                      variant="primary"
                    >
                      Créer un abonnement
                    </ButtonLink>
                  }
                />
              )}
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}

export async function generateMetadata({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const project = await db.project.findUnique({ where: { id }, select: { name: true } });
  return { title: project?.name ?? "Projet" };
}
