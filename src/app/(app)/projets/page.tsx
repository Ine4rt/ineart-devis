import type { Prisma, ProjectStatus } from "@prisma/client";
import { FolderKanban, Plus } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { ListToolbar } from "@/components/domain/list-toolbar";
import { filterOptions } from "@/lib/filters";
import { Badge, EnumBadge } from "@/components/ui/badge";
import { ButtonLink } from "@/components/ui/button";
import { PageHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { Progress } from "@/components/ui/progress";
import {
  PrimaryCell,
  ResultCount,
  Table,
  TableScroll,
  TableWrapper,
} from "@/components/ui/table";
import { PROJECT_STATUS, PROJECT_STATUS_LIST, PROJECT_TYPE, PROJECT_TYPE_LIST } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatDate } from "@/lib/format";
import { parseJsonArray } from "@/lib/utils";

export const metadata: Metadata = { title: "Projets" };
export const dynamic = "force-dynamic";

export default async function ProjectsPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; statut?: string; type?: string; couverture?: string }>;
}) {
  const params = await searchParams;
  const query = params.q?.trim();

  const where: Prisma.ProjectWhereInput = {
    ...(params.statut && params.statut in PROJECT_STATUS
      ? { status: params.statut as ProjectStatus }
      : {}),
    ...(params.type && params.type in PROJECT_TYPE
      ? { type: params.type as keyof typeof PROJECT_TYPE }
      : {}),
    // Filtre issu d'une alerte du tableau de bord : les sites en ligne qui ne
    // rapportent rien. Le lien de l'alerte mène directement ici.
    ...(params.couverture === "sans-abonnement"
      ? { status: { in: ["LIVE", "MAINTENANCE"] }, subscriptions: { none: {} } }
      : {}),
    ...(query
      ? {
          OR: [
            { name: { contains: query } },
            { url: { contains: query } },
            { client: { company: { contains: query } } },
          ],
        }
      : {}),
  };

  const projects = await db.project.findMany({
    where,
    orderBy: [{ updatedAt: "desc" }],
    include: {
      client: { select: { id: true, company: true } },
      _count: { select: { subscriptions: true, domains: true } },
    },
  });

  return (
    <div className="space-y-4">
      <PageHeader
        title="Projets"
        description="Chaque site que vous concevez, de l'idée à la maintenance."
        actions={
          <ButtonLink href="/projets/nouveau" variant="primary">
            <Plus className="h-3.5 w-3.5" />
            Nouveau projet
          </ButtonLink>
        }
      />

      <ListToolbar
        searchPlaceholder="Nom du projet, URL, client…"
        filters={[
          { name: "statut", label: "Statut", options: filterOptions(PROJECT_STATUS_LIST) },
          { name: "type", label: "Type", options: filterOptions(PROJECT_TYPE_LIST) },
          {
            name: "couverture",
            label: "Couverture",
            options: [{ value: "sans-abonnement", label: "En ligne sans abonnement" }],
          },
        ]}
      >
        <ResultCount count={projects.length} singular="projet" plural="projets" />
      </ListToolbar>

      {projects.length === 0 ? (
        <TableWrapper>
          <EmptyState
            icon={<FolderKanban className="h-4 w-4" />}
            title="Aucun projet"
            description="Créez un projet et rattachez-lui un domaine, un hébergement puis un abonnement."
            action={
              <ButtonLink href="/projets/nouveau" variant="primary">
                Créer un projet
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
                  <th>Projet</th>
                  <th>Client</th>
                  <th>Statut</th>
                  <th>Type</th>
                  <th className="w-40">Avancement</th>
                  <th>Technologies</th>
                  <th>Mise en ligne</th>
                </tr>
              </thead>
              <tbody>
                {projects.map((project) => {
                  const stack = parseJsonArray(project.techStack);
                  const uncovered =
                    ["LIVE", "MAINTENANCE"].includes(project.status) &&
                    project._count.subscriptions === 0;

                  return (
                    <tr key={project.id}>
                      <td>
                        <Link href={`/projets/${project.id}`} className="block">
                          <PrimaryCell
                            title={project.name}
                            subtitle={project.url?.replace(/^https?:\/\//, "") ?? "Pas encore d'URL"}
                          />
                        </Link>
                      </td>
                      <td>
                        <Link
                          href={`/clients/${project.client.id}`}
                          className="text-ink-secondary hover:text-accent-text"
                        >
                          {project.client.company}
                        </Link>
                      </td>
                      <td>
                        <div className="flex items-center gap-1.5">
                          <EnumBadge meta={PROJECT_STATUS[project.status]} />
                          {uncovered ? (
                            <Badge tone="info" dot={false} title="Site en ligne sans abonnement actif">
                              non couvert
                            </Badge>
                          ) : null}
                        </div>
                      </td>
                      <td className="text-ink-secondary">
                        {project.isCustom ? "Sur mesure" : (project.cms ?? PROJECT_TYPE[project.type].label)}
                      </td>
                      <td>
                        <Progress value={project.progress} showValue size="sm" />
                      </td>
                      <td>
                        <div className="flex flex-wrap gap-1">
                          {stack.slice(0, 2).map((tech) => (
                            <Badge key={tech} tone="muted" dot={false} square>
                              {tech}
                            </Badge>
                          ))}
                          {stack.length > 2 ? (
                            <span className="text-xs text-ink-muted">+{stack.length - 2}</span>
                          ) : null}
                        </div>
                      </td>
                      <td className="whitespace-nowrap text-ink-muted">
                        {formatDate(project.launchedAt, "short")}
                      </td>
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
