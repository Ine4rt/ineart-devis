import { FolderKanban, Plus } from "lucide-react";
import Link from "next/link";

import { EnumBadge } from "@/components/ui/badge";
import { ButtonLink } from "@/components/ui/button";
import { Card, CardHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { Progress } from "@/components/ui/progress";
import { PROJECT_STATUS, PROJECT_TYPE } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatDate } from "@/lib/format";

export const dynamic = "force-dynamic";

export default async function ClientProjectsPage({
  params,
}: {
  params: Promise<{ id: string }>;
}) {
  const { id } = await params;

  const projects = await db.project.findMany({
    where: { clientId: id },
    orderBy: { updatedAt: "desc" },
    include: { _count: { select: { domains: true, subscriptions: true } } },
  });

  return (
    <Card>
      <CardHeader
        title="Projets du client"
        description="Tous les sites conçus pour cette entreprise."
        icon={<FolderKanban className="h-4 w-4" />}
        action={
          <ButtonLink href={`/projets/nouveau?client=${id}`} size="sm" variant="primary">
            <Plus className="h-3.5 w-3.5" />
            Nouveau projet
          </ButtonLink>
        }
      />

      <div className="divide-y divide-line">
        {projects.length > 0 ? (
          projects.map((project) => (
            <Link
              key={project.id}
              href={`/projets/${project.id}`}
              className="flex flex-wrap items-center gap-4 px-4 py-3 transition-colors hover:bg-surface-inset"
            >
              <div className="min-w-[200px] flex-1">
                <div className="flex flex-wrap items-center gap-2">
                  <span className="text-[13px] font-medium text-ink">{project.name}</span>
                  <EnumBadge meta={PROJECT_STATUS[project.status]} />
                </div>
                <p className="mt-0.5 truncate text-xs text-ink-muted">
                  {project.url?.replace(/^https?:\/\//, "") ?? "Pas encore d'URL"}
                  {" · "}
                  {project.isCustom ? "Sur mesure" : (project.cms ?? PROJECT_TYPE[project.type].label)}
                </p>
              </div>

              <div className="w-40">
                <Progress value={project.progress} showValue size="sm" />
              </div>

              <div className="w-32 text-right text-xs text-ink-muted">
                {project.launchedAt
                  ? `En ligne le ${formatDate(project.launchedAt, "short")}`
                  : `Créé le ${formatDate(project.createdAt, "short")}`}
              </div>
            </Link>
          ))
        ) : (
          <EmptyState
            icon={<FolderKanban className="h-4 w-4" />}
            title="Aucun projet"
            description="Créez le premier site pour ce client."
            action={
              <ButtonLink href={`/projets/nouveau?client=${id}`} variant="primary">
                Créer un projet
              </ButtonLink>
            }
          />
        )}
      </div>
    </Card>
  );
}
