import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { ProjectForm } from "@/components/domain/entity-forms";
import { PageHeader } from "@/components/ui/card";
import { updateProject } from "@/lib/actions/projects";
import { db } from "@/lib/db";
import { getClientOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Modifier le projet" };
export const dynamic = "force-dynamic";

export default async function EditProjectPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const project = await db.project.findUnique({ where: { id } });
  if (!project) notFound();

  const clients = await getClientOptions(project.clientId);

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader eyebrow="Projets" title={`Modifier — ${project.name}`} />
      <ProjectForm
        project={project}
        clients={clients}
        action={updateProject.bind(null, project.id)}
        submitLabel="Enregistrer"
        cancelHref={`/projets/${project.id}`}
      />
    </div>
  );
}
