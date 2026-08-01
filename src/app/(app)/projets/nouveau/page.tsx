import type { Metadata } from "next";

import { ProjectForm } from "@/components/domain/entity-forms";
import { PageHeader } from "@/components/ui/card";
import { createProject } from "@/lib/actions/projects";
import { getClientOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Nouveau projet" };
export const dynamic = "force-dynamic";

export default async function NewProjectPage({
  searchParams,
}: {
  searchParams: Promise<{ client?: string }>;
}) {
  const [{ client }, clients] = await Promise.all([searchParams, getClientOptions()]);

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader
        eyebrow="Projets"
        title="Nouveau projet"
        description="Un projet appartient toujours à un client. Domaine et hébergement se rattachent ensuite."
      />
      <ProjectForm
        clients={clients}
        defaultClientId={client}
        action={createProject}
        submitLabel="Créer le projet"
        cancelHref="/projets"
      />
    </div>
  );
}
