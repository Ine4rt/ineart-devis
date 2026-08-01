import type { Metadata } from "next";

import { DomainForm } from "@/components/domain/entity-forms";
import { PageHeader } from "@/components/ui/card";
import { createDomain } from "@/lib/actions/domains";
import { getClientOptions, getProjectOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Nouveau domaine" };
export const dynamic = "force-dynamic";

export default async function NewDomainPage({
  searchParams,
}: {
  searchParams: Promise<{ client?: string }>;
}) {
  const { client } = await searchParams;
  const [clients, projects] = await Promise.all([
    getClientOptions(),
    getProjectOptions(client ?? null),
  ]);

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader
        eyebrow="Domaines"
        title="Nouveau domaine"
        description="Renseignez la date d'expiration pour activer les rappels automatiques."
      />
      <DomainForm
        clients={clients}
        projects={projects}
        defaultClientId={client}
        action={createDomain}
        submitLabel="Enregistrer le domaine"
        cancelHref="/domaines"
      />
    </div>
  );
}
