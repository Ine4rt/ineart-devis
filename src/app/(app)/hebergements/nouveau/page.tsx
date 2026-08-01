import type { Metadata } from "next";

import { HostingForm } from "@/components/domain/entity-forms";
import { PageHeader } from "@/components/ui/card";
import { createHosting } from "@/lib/actions/hostings";
import { getClientOptions, getProjectOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Nouvel hébergement" };
export const dynamic = "force-dynamic";

export default async function NewHostingPage({
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
        eyebrow="Hébergements"
        title="Nouvel hébergement"
        description="Le prix saisi est votre coût fournisseur — il alimente le calcul de rentabilité."
      />
      <HostingForm
        clients={clients}
        projects={projects}
        defaultClientId={client}
        action={createHosting}
        submitLabel="Enregistrer l'hébergement"
        cancelHref="/hebergements"
      />
    </div>
  );
}
