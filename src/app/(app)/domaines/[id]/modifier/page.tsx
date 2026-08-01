import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { DomainForm } from "@/components/domain/entity-forms";
import { PageHeader } from "@/components/ui/card";
import { updateDomain } from "@/lib/actions/domains";
import { db } from "@/lib/db";
import { getClientOptions, getProjectOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Modifier le domaine" };
export const dynamic = "force-dynamic";

export default async function EditDomainPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const domain = await db.domain.findUnique({ where: { id } });
  if (!domain) notFound();

  const [clients, projects] = await Promise.all([
    getClientOptions(domain.clientId),
    getProjectOptions(domain.clientId),
  ]);

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader eyebrow="Domaines" title={`Modifier — ${domain.name}`} />
      <DomainForm
        domain={domain}
        clients={clients}
        projects={projects}
        action={updateDomain.bind(null, domain.id)}
        submitLabel="Enregistrer"
        cancelHref={`/domaines/${domain.id}`}
      />
    </div>
  );
}
