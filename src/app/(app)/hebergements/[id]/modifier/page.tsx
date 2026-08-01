import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { HostingForm } from "@/components/domain/entity-forms";
import { PageHeader } from "@/components/ui/card";
import { updateHosting } from "@/lib/actions/hostings";
import { db } from "@/lib/db";
import { getClientOptions, getProjectOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Modifier l'hébergement" };
export const dynamic = "force-dynamic";

export default async function EditHostingPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const hosting = await db.hosting.findUnique({ where: { id } });
  if (!hosting) notFound();

  const [clients, projects] = await Promise.all([
    getClientOptions(hosting.clientId),
    getProjectOptions(hosting.clientId),
  ]);

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader eyebrow="Hébergements" title={`Modifier — ${hosting.provider}`} />
      <HostingForm
        hosting={hosting}
        clients={clients}
        projects={projects}
        action={updateHosting.bind(null, hosting.id)}
        submitLabel="Enregistrer"
        cancelHref={`/hebergements/${hosting.id}`}
      />
    </div>
  );
}
