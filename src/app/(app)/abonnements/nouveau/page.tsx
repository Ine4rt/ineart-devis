import type { Metadata } from "next";

import { SubscriptionForm } from "@/components/domain/entity-forms";
import { PageHeader } from "@/components/ui/card";
import { createSubscription } from "@/lib/actions/subscriptions";
import { getClientOptions, getProjectOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Nouvel abonnement" };
export const dynamic = "force-dynamic";

export default async function NewSubscriptionPage({
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
        eyebrow="Abonnements"
        title="Nouvel abonnement"
        description="La première échéance de paiement est créée automatiquement."
      />
      <SubscriptionForm
        clients={clients}
        projects={projects}
        defaultClientId={client}
        action={createSubscription}
        submitLabel="Créer l'abonnement"
        cancelHref="/abonnements"
      />
    </div>
  );
}
