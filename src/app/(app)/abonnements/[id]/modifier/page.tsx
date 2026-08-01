import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { SubscriptionForm } from "@/components/domain/entity-forms";
import { PageHeader } from "@/components/ui/card";
import { updateSubscription } from "@/lib/actions/subscriptions";
import { db } from "@/lib/db";
import { getClientOptions, getProjectOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Modifier l'abonnement" };
export const dynamic = "force-dynamic";

export default async function EditSubscriptionPage({
  params,
}: {
  params: Promise<{ id: string }>;
}) {
  const { id } = await params;
  const subscription = await db.subscription.findUnique({ where: { id } });
  if (!subscription) notFound();

  const [clients, projects] = await Promise.all([
    getClientOptions(subscription.clientId),
    getProjectOptions(subscription.clientId),
  ]);

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader eyebrow="Abonnements" title={`Modifier — ${subscription.name}`} />
      <SubscriptionForm
        subscription={subscription}
        clients={clients}
        projects={projects}
        action={updateSubscription.bind(null, subscription.id)}
        submitLabel="Enregistrer"
        cancelHref={`/abonnements/${subscription.id}`}
      />
    </div>
  );
}
