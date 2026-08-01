import type { Metadata } from "next";

import { ClientForm } from "@/components/domain/client-form";
import { PageHeader } from "@/components/ui/card";
import { createClient } from "@/lib/actions/clients";

export const metadata: Metadata = { title: "Nouveau client" };

export default function NewClientPage() {
  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader
        eyebrow="Clients"
        title="Nouveau client"
        description="Seule la société est requise. Tout le reste peut être complété au fil de la relation."
      />
      <ClientForm action={createClient} submitLabel="Créer le client" />
    </div>
  );
}
