import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { ClientForm } from "@/components/domain/client-form";
import { updateClient } from "@/lib/actions/clients";
import { db } from "@/lib/db";

export const metadata: Metadata = { title: "Modifier le client" };

export default async function EditClientPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const client = await db.client.findUnique({ where: { id } });
  if (!client) notFound();

  // `bind` fige l'identifiant côté serveur : il n'est pas modifiable depuis le
  // formulaire, contrairement à un champ caché.
  const action = updateClient.bind(null, client.id);

  return (
    <div className="mx-auto max-w-4xl">
      <ClientForm client={client} action={action} submitLabel="Enregistrer les modifications" />
    </div>
  );
}
