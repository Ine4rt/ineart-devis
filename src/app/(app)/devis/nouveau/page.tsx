import type { Metadata } from "next";

import { QuoteEditor } from "@/components/domain/quote-editor";
import { PageHeader } from "@/components/ui/card";
import { createQuote } from "@/lib/actions/quotes";
import { getClientOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Nouveau devis" };
export const dynamic = "force-dynamic";

export default async function NewQuotePage({
  searchParams,
}: {
  searchParams: Promise<{ client?: string }>;
}) {
  const [{ client }, clients] = await Promise.all([searchParams, getClientOptions()]);

  return (
    <div className="mx-auto max-w-5xl space-y-4">
      <PageHeader
        eyebrow="Devis"
        title="Nouveau devis"
        description="Le numéro est attribué automatiquement (DEV-AAAA-001)."
      />
      <QuoteEditor
        clients={clients}
        defaultClientId={client}
        action={createQuote}
        submitLabel="Créer le devis"
        cancelHref="/devis"
      />
    </div>
  );
}
