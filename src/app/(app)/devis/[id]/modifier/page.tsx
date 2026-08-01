import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { QuoteEditor } from "@/components/domain/quote-editor";
import { PageHeader } from "@/components/ui/card";
import { updateQuote } from "@/lib/actions/quotes";
import { db } from "@/lib/db";
import { getClientOptions } from "@/lib/queries/options";

export const metadata: Metadata = { title: "Modifier le devis" };
export const dynamic = "force-dynamic";

export default async function EditQuotePage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const quote = await db.quote.findUnique({
    where: { id },
    include: { items: { orderBy: { position: "asc" } } },
  });
  if (!quote) notFound();

  const clients = await getClientOptions(quote.clientId);

  return (
    <div className="mx-auto max-w-5xl space-y-4">
      <PageHeader eyebrow="Devis" title={`Modifier — ${quote.number}`} />
      <QuoteEditor
        quote={quote}
        items={quote.items}
        clients={clients}
        action={updateQuote.bind(null, quote.id)}
        submitLabel="Enregistrer"
        cancelHref={`/devis/${quote.id}`}
      />
    </div>
  );
}
