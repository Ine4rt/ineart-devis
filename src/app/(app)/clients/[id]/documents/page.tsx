import { DocumentPanel } from "@/components/domain/document-panel";
import { db } from "@/lib/db";

export const dynamic = "force-dynamic";

export default async function ClientDocumentsPage({
  params,
}: {
  params: Promise<{ id: string }>;
}) {
  const { id } = await params;

  const documents = await db.document.findMany({
    where: { clientId: id },
    orderBy: { createdAt: "desc" },
  });

  return (
    <DocumentPanel
      documents={documents}
      clientId={id}
      description="Contrat signé, factures, captures, cahier des charges — tout ce qui concerne ce client."
    />
  );
}
