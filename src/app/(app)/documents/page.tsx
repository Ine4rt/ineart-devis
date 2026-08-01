import type { DocumentCategory, Prisma } from "@prisma/client";
import { Download, Paperclip } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { ListToolbar } from "@/components/domain/list-toolbar";
import { filterOptions } from "@/lib/filters";
import { Badge } from "@/components/ui/badge";
import { PageHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import {
  NumCell,
  NumHead,
  PrimaryCell,
  ResultCount,
  Table,
  TableScroll,
  TableWrapper,
} from "@/components/ui/table";
import { DOCUMENT_CATEGORY, DOCUMENT_CATEGORY_LIST } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatDate, formatFileSize } from "@/lib/format";

export const metadata: Metadata = { title: "Documents" };
export const dynamic = "force-dynamic";

/**
 * Bibliothèque transversale : tous les fichiers, tous clients confondus.
 * L'ajout se fait depuis la fiche du client concerné — un document orphelin
 * n'a pas d'intérêt dans cet outil.
 */
export default async function DocumentsPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; categorie?: string }>;
}) {
  const params = await searchParams;
  const query = params.q?.trim();

  const where: Prisma.DocumentWhereInput = {
    ...(params.categorie && params.categorie in DOCUMENT_CATEGORY
      ? { category: params.categorie as DocumentCategory }
      : {}),
    ...(query
      ? {
          OR: [
            { name: { contains: query } },
            { originalName: { contains: query } },
            { client: { company: { contains: query } } },
          ],
        }
      : {}),
  };

  const documents = await db.document.findMany({
    where,
    orderBy: { createdAt: "desc" },
    include: {
      client: { select: { id: true, company: true } },
      project: { select: { id: true, name: true } },
    },
  });

  const totalSize = documents.reduce((sum, document) => sum + document.size, 0);

  return (
    <div className="space-y-4">
      <PageHeader
        title="Documents"
        description="Tous les fichiers joints à vos clients et projets. L'ajout se fait depuis la fiche concernée."
      />

      <ListToolbar
        searchPlaceholder="Nom du fichier, client…"
        filters={[
          {
            name: "categorie",
            label: "Catégorie",
            options: filterOptions(DOCUMENT_CATEGORY_LIST),
          },
        ]}
      >
        <ResultCount count={documents.length} singular="document" plural="documents" />
      </ListToolbar>

      {documents.length === 0 ? (
        <TableWrapper>
          <EmptyState
            icon={<Paperclip className="h-4 w-4" />}
            title="Aucun document"
            description="Ouvrez une fiche client, onglet Documents, pour joindre un contrat ou une facture."
          />
        </TableWrapper>
      ) : (
        <TableWrapper>
          <TableScroll>
            <Table>
              <thead>
                <tr>
                  <th>Document</th>
                  <th>Client</th>
                  <th>Projet</th>
                  <th>Catégorie</th>
                  <th>Ajouté le</th>
                  <NumHead>Taille</NumHead>
                  <th />
                </tr>
              </thead>
              <tbody>
                {documents.map((document) => (
                  <tr key={document.id}>
                    <td>
                      <Link href={`/api/documents/${document.id}`} target="_blank" className="block">
                        <PrimaryCell title={document.name} subtitle={document.originalName} />
                      </Link>
                    </td>
                    <td>
                      {document.client ? (
                        <Link
                          href={`/clients/${document.client.id}/documents`}
                          className="text-ink-secondary hover:text-accent-text"
                        >
                          {document.client.company}
                        </Link>
                      ) : (
                        <span className="text-ink-muted">—</span>
                      )}
                    </td>
                    <td className="text-ink-secondary">{document.project?.name ?? "—"}</td>
                    <td>
                      <Badge tone={DOCUMENT_CATEGORY[document.category].tone} square dot={false}>
                        {DOCUMENT_CATEGORY[document.category].label}
                      </Badge>
                    </td>
                    <td className="whitespace-nowrap text-ink-muted">
                      {formatDate(document.createdAt, "short")}
                    </td>
                    <NumCell className="text-ink-muted">{formatFileSize(document.size)}</NumCell>
                    <td className="text-right">
                      <Link
                        href={`/api/documents/${document.id}`}
                        target="_blank"
                        className="btn btn-ghost h-8 w-8 p-0"
                        aria-label="Télécharger"
                      >
                        <Download className="h-3.5 w-3.5" />
                      </Link>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={5} className="text-xs font-medium text-ink-muted">
                    Espace utilisé
                  </td>
                  <NumCell className="font-semibold text-ink">{formatFileSize(totalSize)}</NumCell>
                  <td />
                </tr>
              </tfoot>
            </Table>
          </TableScroll>
        </TableWrapper>
      )}
    </div>
  );
}
