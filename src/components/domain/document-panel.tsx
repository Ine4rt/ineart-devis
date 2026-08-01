import type { Document, DocumentCategory } from "@prisma/client";
import { Download, FileText, Image as ImageIcon, Paperclip, Trash2, Upload } from "lucide-react";
import Link from "next/link";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardBody, CardHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { EnumSelect, Field, Input } from "@/components/ui/form";
import { SubmitButton } from "@/components/ui/submit-button";
import { DOCUMENT_CATEGORY, DOCUMENT_CATEGORY_LIST } from "@/lib/constants";
import { deleteDocument, uploadDocument } from "@/lib/actions/documents";
import { formatDate, formatFileSize } from "@/lib/format";

/**
 * Dépôt de fichiers d'un client ou d'un projet.
 *
 * Le formulaire poste directement vers une server action (`multipart/form-data`
 * est géré nativement) : pas de route d'upload séparée, pas de state machine
 * côté client.
 */

function iconFor(mimeType: string) {
  if (mimeType.startsWith("image/")) return <ImageIcon className="h-3.5 w-3.5" />;
  if (mimeType === "application/pdf") return <FileText className="h-3.5 w-3.5" />;
  return <Paperclip className="h-3.5 w-3.5" />;
}

export function DocumentPanel({
  documents,
  clientId,
  projectId,
  title = "Documents",
  description = "Contrats, factures, captures et pièces jointes.",
}: {
  documents: Document[];
  clientId?: string;
  projectId?: string;
  title?: string;
  description?: string;
}) {
  return (
    <Card>
      <CardHeader
        title={title}
        description={description}
        icon={<Paperclip className="h-4 w-4" />}
      />

      <CardBody className="border-b border-line">
        <form action={uploadDocument} className="flex flex-wrap items-end gap-2">
          {clientId ? <input type="hidden" name="clientId" value={clientId} /> : null}
          {projectId ? <input type="hidden" name="projectId" value={projectId} /> : null}

          <Field label="Fichier" htmlFor="doc-file" className="min-w-[200px] flex-1">
            <Input
              id="doc-file"
              name="file"
              type="file"
              required
              className="file:mr-2 file:rounded file:border-0 file:bg-surface-inset file:px-2 file:py-1 file:text-xs file:text-ink-secondary"
            />
          </Field>
          <Field label="Nom affiché" htmlFor="doc-name" className="min-w-[160px] flex-1">
            <Input id="doc-name" name="name" placeholder="Par défaut : le nom du fichier" />
          </Field>
          <Field label="Catégorie" htmlFor="doc-category" className="w-40">
            <EnumSelect
              id="doc-category"
              name="category"
              options={DOCUMENT_CATEGORY_LIST}
              defaultValue="CONTRACT"
            />
          </Field>
          <SubmitButton pendingLabel="Envoi…">
            <Upload className="h-3.5 w-3.5" />
            Joindre
          </SubmitButton>
        </form>
      </CardBody>

      <div className="divide-y divide-line">
        {documents.length > 0 ? (
          documents.map((document) => {
            const remove = deleteDocument.bind(null, document.id);
            return (
              <div key={document.id} className="group flex items-center gap-3 px-4 py-2.5">
                <span className="flex h-7 w-7 shrink-0 items-center justify-center rounded-md border border-line bg-surface-inset text-ink-muted">
                  {iconFor(document.mimeType)}
                </span>

                <div className="min-w-0 flex-1">
                  <Link
                    href={`/api/documents/${document.id}`}
                    target="_blank"
                    className="block truncate text-[13px] font-medium text-ink hover:text-accent-text"
                  >
                    {document.name}
                  </Link>
                  <p className="truncate text-xs text-ink-muted">
                    {formatFileSize(document.size)} · {formatDate(document.createdAt, "short")}
                    {document.notes ? ` · ${document.notes}` : ""}
                  </p>
                </div>

                <Badge tone={DOCUMENT_CATEGORY[document.category as DocumentCategory].tone} square dot={false}>
                  {DOCUMENT_CATEGORY[document.category as DocumentCategory].label}
                </Badge>

                <div className="flex shrink-0 items-center gap-1">
                  <Link
                    href={`/api/documents/${document.id}`}
                    target="_blank"
                    className="btn btn-ghost h-8 w-8 p-0"
                    aria-label="Télécharger"
                  >
                    <Download className="h-3.5 w-3.5" />
                  </Link>
                  <form action={remove} className="opacity-0 transition-opacity group-hover:opacity-100">
                    <Button variant="ghost" size="icon" type="submit" aria-label="Supprimer">
                      <Trash2 className="h-3.5 w-3.5" />
                    </Button>
                  </form>
                </div>
              </div>
            );
          })
        ) : (
          <EmptyState
            compact
            icon={<Paperclip className="h-4 w-4" />}
            title="Aucun document"
            description="Joignez le contrat signé, les factures ou toute pièce utile."
          />
        )}
      </div>
    </Card>
  );
}
