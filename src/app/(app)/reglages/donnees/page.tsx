import { Database, Download, HardDrive, KeyRound } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

import { Card, CardBody, CardHeader, PageHeader } from "@/components/ui/card";
import { Tabs } from "@/components/ui/tabs";
import { db } from "@/lib/db";
import { formatFileSize } from "@/lib/format";

export const metadata: Metadata = { title: "Données & export" };
export const dynamic = "force-dynamic";

export default async function DataSettingsPage() {
  const documents = await db.document.aggregate({ _sum: { size: true }, _count: { _all: true } });

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader
        title="Réglages"
        description="Sauvegarde, export et rappels d'exploitation."
      />

      <Tabs
        items={[
          { href: "/reglages", label: "Compte", exact: true },
          { href: "/reglages/donnees", label: "Données & export" },
        ]}
      />

      <div className="space-y-4 pt-1">
        <Card>
          <CardHeader
            title="Export complet"
            description="Toutes vos données au format JSON, réutilisables ailleurs."
            icon={<Download className="h-4 w-4" />}
          />
          <CardBody className="flex flex-wrap items-center justify-between gap-4">
            <p className="max-w-md text-xs leading-relaxed text-ink-muted">
              Clients, projets, domaines, hébergements, abonnements, échéances, devis, tâches, notes
              et métadonnées de documents. Les secrets du coffre sont exportés chiffrés : sans la clé{" "}
              <code className="font-mono">APP_ENCRYPTION_KEY</code>, ils restent illisibles.
            </p>
            <Link href="/api/export" className="btn btn-primary h-8 px-3" download>
              <Download className="h-3.5 w-3.5" />
              Télécharger l'export
            </Link>
          </CardBody>
        </Card>

        <Card>
          <CardHeader
            title="Sauvegarde"
            description="Ce qu'il faut copier pour ne rien perdre."
            icon={<Database className="h-4 w-4" />}
          />
          <CardBody className="space-y-3 text-xs leading-relaxed text-ink-secondary">
            <p>
              Trois éléments composent l'état complet de la console. Sauvegardez-les ensemble :
            </p>
            <ol className="ml-4 list-decimal space-y-1.5">
              <li>
                <code className="font-mono text-ink">prisma/dev.db</code> — la base SQLite. Une
                simple copie du fichier suffit ; il n'y a pas de serveur à arrêter pour cela.
              </li>
              <li>
                <code className="font-mono text-ink">storage/documents/</code> — les fichiers joints
                (contrats, factures, captures).
              </li>
              <li>
                <code className="font-mono text-ink">.env</code> — sans{" "}
                <code className="font-mono text-ink">APP_ENCRYPTION_KEY</code>, les identifiants du
                coffre sont définitivement illisibles.
              </li>
            </ol>
            <p className="text-ink-muted">
              Une copie hebdomadaire vers un disque externe ou un stockage distant suffit largement
              à cette échelle.
            </p>
          </CardBody>
        </Card>

        <div className="grid gap-4 sm:grid-cols-2">
          <Card>
            <CardHeader title="Fichiers stockés" icon={<HardDrive className="h-4 w-4" />} />
            <CardBody>
              <p className="text-2xl font-semibold text-ink tabular-nums">
                {formatFileSize(documents._sum.size ?? 0)}
              </p>
              <p className="mt-1 text-xs text-ink-muted">
                {documents._count._all} document(s) dans le stockage local
              </p>
            </CardBody>
          </Card>

          <Card>
            <CardHeader title="Clés de chiffrement" icon={<KeyRound className="h-4 w-4" />} />
            <CardBody className="space-y-2 text-xs leading-relaxed text-ink-secondary">
              <p>
                Changer <code className="font-mono text-ink">APP_SESSION_SECRET</code> déconnecte
                immédiatement toutes les sessions ouvertes — utile en cas de doute.
              </p>
              <p>
                Ne changez jamais <code className="font-mono text-ink">APP_ENCRYPTION_KEY</code> sans
                avoir d'abord noté ailleurs les identifiants du coffre.
              </p>
            </CardBody>
          </Card>
        </div>
      </div>
    </div>
  );
}
