import { Monitor, Palette, ShieldCheck } from "lucide-react";
import type { Metadata } from "next";

import { PasswordForm, ProfileForm } from "@/app/(app)/reglages/settings-forms";
import { ThemeToggle } from "@/components/layout/theme";
import { Card, CardBody, CardHeader, PageHeader } from "@/components/ui/card";
import { DetailList, DetailRow } from "@/components/ui/detail-list";
import { Tabs } from "@/components/ui/tabs";
import { requireUser } from "@/lib/auth";
import { db } from "@/lib/db";

export const metadata: Metadata = { title: "Réglages" };
export const dynamic = "force-dynamic";

export default async function SettingsPage() {
  const user = await requireUser();

  const [clients, projects, domains, hostings, subscriptions, documents, credentials] =
    await Promise.all([
      db.client.count(),
      db.project.count(),
      db.domain.count(),
      db.hosting.count(),
      db.subscription.count(),
      db.document.count(),
      db.credential.count(),
    ]);

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <PageHeader title="Réglages" description="Compte, apparence et état de l'instance." />

      <Tabs
        items={[
          { href: "/reglages", label: "Compte", exact: true },
          { href: "/reglages/donnees", label: "Données & export" },
        ]}
      />

      <div className="space-y-4 pt-1">
        <ProfileForm user={user} />
        <PasswordForm />

        <Card>
          <CardHeader
            title="Apparence"
            description="Le choix est mémorisé sur cet appareil."
            icon={<Palette className="h-4 w-4" />}
          />
          <CardBody className="flex items-center justify-between gap-4">
            <div>
              <p className="text-[13px] font-medium text-ink">Thème de l'interface</p>
              <p className="mt-0.5 text-xs text-ink-muted">
                « Système » suit automatiquement les réglages de votre ordinateur.
              </p>
            </div>
            <ThemeToggle />
          </CardBody>
        </Card>

        <Card>
          <CardHeader
            title="Sécurité"
            description="Comment vos données sensibles sont protégées."
            icon={<ShieldCheck className="h-4 w-4" />}
          />
          <CardBody className="pt-0">
            <DetailList>
              <DetailRow label="Accès">Compte unique, session signée (cookie httpOnly)</DetailRow>
              <DetailRow label="Mots de passe">Hachés avec scrypt et sel aléatoire</DetailRow>
              <DetailRow label="Coffre à identifiants">
                Chiffrement AES-256-GCM, déchiffrement à la demande uniquement
              </DetailRow>
              <DetailRow label="Documents">
                Stockés hors du dossier public, servis par une route authentifiée
              </DetailRow>
              <DetailRow label="Indexation">
                Désactivée (en-tête <code className="font-mono text-xs">X-Robots-Tag: noindex</code>)
              </DetailRow>
            </DetailList>
          </CardBody>
        </Card>

        <Card>
          <CardHeader
            title="Contenu de l'instance"
            icon={<Monitor className="h-4 w-4" />}
          />
          <CardBody className="pt-0">
            <DetailList>
              <DetailRow label="Clients">{clients}</DetailRow>
              <DetailRow label="Projets">{projects}</DetailRow>
              <DetailRow label="Domaines">{domains}</DetailRow>
              <DetailRow label="Hébergements">{hostings}</DetailRow>
              <DetailRow label="Abonnements">{subscriptions}</DetailRow>
              <DetailRow label="Documents">{documents}</DetailRow>
              <DetailRow label="Identifiants stockés">{credentials}</DetailRow>
            </DetailList>
          </CardBody>
        </Card>
      </div>
    </div>
  );
}
