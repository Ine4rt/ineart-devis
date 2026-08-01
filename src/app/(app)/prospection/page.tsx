import { Plus, Target } from "lucide-react";
import type { Metadata } from "next";

import { PipelineBoard, type ProspectCard } from "@/app/(app)/prospection/pipeline-board";
import { ButtonLink } from "@/components/ui/button";
import { Card, PageHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { StatTile } from "@/components/ui/stat-tile";
import { PIPELINE_STAGE } from "@/lib/constants";
import { db } from "@/lib/db";
import { formatMoney, toNumber } from "@/lib/format";

export const metadata: Metadata = { title: "Prospection" };
export const dynamic = "force-dynamic";

/**
 * Pipeline commercial.
 *
 * Un prospect n'est pas une entité distincte : c'est un client au statut
 * PROSPECT. Le convertir revient à le déposer dans « Gagné », ce qui bascule
 * son statut en ACTIF — sans ressaisie, et sans risque de doublon entre une
 * table « leads » et une table « clients ».
 */
export default async function ProspectionPage() {
  const [prospects, wonThisYear, lostCount] = await Promise.all([
    db.client.findMany({
      where: { status: "PROSPECT" },
      orderBy: { updatedAt: "desc" },
      select: {
        id: true,
        company: true,
        firstName: true,
        lastName: true,
        city: true,
        logoUrl: true,
        potentialValue: true,
        pipelineStage: true,
        updatedAt: true,
      },
    }),
    db.client.count({
      where: { status: "ACTIVE", wonAt: { gte: new Date(new Date().getFullYear(), 0, 1) } },
    }),
    db.client.count({ where: { status: "LOST" } }),
  ]);

  const cards: ProspectCard[] = prospects.map((prospect) => ({
    id: prospect.id,
    company: prospect.company,
    contact: [prospect.firstName, prospect.lastName].filter(Boolean).join(" ") || null,
    city: prospect.city,
    logoUrl: prospect.logoUrl,
    potentialValue: toNumber(prospect.potentialValue),
    stage: prospect.pipelineStage ?? "IDENTIFIED",
    updatedAt: prospect.updatedAt.toISOString(),
  }));

  const pipelineValue = cards.reduce((sum, card) => sum + card.potentialValue, 0);
  const engaged = cards.filter((card) =>
    ["MEETING", "QUOTE_SENT", "NEGOTIATION"].includes(card.stage),
  ).length;

  const conversionRate =
    wonThisYear + lostCount > 0 ? Math.round((wonThisYear / (wonThisYear + lostCount)) * 100) : null;

  return (
    <div className="space-y-4">
      <PageHeader
        title="Prospection"
        description="Les entreprises sans site web que vous démarchez. Glissez une fiche pour la faire avancer."
        actions={
          <ButtonLink href="/clients/nouveau" variant="primary">
            <Plus className="h-3.5 w-3.5" />
            Nouveau prospect
          </ButtonLink>
        }
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile label="Prospects actifs" value={cards.length} />
        <StatTile
          label="Potentiel annuel"
          value={formatMoney(pipelineValue)}
          hint="Somme des estimations saisies"
          tone="accent"
        />
        <StatTile
          label="En discussion avancée"
          value={engaged}
          hint="Rendez-vous, devis ou négociation"
          tone="info"
        />
        <StatTile
          label="Convertis cette année"
          value={wonThisYear}
          hint={conversionRate !== null ? `${conversionRate}% de taux de conversion` : undefined}
          tone="success"
        />
      </div>

      {cards.length === 0 ? (
        <Card>
          <EmptyState
            icon={<Target className="h-4 w-4" />}
            title="Aucun prospect en cours"
            description="Créez une fiche au statut « Prospect » pour la faire apparaître dans le pipeline. Chaque entreprise sans site web est une piste."
            action={
              <ButtonLink href="/clients/nouveau" variant="primary">
                Ajouter un prospect
              </ButtonLink>
            }
          />
        </Card>
      ) : (
        <PipelineBoard prospects={cards} />
      )}

      <p className="text-center text-[11px] text-ink-muted">
        Étapes du pipeline :{" "}
        {Object.values(PIPELINE_STAGE)
          .map((stage) => stage.label)
          .join(" → ")}
      </p>
    </div>
  );
}
