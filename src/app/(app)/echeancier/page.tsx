import { CalendarClock } from "lucide-react";
import type { Metadata } from "next";

import { RenewalRow, UrgencySummary } from "@/components/domain/renewal-list";
import { SegmentedLinks } from "@/components/ui/tabs";
import { Card, CardHeader, PageHeader } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { StatTile } from "@/components/ui/stat-tile";
import { formatMoney } from "@/lib/format";
import { getRenewalFeed, RENEWAL_KIND_LABEL, type RenewalKind } from "@/lib/queries/renewal-feed";
import { ALERT_THRESHOLDS, urgencyFromDays } from "@/lib/renewals";

export const metadata: Metadata = { title: "Échéancier" };
export const dynamic = "force-dynamic";

/**
 * Vue chronologique unique de tout ce qui expire.
 *
 * Les éléments sont regroupés par palier d'alerte (retard, 7, 15, 30, 60,
 * 90 jours, au-delà) plutôt que listés à plat : on veut voir d'abord ce qui
 * brûle, pas le premier élément par ordre de date.
 */
export default async function SchedulePage({
  searchParams,
}: {
  searchParams: Promise<{ type?: string; auto?: string; horizon?: string }>;
}) {
  const params = await searchParams;
  const horizon = Number(params.horizon) || 365;

  let items = await getRenewalFeed(horizon);

  if (params.type && params.type in RENEWAL_KIND_LABEL) {
    items = items.filter((item) => item.kind === (params.type as RenewalKind));
  }
  if (params.auto === "manuel") {
    items = items.filter((item) => !item.autoRenew);
  }

  const buckets: { label: string; max: number; items: typeof items }[] = [
    { label: "En retard", max: -1, items: [] },
    { label: "Sous 7 jours", max: 7, items: [] },
    { label: "Sous 15 jours", max: 15, items: [] },
    { label: "Sous 30 jours", max: 30, items: [] },
    { label: "Sous 60 jours", max: 60, items: [] },
    { label: "Sous 90 jours", max: 90, items: [] },
    { label: "Au-delà de 90 jours", max: Number.POSITIVE_INFINITY, items: [] },
  ];

  for (const item of items) {
    const bucket =
      item.days < 0
        ? buckets[0]
        : (buckets.find((candidate) => candidate.max >= item.days && candidate.max > 0) ??
          buckets[buckets.length - 1]);
    bucket.items.push(item);
  }

  const incomingRevenue = items
    .filter((item) => item.isRevenue && item.days <= 90)
    .reduce((sum, item) => sum + item.amount, 0);
  const outgoingCost = items
    .filter((item) => !item.isRevenue && item.days <= 90)
    .reduce((sum, item) => sum + item.amount, 0);
  const manualCount = items.filter((item) => !item.autoRenew).length;

  const typeViews = [
    { href: "/echeancier", label: "Tout", key: undefined },
    { href: "/echeancier?type=subscription", label: "Abonnements", key: "subscription" },
    { href: "/echeancier?type=domain", label: "Domaines", key: "domain" },
    { href: "/echeancier?type=hosting", label: "Hébergements", key: "hosting" },
    { href: "/echeancier?type=ssl", label: "SSL", key: "ssl" },
  ];

  return (
    <div className="space-y-4">
      <PageHeader
        title="Échéancier"
        description="Abonnements, domaines, hébergements et certificats réunis sur une seule ligne de temps."
      />

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile label="Échéances suivies" value={items.length} />
        <StatTile
          label="Revenus attendus (90 j)"
          value={formatMoney(incomingRevenue)}
          tone="success"
        />
        <StatTile label="Coûts à engager (90 j)" value={formatMoney(outgoingCost)} tone="warning" />
        <StatTile
          label="Reconductions manuelles"
          value={manualCount}
          tone={manualCount > 0 ? "caution" : "neutral"}
          hint="Expireront sans action de votre part"
          href="/echeancier?auto=manuel"
        />
      </div>

      <Card>
        <CardHeader
          title="Répartition par palier"
          description={`Seuils de rappel : ${ALERT_THRESHOLDS.join(" / ")} jours`}
          icon={<CalendarClock className="h-4 w-4" />}
        />
        <div className="p-4">
          <UrgencySummary items={items} />
        </div>
      </Card>

      <div className="flex flex-wrap items-center gap-3">
        <SegmentedLinks
          items={typeViews.map((view) => ({
            href: view.href,
            label: view.label,
            active: params.type === view.key,
          }))}
        />
        <SegmentedLinks
          items={[
            { href: "/echeancier", label: "Toutes", active: params.auto !== "manuel" },
            { href: "/echeancier?auto=manuel", label: "Manuelles", active: params.auto === "manuel" },
          ]}
        />
      </div>

      {items.length === 0 ? (
        <Card>
          <EmptyState
            icon={<CalendarClock className="h-4 w-4" />}
            title="Aucune échéance"
            description="Renseignez les dates d'expiration de vos domaines, hébergements et abonnements pour les voir apparaître ici."
          />
        </Card>
      ) : (
        <div className="space-y-4">
          {buckets
            .filter((bucket) => bucket.items.length > 0)
            .map((bucket) => {
              const tone = urgencyFromDays(bucket.max < 0 ? -1 : bucket.max).tone;
              return (
                <Card key={bucket.label}>
                  <div
                    data-tone={tone}
                    className="flex items-center justify-between gap-3 border-b border-line px-4 py-2.5"
                  >
                    <div className="flex items-center gap-2">
                      <span className="dot" />
                      <h2 className="text-sm font-semibold text-ink">{bucket.label}</h2>
                    </div>
                    <span className="text-xs text-ink-muted tabular-nums">
                      {bucket.items.length} élément(s)
                    </span>
                  </div>
                  <div className="divide-y divide-line">
                    {bucket.items.map((item) => (
                      <RenewalRow key={item.id} item={item} />
                    ))}
                  </div>
                </Card>
              );
            })}
        </div>
      )}
    </div>
  );
}
