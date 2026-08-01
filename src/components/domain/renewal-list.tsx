import { AlertTriangle, Globe, RefreshCw, Server, ShieldCheck, Waypoints } from "lucide-react";
import Link from "next/link";

import { Badge } from "@/components/ui/badge";
import { formatDate, formatMoney } from "@/lib/format";
import { RENEWAL_KIND_LABEL, type RenewalItem, type RenewalKind } from "@/lib/queries/renewal-feed";
import { countdownLabel } from "@/lib/renewals";
import { cn } from "@/lib/utils";

/**
 * Rendu d'une échéance. Partagé par le tableau de bord, l'échéancier et les
 * fiches client : une échéance se lit exactement de la même façon partout.
 */

const KIND_ICON: Record<RenewalKind, typeof Globe> = {
  subscription: Waypoints,
  domain: Globe,
  hosting: Server,
  ssl: ShieldCheck,
};

export function RenewalRow({ item, showKind = true }: { item: RenewalItem; showKind?: boolean }) {
  const Icon = KIND_ICON[item.kind];

  return (
    <Link
      href={item.href}
      className="group flex items-center gap-3 px-4 py-2.5 transition-colors hover:bg-surface-inset"
    >
      <span
        data-tone={item.urgency.tone}
        className="flex h-7 w-7 shrink-0 items-center justify-center rounded-md border border-[color:var(--tone-line)] bg-[color:var(--tone-soft)] text-[color:var(--tone-fg)]"
      >
        <Icon className="h-3.5 w-3.5" />
      </span>

      <div className="min-w-0 flex-1">
        <div className="flex items-center gap-2">
          <span className="truncate text-[13px] font-medium text-ink">{item.label}</span>
          {!item.autoRenew ? (
            <span
              title="Reconduction automatique désactivée — action manuelle requise"
              data-tone="warning"
              className="flex shrink-0 items-center text-[color:var(--tone-fg)]"
            >
              <AlertTriangle className="h-3 w-3" />
            </span>
          ) : null}
        </div>
        <div className="truncate text-xs text-ink-muted">
          {showKind ? `${RENEWAL_KIND_LABEL[item.kind]} · ` : ""}
          {item.sublabel}
        </div>
      </div>

      <div className="shrink-0 text-right">
        <div
          data-tone={item.urgency.tone}
          className="text-xs font-medium text-[color:var(--tone-fg)] tabular-nums"
        >
          {countdownLabel(item.days)}
        </div>
        <div className="text-[11px] text-ink-muted tabular-nums">{formatDate(item.date, "short")}</div>
      </div>

      {item.amount > 0 ? (
        <div className="w-20 shrink-0 text-right">
          <div
            className={cn(
              "text-xs font-medium tabular-nums",
              item.isRevenue ? "text-ink" : "text-ink-muted",
            )}
          >
            {item.isRevenue ? "+" : "−"}
            {formatMoney(item.amount, { currency: item.currency, compact: true })}
          </div>
          <div className="text-[11px] text-ink-muted">{item.isRevenue ? "revenu" : "coût"}</div>
        </div>
      ) : (
        <div className="w-20 shrink-0" />
      )}
    </Link>
  );
}

/** Compteurs par palier — la synthèse « 90/60/30/15/7 jours » demandée. */
export function UrgencySummary({ items }: { items: RenewalItem[] }) {
  const buckets = [
    { label: "En retard", tone: "danger" as const, count: items.filter((i) => i.days < 0).length },
    { label: "7 jours", tone: "critical" as const, count: items.filter((i) => i.days >= 0 && i.days <= 7).length },
    { label: "15 jours", tone: "warning" as const, count: items.filter((i) => i.days > 7 && i.days <= 15).length },
    { label: "30 jours", tone: "warning" as const, count: items.filter((i) => i.days > 15 && i.days <= 30).length },
    { label: "60 jours", tone: "caution" as const, count: items.filter((i) => i.days > 30 && i.days <= 60).length },
    { label: "90 jours", tone: "info" as const, count: items.filter((i) => i.days > 60 && i.days <= 90).length },
  ];

  return (
    <div className="grid grid-cols-3 gap-px overflow-hidden rounded-md border border-line bg-line sm:grid-cols-6">
      {buckets.map((bucket) => (
        <div key={bucket.label} data-tone={bucket.tone} className="bg-surface px-3 py-2.5">
          <div
            className={cn(
              "text-lg font-semibold tabular-nums",
              bucket.count > 0 ? "text-[color:var(--tone-fg)]" : "text-ink-muted",
            )}
          >
            {bucket.count}
          </div>
          <div className="mt-0.5 truncate text-[11px] text-ink-muted">{bucket.label}</div>
        </div>
      ))}
    </div>
  );
}

export function AutoRenewBadge({ autoRenew }: { autoRenew: boolean }) {
  return autoRenew ? (
    <Badge tone="success" dot={false}>
      <RefreshCw className="h-3 w-3" />
      Auto
    </Badge>
  ) : (
    <Badge tone="warning" dot={false}>
      <AlertTriangle className="h-3 w-3" />
      Manuel
    </Badge>
  );
}
