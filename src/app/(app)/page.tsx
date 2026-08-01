import {
  ArrowUpRight,
  Banknote,
  Building2,
  CalendarClock,
  CircleAlert,
  Globe,
  Info,
  ListChecks,
  Sparkles,
  TrendingUp,
  TriangleAlert,
  Wallet,
} from "lucide-react";
import Link from "next/link";

import { RenewalRow, UrgencySummary } from "@/components/domain/renewal-list";
import { AreaChart, BarChart, DonutChart } from "@/components/charts/chart-kit";
import { Badge } from "@/components/ui/badge";
import { Card, CardBody, CardHeader, CardLink } from "@/components/ui/card";
import { EmptyState } from "@/components/ui/empty-state";
import { StatTile } from "@/components/ui/stat-tile";
import { TASK_PRIORITY } from "@/lib/constants";
import { formatDate, formatMoney, formatPercent, formatRelative } from "@/lib/format";
import { getDashboardData } from "@/lib/queries/dashboard";
import { countdownLabel, daysUntil } from "@/lib/renewals";
import { requireUser } from "@/lib/auth";
import { cn } from "@/lib/utils";

export const dynamic = "force-dynamic";

const ALERT_ICON = {
  danger: TriangleAlert,
  warning: CircleAlert,
  info: Info,
} as const;

export default async function DashboardPage() {
  const [user, data] = await Promise.all([requireUser(), getDashboardData()]);

  const hour = data.now.getHours();
  const greeting = hour < 6 ? "Bonne nuit" : hour < 12 ? "Bonjour" : hour < 18 ? "Bon après-midi" : "Bonsoir";
  const firstName = user.name.split(" ")[0];

  const immediateRenewals = data.renewals.filter((item) => item.days <= 90).slice(0, 8);
  const hasData = data.clients.total > 0;

  return (
    <div className="space-y-5">
      <header className="flex flex-wrap items-end justify-between gap-4">
        <div>
          <p className="section-label">
            {new Intl.DateTimeFormat("fr-BE", {
              weekday: "long",
              day: "numeric",
              month: "long",
            }).format(data.now)}
          </p>
          <h1 className="mt-1 text-xl font-semibold text-ink">
            {greeting}, {firstName}
          </h1>
          <p className="mt-1 text-sm text-ink-secondary">
            {data.alerts.length === 0
              ? "Rien ne requiert votre attention immédiate."
              : `${data.alerts.length} point${data.alerts.length > 1 ? "s" : ""} demandent votre attention.`}
          </p>
        </div>

        <div className="flex items-center gap-2">
          <CardLink href="/echeancier">Voir l'échéancier</CardLink>
        </div>
      </header>

      {!hasData ? (
        <Card>
          <EmptyState
            icon={<Sparkles className="h-4 w-4" />}
            title="Votre console est prête"
            description="Commencez par créer un client, puis rattachez-lui un projet, un domaine et un abonnement. Les indicateurs se rempliront automatiquement."
            action={
              <Link href="/clients/nouveau" className="btn btn-primary h-8 px-3">
                Créer le premier client
              </Link>
            }
          />
        </Card>
      ) : null}

      {/* ---- Alertes : toujours en tête, jamais enfouies dans une carte ---- */}
      {data.alerts.length > 0 ? (
        <div className="stagger grid gap-2 sm:grid-cols-2 xl:grid-cols-4">
          {data.alerts.map((alert, index) => {
            const Icon = ALERT_ICON[alert.tone];
            return (
              <Link
                key={alert.id}
                href={alert.href}
                data-tone={alert.tone}
                style={{ "--i": index } as React.CSSProperties}
                className="group flex items-start gap-2.5 rounded-lg border border-[color:var(--tone-line)] bg-[color:var(--tone-soft)] px-3 py-2.5 transition-all duration-200 hover:-translate-y-px hover:shadow-[var(--shadow-sm)]"
              >
                <Icon className="mt-px h-4 w-4 shrink-0 text-[color:var(--tone-fg)]" />
                <span className="min-w-0 flex-1">
                  <span className="block text-xs font-semibold text-[color:var(--tone-fg)]">
                    {alert.title}
                  </span>
                  <span className="mt-0.5 block text-[11px] leading-relaxed text-ink-secondary">
                    {alert.detail}
                  </span>
                </span>
                <ArrowUpRight className="h-3.5 w-3.5 shrink-0 text-[color:var(--tone-fg)] opacity-0 transition-opacity group-hover:opacity-100" />
              </Link>
            );
          })}
        </div>
      ) : null}

      {/* ---- Indicateurs principaux ---- */}
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile
          label="Revenus annuels récurrents"
          value={formatMoney(data.revenue.arr)}
          hint={`${formatMoney(data.revenue.mrr)} / mois estimés`}
          icon={<TrendingUp className="h-4 w-4" />}
          tone="success"
          href="/revenus"
        />
        <StatTile
          label="Marge nette annuelle"
          value={formatMoney(data.revenue.margin)}
          hint={`${formatPercent(data.revenue.marginRate)} de marge · ${formatMoney(
            data.infrastructure.annualCost,
          )} de coûts`}
          icon={<Wallet className="h-4 w-4" />}
          tone="accent"
          href="/revenus"
        />
        <StatTile
          label="À encaisser"
          value={formatMoney(data.billing.unpaidTotal)}
          hint={
            data.billing.overdueCount > 0
              ? `dont ${formatMoney(data.billing.overdueTotal)} en retard`
              : `${data.billing.unpaidCount} échéance(s) ouverte(s)`
          }
          icon={<Banknote className="h-4 w-4" />}
          tone={data.billing.overdueCount > 0 ? "danger" : "neutral"}
          href="/paiements"
        />
        <StatTile
          label="Clients actifs"
          value={data.clients.active}
          hint={`${data.clients.total} au total · ${data.clients.prospects} prospect(s)`}
          icon={<Building2 className="h-4 w-4" />}
          href="/clients"
        />
      </div>

      {/* ---- Indicateurs secondaires ---- */}
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile
          label="Sites en ligne"
          value={data.projects.live}
          hint={`${data.projects.inProgress} en cours de production`}
          href="/projets"
        />
        <StatTile
          label="Noms de domaine"
          value={data.infrastructure.domains}
          hint={`${data.infrastructure.hostings} hébergement(s)`}
          href="/domaines"
        />
        <StatTile
          label="Abonnements actifs"
          value={data.billing.activeSubscriptions}
          hint={`${formatMoney(data.revenue.arr / Math.max(data.billing.activeSubscriptions, 1))} par contrat`}
          href="/abonnements"
        />
        <StatTile
          label="Pipeline commercial"
          value={data.pipeline.count}
          hint={
            data.pipeline.value > 0
              ? `${formatMoney(data.pipeline.value)} de potentiel annuel`
              : "Potentiel non estimé"
          }
          href="/prospection"
        />
      </div>

      <div className="grid gap-4 xl:grid-cols-3">
        {/* ---- Colonne principale ---- */}
        <div className="space-y-4 xl:col-span-2">
          <Card>
            <CardHeader
              title="Trajectoire des revenus récurrents"
              description="ARR reconstitué sur les douze derniers mois"
              icon={<TrendingUp className="h-4 w-4" />}
              action={
                <Badge tone={data.revenue.arr > 0 ? "success" : "muted"} dot={false}>
                  {formatMoney(data.revenue.arr, { compact: true })} aujourd'hui
                </Badge>
              }
            />
            <CardBody>
              {data.revenue.arr > 0 ? (
                <AreaChart
                  data={data.revenue.trend}
                  format="money"
                  height={190}
                />
              ) : (
                <EmptyState
                  compact
                  title="Aucun abonnement actif"
                  description="Les revenus récurrents apparaîtront dès le premier contrat enregistré."
                />
              )}
            </CardBody>
          </Card>

          <Card>
            <CardHeader
              title="Prochaines échéances"
              description="Abonnements, domaines, hébergements et certificats confondus"
              icon={<CalendarClock className="h-4 w-4" />}
              action={<CardLink href="/echeancier">Tout voir</CardLink>}
            />
            <CardBody className="pb-3">
              <UrgencySummary items={data.renewals} />
            </CardBody>
            <div className="divide-y divide-line border-t border-line">
              {immediateRenewals.length > 0 ? (
                immediateRenewals.map((item) => <RenewalRow key={item.id} item={item} />)
              ) : (
                <EmptyState
                  compact
                  title="Aucune échéance sous 90 jours"
                  description="Tout est à jour de ce côté."
                />
              )}
            </div>
          </Card>

          <Card>
            <CardHeader
              title="Encaissements de l'année"
              description={`${formatMoney(data.revenue.collectedThisYear)} encaissés depuis janvier`}
              icon={<Banknote className="h-4 w-4" />}
            />
            <CardBody>
              <BarChart
                data={data.revenue.monthlyCollected}
                format="money"
                highlightIndex={data.now.getMonth()}
                height={150}
              />
            </CardBody>
          </Card>
        </div>

        {/* ---- Colonne latérale ---- */}
        <div className="space-y-4">
          <Card>
            <CardHeader
              title="À encaisser"
              description={`${data.billing.unpaidCount} échéance(s) en attente`}
              action={<CardLink href="/paiements">Détail</CardLink>}
            />
            <div className="divide-y divide-line">
              {data.billing.upcomingUnpaid.length > 0 ? (
                data.billing.upcomingUnpaid.map((period) => {
                  const days = daysUntil(period.dueAt);
                  const late = days < 0;
                  return (
                    <Link
                      key={period.id}
                      href={`/clients/${period.subscription.client.id}`}
                      className="flex items-center gap-3 px-4 py-2.5 transition-colors hover:bg-surface-inset"
                    >
                      <div className="min-w-0 flex-1">
                        <div className="truncate text-[13px] font-medium text-ink">
                          {period.subscription.client.company}
                        </div>
                        <div className="truncate text-xs text-ink-muted">
                          {period.subscription.name}
                        </div>
                      </div>
                      <div className="shrink-0 text-right">
                        <div className="text-xs font-semibold text-ink tabular-nums">
                          {formatMoney(period.amount, { currency: period.subscription.currency })}
                        </div>
                        <div
                          data-tone={late ? "danger" : "warning"}
                          className="text-[11px] text-[color:var(--tone-fg)]"
                        >
                          {countdownLabel(days)}
                        </div>
                      </div>
                    </Link>
                  );
                })
              ) : (
                <EmptyState compact title="Tout est encaissé" description="Aucune échéance ouverte." />
              )}
            </div>
          </Card>

          <Card>
            <CardHeader
              title="Répartition de l'ARR"
              description="Par type d'abonnement"
              action={<CardLink href="/revenus">Analyser</CardLink>}
            />
            <CardBody>
              {data.revenue.mix.length > 0 ? (
                <DonutChart
                  data={data.revenue.mix}
                  format="money-compact"
                  centerLabel="ARR"
                  centerValue={formatMoney(data.revenue.arr, { compact: true })}
                  size={132}
                />
              ) : (
                <EmptyState compact title="Pas encore de revenus" />
              )}
            </CardBody>
          </Card>

          <Card>
            <CardHeader
              title="À faire"
              description={
                data.tasks.lateCount > 0 ? `${data.tasks.lateCount} en retard` : "Prochaines échéances"
              }
              icon={<ListChecks className="h-4 w-4" />}
              action={<CardLink href="/taches">Tout voir</CardLink>}
            />
            <div className="divide-y divide-line">
              {data.tasks.open.length > 0 ? (
                data.tasks.open.map((task) => {
                  const days = task.dueAt ? daysUntil(task.dueAt) : null;
                  return (
                    <Link
                      key={task.id}
                      href="/taches"
                      className="flex items-start gap-2.5 px-4 py-2.5 transition-colors hover:bg-surface-inset"
                    >
                      <span
                        data-tone={TASK_PRIORITY[task.priority].tone}
                        className="dot mt-1.5"
                        title={TASK_PRIORITY[task.priority].label}
                      />
                      <span className="min-w-0 flex-1">
                        <span className="block truncate text-[13px] text-ink">{task.title}</span>
                        <span className="block truncate text-xs text-ink-muted">
                          {task.client?.company ?? "Sans client"}
                        </span>
                      </span>
                      {days !== null ? (
                        <span
                          data-tone={days < 0 ? "danger" : days <= 3 ? "warning" : "muted"}
                          className="shrink-0 text-[11px] text-[color:var(--tone-fg)] tabular-nums"
                        >
                          {countdownLabel(days)}
                        </span>
                      ) : null}
                    </Link>
                  );
                })
              ) : (
                <EmptyState compact title="Aucune tâche ouverte" description="Rien en attente." />
              )}
            </div>
          </Card>

          <Card>
            <CardHeader title="Activité récente" />
            <div className="divide-y divide-line">
              {data.activity.length > 0 ? (
                data.activity.map((entry) => (
                  <div key={entry.id} className="px-4 py-2.5">
                    <p className="text-xs text-ink">
                      {entry.summary}
                      {entry.client ? (
                        <Link
                          href={`/clients/${entry.client.id}`}
                          className="ml-1 font-medium text-accent-text hover:underline"
                        >
                          {entry.client.company}
                        </Link>
                      ) : null}
                    </p>
                    <p
                      className="mt-0.5 text-[11px] text-ink-muted"
                      title={formatDate(entry.createdAt, "long")}
                    >
                      {formatRelative(entry.createdAt)}
                    </p>
                  </div>
                ))
              ) : (
                <EmptyState compact title="Aucune activité" description="L'historique se remplira au fil de vos actions." />
              )}
            </div>
          </Card>
        </div>
      </div>

      <p className={cn("pt-2 text-center text-[11px] text-ink-muted")}>
        <Globe className="mr-1 inline h-3 w-3" />
        Données locales — dernière actualisation {formatDate(data.now, "long")}
      </p>
    </div>
  );
}
