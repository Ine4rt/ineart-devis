import { CalendarClock, Check, Pencil, Plus, RefreshCw, Trash2 } from "lucide-react";
import Link from "next/link";
import { notFound } from "next/navigation";

import { AutoRenewBadge } from "@/components/domain/renewal-list";
import { EnumBadge } from "@/components/ui/badge";
import { Button, ButtonLink } from "@/components/ui/button";
import { Card, CardBody, CardHeader } from "@/components/ui/card";
import { DetailList, DetailRow } from "@/components/ui/detail-list";
import { EmptyState } from "@/components/ui/empty-state";
import { EnumSelect, Field, Input } from "@/components/ui/form";
import { SubmitButton } from "@/components/ui/submit-button";
import {
  BILLING_CYCLE,
  PERIOD_STATUS,
  PERIOD_STATUS_LIST,
  SUBSCRIPTION_STATUS,
  SUBSCRIPTION_TYPE,
} from "@/lib/constants";
import {
  createPeriod,
  deletePeriod,
  markPeriodPaid,
  renewSubscription,
} from "@/lib/actions/subscriptions";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { countdownLabel, daysUntil, toYearly, urgencyFor } from "@/lib/renewals";

export const dynamic = "force-dynamic";

export default async function SubscriptionPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;

  const subscription = await db.subscription.findUnique({
    where: { id },
    include: {
      client: { select: { id: true, company: true } },
      project: { select: { id: true, name: true } },
      periods: { orderBy: { periodStart: "desc" } },
    },
  });

  if (!subscription) notFound();

  const urgency = urgencyFor(subscription.nextRenewalAt);
  const days = daysUntil(subscription.nextRenewalAt);

  const paid = subscription.periods.filter((period) => period.status === "PAID");
  const unpaid = subscription.periods.filter((period) =>
    ["DUE", "OVERDUE"].includes(period.status),
  );
  const collected = paid.reduce((sum, period) => sum + toNumber(period.amount), 0);
  const outstanding = unpaid.reduce((sum, period) => sum + toNumber(period.amount), 0);

  const renew = renewSubscription.bind(null, subscription.id);
  const addPeriod = createPeriod.bind(null, subscription.id);

  return (
    <div className="space-y-4">
      <header className="flex flex-wrap items-start justify-between gap-4">
        <div className="min-w-0">
          <p className="section-label">
            <Link href={`/clients/${subscription.client.id}`} className="hover:text-accent-text">
              {subscription.client.company}
            </Link>
          </p>
          <div className="mt-1 flex flex-wrap items-center gap-2">
            <h1 className="text-xl font-semibold text-ink">{subscription.name}</h1>
            <EnumBadge meta={SUBSCRIPTION_STATUS[subscription.status]} />
            <AutoRenewBadge autoRenew={subscription.autoRenew} />
          </div>
        </div>

        <div className="flex items-center gap-2">
          <form action={renew}>
            <SubmitButton variant="default" pendingLabel="Reconduction…">
              <RefreshCw className="h-3.5 w-3.5" />
              Reconduire d'un cycle
            </SubmitButton>
          </form>
          <ButtonLink href={`/abonnements/${subscription.id}/modifier`}>
            <Pencil className="h-3.5 w-3.5" />
            Modifier
          </ButtonLink>
        </div>
      </header>

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <div className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Montant</p>
          <p className="mt-1.5 text-xl font-semibold text-ink tabular-nums">
            {formatMoney(subscription.amount, { currency: subscription.currency })}
          </p>
          <p className="mt-0.5 text-xs text-ink-muted">
            {BILLING_CYCLE[subscription.billingCycle].label} ·{" "}
            {formatMoney(toYearly(toNumber(subscription.amount), subscription.billingCycle), {
              currency: subscription.currency,
            })}
            /an
          </p>
        </div>

        <div data-tone={urgency.tone} className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Prochain renouvellement</p>
          <p className="mt-1.5 text-xl font-semibold text-ink tabular-nums">
            {formatDate(subscription.nextRenewalAt, "short")}
          </p>
          <p className="mt-0.5 text-xs font-medium text-[color:var(--tone-fg)]">
            {countdownLabel(days)}
          </p>
        </div>

        <div className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Encaissé à ce jour</p>
          <p className="mt-1.5 text-xl font-semibold text-ink tabular-nums">
            {formatMoney(collected, { currency: subscription.currency })}
          </p>
          <p className="mt-0.5 text-xs text-ink-muted">{paid.length} période(s) réglée(s)</p>
        </div>

        <div data-tone={outstanding > 0 ? "danger" : "success"} className="card p-4">
          <p className="text-xs font-medium text-ink-secondary">Reste à encaisser</p>
          <p className="mt-1.5 text-xl font-semibold text-ink tabular-nums">
            {formatMoney(outstanding, { currency: subscription.currency })}
          </p>
          <p className="mt-0.5 text-xs font-medium text-[color:var(--tone-fg)]">
            {unpaid.length === 0 ? "Aucun impayé" : `${unpaid.length} échéance(s) ouverte(s)`}
          </p>
        </div>
      </div>

      <div className="grid gap-4 lg:grid-cols-3">
        <Card className="lg:col-span-1">
          <CardHeader title="Contrat" />
          <CardBody className="pt-0">
            <DetailList>
              <DetailRow label="Type">{SUBSCRIPTION_TYPE[subscription.type].label}</DetailRow>
              <DetailRow label="Souscrit le">{formatDate(subscription.startedAt)}</DetailRow>
              <DetailRow label="Période en cours">
                {formatDate(subscription.currentPeriodStart, "short")} →{" "}
                {formatDate(subscription.currentPeriodEnd, "short")}
              </DetailRow>
              <DetailRow label="Moyen de paiement">{subscription.paymentMethod}</DetailRow>
              <DetailRow
                label="Projet"
                href={subscription.project ? `/projets/${subscription.project.id}` : null}
              >
                {subscription.project?.name}
              </DetailRow>
              <DetailRow label="Résilié le">{formatDate(subscription.cancelledAt)}</DetailRow>
              <DetailRow label="Remarques">
                {subscription.notes ? (
                  <span className="whitespace-pre-wrap">{subscription.notes}</span>
                ) : null}
              </DetailRow>
            </DetailList>
          </CardBody>
        </Card>

        <Card className="lg:col-span-2">
          <CardHeader
            title="Échéances"
            description="Chaque ligne couvre une période facturée. C'est l'historique de paiement du client."
            icon={<CalendarClock className="h-4 w-4" />}
          />

          <CardBody className="border-b border-line">
            <form action={addPeriod} className="flex flex-wrap items-end gap-2">
              <Field label="Début de période" htmlFor="periodStart">
                <Input id="periodStart" name="periodStart" type="date" required />
              </Field>
              <Field label="Échéance" htmlFor="dueAt">
                <Input id="dueAt" name="dueAt" type="date" />
              </Field>
              <Field label="Montant" htmlFor="amount" className="w-28">
                <Input
                  id="amount"
                  name="amount"
                  inputMode="decimal"
                  defaultValue={toNumber(subscription.amount)}
                />
              </Field>
              <Field label="Statut" htmlFor="status" className="w-32">
                <EnumSelect
                  id="status"
                  name="status"
                  options={PERIOD_STATUS_LIST}
                  defaultValue="DUE"
                />
              </Field>
              <SubmitButton variant="default">
                <Plus className="h-3.5 w-3.5" />
                Ajouter
              </SubmitButton>
            </form>
          </CardBody>

          <div className="divide-y divide-line">
            {subscription.periods.length > 0 ? (
              subscription.periods.map((period) => {
                const markPaid = markPeriodPaid.bind(null, period.id);
                const remove = deletePeriod.bind(null, period.id, subscription.id);
                const late = period.status === "OVERDUE";

                return (
                  <div key={period.id} className="flex flex-wrap items-center gap-3 px-4 py-2.5">
                    <div className="min-w-0 flex-1">
                      <p className="text-[13px] text-ink">
                        {formatDate(period.periodStart, "short")}
                        <span className="mx-1.5 text-ink-muted">→</span>
                        {formatDate(period.periodEnd, "short")}
                      </p>
                      <p className="text-xs text-ink-muted">
                        Échéance {formatDate(period.dueAt, "short")}
                        {period.paidAt ? ` · Payé le ${formatDate(period.paidAt, "short")}` : ""}
                        {late ? ` · ${countdownLabel(daysUntil(period.dueAt))}` : ""}
                      </p>
                    </div>

                    <EnumBadge meta={PERIOD_STATUS[period.status]} />

                    <span className="w-24 text-right text-sm font-medium text-ink tabular-nums">
                      {formatMoney(period.amount, { currency: subscription.currency })}
                    </span>

                    <div className="flex items-center gap-1">
                      {period.status !== "PAID" ? (
                        <form action={markPaid}>
                          <Button type="submit" size="sm" variant="default">
                            <Check className="h-3.5 w-3.5" />
                            Marquer payé
                          </Button>
                        </form>
                      ) : null}
                      <form action={remove}>
                        <Button variant="ghost" size="icon" type="submit" aria-label="Supprimer">
                          <Trash2 className="h-3.5 w-3.5" />
                        </Button>
                      </form>
                    </div>
                  </div>
                );
              })
            ) : (
              <EmptyState compact title="Aucune échéance" description="Ajoutez la première période." />
            )}
          </div>
        </Card>
      </div>
    </div>
  );
}

export async function generateMetadata({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const subscription = await db.subscription.findUnique({ where: { id }, select: { name: true } });
  return { title: subscription?.name ?? "Abonnement" };
}
