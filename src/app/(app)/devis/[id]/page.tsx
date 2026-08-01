import type { QuoteStatus } from "@prisma/client";
import { Check, Pencil, Send, X } from "lucide-react";
import Link from "next/link";
import { notFound } from "next/navigation";

import { EnumBadge } from "@/components/ui/badge";
import { Button, ButtonLink } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { PrintButton } from "@/components/ui/print-button";
import { QUOTE_STATUS } from "@/lib/constants";
import { setQuoteStatus } from "@/lib/actions/quotes";
import { db } from "@/lib/db";
import { formatDate, formatMoney, toNumber } from "@/lib/format";
import { computeQuoteTotals } from "@/lib/quotes";

export const dynamic = "force-dynamic";

/**
 * Aperçu d'un devis, mis en page pour l'impression.
 *
 * Pas de génération PDF côté serveur : l'impression du navigateur (Ctrl+P →
 * « Enregistrer au format PDF ») produit exactement ce document, avec les
 * styles `@media print` définis dans globals.css. Une dépendance de moins, et
 * un rendu qui reste fidèle à ce qu'on voit à l'écran.
 */
export default async function QuotePage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;

  const quote = await db.quote.findUnique({
    where: { id },
    include: {
      client: true,
      items: { orderBy: { position: "asc" } },
    },
  });

  if (!quote) notFound();

  const totals = computeQuoteTotals(
    quote.items.map((item) => ({
      quantity: toNumber(item.quantity),
      unitPrice: toNumber(item.unitPrice),
    })),
    toNumber(quote.discountPct),
    toNumber(quote.taxRate),
  );

  const transitions: { status: QuoteStatus; label: string; icon: typeof Send }[] = [
    { status: "SENT", label: "Marquer envoyé", icon: Send },
    { status: "ACCEPTED", label: "Accepté", icon: Check },
    { status: "REJECTED", label: "Refusé", icon: X },
  ];

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <header className="no-print flex flex-wrap items-start justify-between gap-4">
        <div className="min-w-0">
          <p className="section-label">
            <Link href={`/clients/${quote.client.id}`} className="hover:text-accent-text">
              {quote.client.company}
            </Link>
          </p>
          <div className="mt-1 flex flex-wrap items-center gap-2">
            <h1 className="text-xl font-semibold text-ink">{quote.number}</h1>
            <EnumBadge meta={QUOTE_STATUS[quote.status]} />
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {transitions
            .filter((transition) => transition.status !== quote.status)
            .map((transition) => {
              const apply = setQuoteStatus.bind(null, quote.id, transition.status);
              const Icon = transition.icon;
              return (
                <form action={apply} key={transition.status}>
                  <Button type="submit" size="sm">
                    <Icon className="h-3.5 w-3.5" />
                    {transition.label}
                  </Button>
                </form>
              );
            })}
          <ButtonLink href={`/devis/${quote.id}/modifier`} size="sm">
            <Pencil className="h-3.5 w-3.5" />
            Modifier
          </ButtonLink>
          <PrintButton />
        </div>
      </header>

      {/* --- Document imprimable --- */}
      <Card className="p-8 sm:p-10">
        <div className="flex flex-wrap items-start justify-between gap-6 border-b border-line pb-6">
          <div>
            <div className="flex items-center gap-2.5">
              <span className="flex h-9 w-9 items-center justify-center rounded-md bg-[var(--accent)] text-sm font-bold text-white">
                i4
              </span>
              <div>
                <p className="text-sm font-semibold text-ink">Ine4rt</p>
                <p className="text-xs text-ink-muted">Développement web indépendant</p>
              </div>
            </div>
          </div>

          <div className="text-right">
            <p className="text-lg font-semibold text-ink">Devis {quote.number}</p>
            <p className="mt-0.5 text-xs text-ink-muted">
              Émis le {formatDate(quote.issuedAt, "long")}
            </p>
            {quote.validUntil ? (
              <p className="text-xs text-ink-muted">
                Valable jusqu'au {formatDate(quote.validUntil, "long")}
              </p>
            ) : null}
          </div>
        </div>

        <div className="grid gap-6 py-6 sm:grid-cols-2">
          <div>
            <p className="section-label">Destinataire</p>
            <p className="mt-1.5 text-sm font-semibold text-ink">{quote.client.company}</p>
            {[quote.client.firstName, quote.client.lastName].filter(Boolean).length > 0 ? (
              <p className="text-sm text-ink-secondary">
                {[quote.client.firstName, quote.client.lastName].filter(Boolean).join(" ")}
              </p>
            ) : null}
            {quote.client.addressLine1 ? (
              <p className="mt-1 text-xs leading-relaxed text-ink-secondary">
                {quote.client.addressLine1}
                {quote.client.addressLine2 ? (
                  <>
                    <br />
                    {quote.client.addressLine2}
                  </>
                ) : null}
                <br />
                {[quote.client.postalCode, quote.client.city].filter(Boolean).join(" ")}
                {quote.client.country ? (
                  <>
                    <br />
                    {quote.client.country}
                  </>
                ) : null}
              </p>
            ) : null}
            {quote.client.vatNumber ? (
              <p className="mt-1 text-xs text-ink-muted">TVA {quote.client.vatNumber}</p>
            ) : null}
          </div>

          <div className="sm:text-right">
            <p className="section-label">Objet</p>
            <p className="mt-1.5 text-sm text-ink">{quote.title}</p>
          </div>
        </div>

        {quote.intro ? (
          <p className="whitespace-pre-wrap border-t border-line pt-5 text-sm leading-relaxed text-ink-secondary">
            {quote.intro}
          </p>
        ) : null}

        <table className="mt-6 w-full text-sm">
          <thead>
            <tr className="border-b border-line">
              <th className="pb-2 text-left text-[11px] font-medium uppercase tracking-wide text-ink-muted">
                Prestation
              </th>
              <th className="pb-2 text-right text-[11px] font-medium uppercase tracking-wide text-ink-muted">
                Qté
              </th>
              <th className="pb-2 text-right text-[11px] font-medium uppercase tracking-wide text-ink-muted">
                P.U.
              </th>
              <th className="pb-2 text-right text-[11px] font-medium uppercase tracking-wide text-ink-muted">
                Total
              </th>
            </tr>
          </thead>
          <tbody>
            {quote.items.map((item) => (
              <tr key={item.id} className="border-b border-line">
                <td className="py-2.5 pr-4">
                  <p className="font-medium text-ink">{item.label}</p>
                  {item.description ? (
                    <p className="mt-0.5 text-xs text-ink-muted">{item.description}</p>
                  ) : null}
                </td>
                <td className="py-2.5 text-right text-ink-secondary tabular-nums">
                  {toNumber(item.quantity)}
                </td>
                <td className="py-2.5 text-right text-ink-secondary tabular-nums">
                  {formatMoney(item.unitPrice, { currency: quote.currency })}
                </td>
                <td className="py-2.5 text-right font-medium text-ink tabular-nums">
                  {formatMoney(toNumber(item.quantity) * toNumber(item.unitPrice), {
                    currency: quote.currency,
                  })}
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        <div className="mt-5 flex justify-end">
          <div className="w-full max-w-xs space-y-1.5 text-sm">
            <div className="flex justify-between text-ink-secondary">
              <span>Sous-total</span>
              <span className="tabular-nums">
                {formatMoney(totals.subtotal, { currency: quote.currency })}
              </span>
            </div>
            {totals.discount > 0 ? (
              <div className="flex justify-between text-ink-secondary">
                <span>Remise ({toNumber(quote.discountPct)} %)</span>
                <span className="tabular-nums">
                  −{formatMoney(totals.discount, { currency: quote.currency })}
                </span>
              </div>
            ) : null}
            <div className="flex justify-between text-ink-secondary">
              <span>Total HT</span>
              <span className="tabular-nums">
                {formatMoney(totals.taxable, { currency: quote.currency })}
              </span>
            </div>
            <div className="flex justify-between text-ink-secondary">
              <span>TVA ({toNumber(quote.taxRate)} %)</span>
              <span className="tabular-nums">
                {formatMoney(totals.tax, { currency: quote.currency })}
              </span>
            </div>
            <div className="flex justify-between border-t border-line pt-2 text-base font-semibold text-ink">
              <span>Total TTC</span>
              <span className="tabular-nums">
                {formatMoney(totals.total, { currency: quote.currency })}
              </span>
            </div>
          </div>
        </div>

        {quote.terms ? (
          <div className="mt-8 border-t border-line pt-5">
            <p className="section-label">Conditions</p>
            <p className="mt-1.5 whitespace-pre-wrap text-xs leading-relaxed text-ink-secondary">
              {quote.terms}
            </p>
          </div>
        ) : null}
      </Card>

      {quote.notes ? (
        <Card className="no-print p-4">
          <p className="section-label">Notes internes</p>
          <p className="mt-1.5 whitespace-pre-wrap text-sm text-ink-secondary">{quote.notes}</p>
        </Card>
      ) : null}
    </div>
  );
}

export async function generateMetadata({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const quote = await db.quote.findUnique({ where: { id }, select: { number: true } });
  return { title: quote?.number ?? "Devis" };
}
