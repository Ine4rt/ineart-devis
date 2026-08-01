"use client";

import type { Quote, QuoteItem } from "@prisma/client";
import { Plus, Trash2 } from "lucide-react";
import Link from "next/link";
import { useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import {
  EnumSelect,
  Field,
  FormActions,
  FormGrid,
  FormSection,
  Input,
  Select,
  Textarea,
} from "@/components/ui/form";
import { SubmitButton } from "@/components/ui/submit-button";
import { QUOTE_STATUS_LIST } from "@/lib/constants";
import { computeQuoteTotals } from "@/lib/quotes";
import { formatMoney, toDateInputValue, toNumber } from "@/lib/format";

/**
 * Éditeur de devis.
 *
 * Les lignes sont un état local (on en ajoute et en retire librement) mais
 * l'envoi reste un `FormData` classique : chaque ligne rend des champs nommés
 * `itemLabel`, `itemQuantity`… que la server action recompose. Les totaux sont
 * recalculés à la frappe avec la même fonction pure que le serveur, donc
 * l'aperçu ne peut pas diverger du montant enregistré.
 */

interface LineDraft {
  key: string;
  label: string;
  description: string;
  quantity: string;
  unitPrice: string;
}

function emptyLine(): LineDraft {
  return {
    key: Math.random().toString(36).slice(2),
    label: "",
    description: "",
    quantity: "1",
    unitPrice: "",
  };
}

export function QuoteEditor({
  quote,
  items,
  clients,
  action,
  submitLabel,
  cancelHref,
  defaultClientId,
}: {
  quote?: Quote | null;
  items?: QuoteItem[];
  clients: { id: string; company: string }[];
  action: (formData: FormData) => Promise<void>;
  submitLabel: string;
  cancelHref: string;
  defaultClientId?: string;
}) {
  const [lines, setLines] = useState<LineDraft[]>(() =>
    items && items.length > 0
      ? items.map((item) => ({
          key: item.id,
          label: item.label,
          description: item.description ?? "",
          quantity: String(toNumber(item.quantity)),
          unitPrice: String(toNumber(item.unitPrice)),
        }))
      : [emptyLine()],
  );
  const [discountPct, setDiscountPct] = useState(String(toNumber(quote?.discountPct) || 0));
  const [taxRate, setTaxRate] = useState(String(quote ? toNumber(quote.taxRate) : 21));

  const totals = useMemo(
    () =>
      computeQuoteTotals(
        lines.map((line) => ({
          quantity: Number.parseFloat(line.quantity.replace(",", ".")) || 0,
          unitPrice: Number.parseFloat(line.unitPrice.replace(",", ".")) || 0,
        })),
        Number.parseFloat(discountPct.replace(",", ".")) || 0,
        Number.parseFloat(taxRate.replace(",", ".")) || 0,
      ),
    [lines, discountPct, taxRate],
  );

  function updateLine(key: string, patch: Partial<LineDraft>) {
    setLines((current) =>
      current.map((line) => (line.key === key ? { ...line, ...patch } : line)),
    );
  }

  return (
    <form action={action}>
      <Card className="overflow-hidden">
        <FormSection title="Devis" description="Client, objet et validité de la proposition.">
          <FormGrid>
            <Field label="Client" htmlFor="clientId" required>
              <Select
                id="clientId"
                name="clientId"
                required
                defaultValue={quote?.clientId ?? defaultClientId ?? ""}
              >
                <option value="">Sélectionner un client</option>
                {clients.map((client) => (
                  <option key={client.id} value={client.id}>
                    {client.company}
                  </option>
                ))}
              </Select>
            </Field>
            <Field label="Statut" htmlFor="status">
              <EnumSelect
                id="status"
                name="status"
                options={QUOTE_STATUS_LIST}
                defaultValue={quote?.status ?? "DRAFT"}
              />
            </Field>
            <Field label="Objet" htmlFor="title" required span={2}>
              <Input
                id="title"
                name="title"
                required
                data-autofocus
                defaultValue={quote?.title ?? ""}
                placeholder="Création d'un site vitrine + hébergement annuel"
              />
            </Field>
            <Field label="Date d'émission" htmlFor="issuedAt">
              <Input
                id="issuedAt"
                name="issuedAt"
                type="date"
                defaultValue={toDateInputValue(quote?.issuedAt ?? new Date())}
              />
            </Field>
            <Field label="Valable jusqu'au" htmlFor="validUntil">
              <Input
                id="validUntil"
                name="validUntil"
                type="date"
                defaultValue={toDateInputValue(quote?.validUntil)}
              />
            </Field>
            <Field label="Introduction" htmlFor="intro" span={2} hint="Apparaît en tête du devis.">
              <Textarea id="intro" name="intro" rows={3} defaultValue={quote?.intro ?? ""} />
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection title="Prestations" description="Une ligne par poste. Les lignes sans intitulé sont ignorées.">
          <div className="space-y-2">
            <div className="hidden gap-2 px-1 text-[11px] font-medium uppercase tracking-wide text-ink-muted sm:grid sm:grid-cols-[1fr_80px_110px_110px_32px]">
              <span>Intitulé</span>
              <span className="text-right">Qté</span>
              <span className="text-right">Prix unitaire</span>
              <span className="text-right">Total</span>
              <span />
            </div>

            {lines.map((line) => {
              const lineTotal =
                (Number.parseFloat(line.quantity.replace(",", ".")) || 0) *
                (Number.parseFloat(line.unitPrice.replace(",", ".")) || 0);

              return (
                <div
                  key={line.key}
                  className="grid gap-2 rounded-md border border-line bg-surface-subtle p-2 sm:grid-cols-[1fr_80px_110px_110px_32px] sm:items-start sm:border-0 sm:bg-transparent sm:p-0"
                >
                  <div className="space-y-1.5">
                    <Input
                      name="itemLabel"
                      value={line.label}
                      onChange={(event) => updateLine(line.key, { label: event.target.value })}
                      placeholder="Conception et développement du site"
                    />
                    <Input
                      name="itemDescription"
                      value={line.description}
                      onChange={(event) => updateLine(line.key, { description: event.target.value })}
                      placeholder="Détail (facultatif)"
                      className="text-xs"
                    />
                  </div>
                  <Input
                    name="itemQuantity"
                    inputMode="decimal"
                    value={line.quantity}
                    onChange={(event) => updateLine(line.key, { quantity: event.target.value })}
                    className="text-right"
                  />
                  <Input
                    name="itemUnitPrice"
                    inputMode="decimal"
                    value={line.unitPrice}
                    onChange={(event) => updateLine(line.key, { unitPrice: event.target.value })}
                    className="text-right"
                    placeholder="0"
                  />
                  <div className="flex h-8 items-center justify-end px-2 text-[13px] font-medium text-ink tabular-nums">
                    {formatMoney(lineTotal)}
                  </div>
                  <Button
                    type="button"
                    variant="ghost"
                    size="icon"
                    aria-label="Supprimer la ligne"
                    onClick={() =>
                      setLines((current) =>
                        current.length > 1
                          ? current.filter((candidate) => candidate.key !== line.key)
                          : [emptyLine()],
                      )
                    }
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                  </Button>
                </div>
              );
            })}

            <Button
              type="button"
              variant="ghost"
              onClick={() => setLines((current) => [...current, emptyLine()])}
            >
              <Plus className="h-3.5 w-3.5" />
              Ajouter une ligne
            </Button>
          </div>

          <div className="mt-5 grid gap-4 sm:grid-cols-2">
            <FormGrid columns={1}>
              <Field label="Remise (%)" htmlFor="discountPct">
                <Input
                  id="discountPct"
                  name="discountPct"
                  inputMode="decimal"
                  value={discountPct}
                  onChange={(event) => setDiscountPct(event.target.value)}
                />
              </Field>
              <Field label="TVA (%)" htmlFor="taxRate">
                <Input
                  id="taxRate"
                  name="taxRate"
                  inputMode="decimal"
                  value={taxRate}
                  onChange={(event) => setTaxRate(event.target.value)}
                />
              </Field>
            </FormGrid>

            <div className="card-inset space-y-1.5 p-3 text-[13px]">
              <div className="flex justify-between text-ink-secondary">
                <span>Sous-total</span>
                <span className="tabular-nums">{formatMoney(totals.subtotal)}</span>
              </div>
              {totals.discount > 0 ? (
                <div className="flex justify-between text-ink-secondary">
                  <span>Remise</span>
                  <span className="tabular-nums">−{formatMoney(totals.discount)}</span>
                </div>
              ) : null}
              <div className="flex justify-between text-ink-secondary">
                <span>TVA</span>
                <span className="tabular-nums">{formatMoney(totals.tax)}</span>
              </div>
              <div className="mt-1 flex justify-between border-t border-line pt-2 text-sm font-semibold text-ink">
                <span>Total TTC</span>
                <span className="tabular-nums">{formatMoney(totals.total)}</span>
              </div>
            </div>
          </div>
        </FormSection>

        <FormSection title="Conditions" description="Modalités de paiement, délais, mentions légales.">
          <FormGrid columns={1}>
            <Field label="Conditions" htmlFor="terms">
              <Textarea
                id="terms"
                name="terms"
                rows={4}
                defaultValue={
                  quote?.terms ??
                  "Acompte de 30 % à la commande, solde à la livraison. Abonnement annuel facturé à la mise en ligne."
                }
              />
            </Field>
            <Field label="Notes internes" htmlFor="notes" hint="Non imprimées sur le devis.">
              <Textarea id="notes" name="notes" rows={2} defaultValue={quote?.notes ?? ""} />
            </Field>
          </FormGrid>
        </FormSection>

        <FormActions>
          <Link href={cancelHref} className="btn btn-ghost h-8 px-3">
            Annuler
          </Link>
          <SubmitButton pendingLabel="Enregistrement…">{submitLabel}</SubmitButton>
        </FormActions>
      </Card>
    </form>
  );
}
