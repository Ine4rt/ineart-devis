"use server";

import type { QuoteStatus } from "@prisma/client";
import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";

import { requireUser } from "@/lib/auth";
import { QUOTE_STATUS } from "@/lib/constants";
import { db } from "@/lib/db";
import {
  date,
  dec,
  enumOf,
  logActivity,
  nextQuoteNumber,
  optStr,
  str,
} from "@/lib/actions/helpers";

const STATUSES = Object.keys(QUOTE_STATUS) as QuoteStatus[];

/**
 * Devis.
 *
 * Les lignes sont envoyées en tableaux parallèles (`items.label[]`,
 * `items.quantity[]`…) par un formulaire dynamique. On les recompose ici, en
 * ignorant les lignes sans intitulé — ce qui rend la suppression d'une ligne
 * côté client triviale : il suffit de vider le champ.
 */
function readItems(formData: FormData) {
  const labels = formData.getAll("itemLabel").map(String);
  const descriptions = formData.getAll("itemDescription").map(String);
  const quantities = formData.getAll("itemQuantity").map(String);
  const prices = formData.getAll("itemUnitPrice").map(String);

  return labels
    .map((label, index) => ({
      label: label.trim(),
      description: descriptions[index]?.trim() || null,
      quantity: Number.parseFloat((quantities[index] ?? "1").replace(",", ".")) || 1,
      unitPrice: Number.parseFloat((prices[index] ?? "0").replace(",", ".")) || 0,
      position: index,
    }))
    .filter((item) => item.label.length > 0);
}

function readQuoteFields(formData: FormData) {
  return {
    title: str(formData, "title") || "Proposition commerciale",
    status: enumOf(formData, "status", STATUSES, "DRAFT"),
    issuedAt: date(formData, "issuedAt") ?? new Date(),
    validUntil: date(formData, "validUntil"),
    discountPct: dec(formData, "discountPct"),
    taxRate: dec(formData, "taxRate", 21),
    currency: str(formData, "currency") || "EUR",
    intro: optStr(formData, "intro"),
    terms: optStr(formData, "terms"),
    notes: optStr(formData, "notes"),
  };
}

export async function createQuote(formData: FormData) {
  await requireUser();

  const clientId = str(formData, "clientId");
  if (!clientId) throw new Error("Un devis doit être rattaché à un client.");

  const quote = await db.quote.create({
    data: {
      ...readQuoteFields(formData),
      clientId,
      number: await nextQuoteNumber(),
      items: { create: readItems(formData) },
    },
  });

  await logActivity({
    entityType: "quote",
    entityId: quote.id,
    clientId,
    action: "created",
    summary: `Devis ${quote.number} créé pour`,
  });

  revalidatePath("/devis");
  redirect(`/devis/${quote.id}`);
}

export async function updateQuote(quoteId: string, formData: FormData) {
  await requireUser();

  const items = readItems(formData);

  // Remplacement intégral des lignes : plus simple et plus sûr qu'un diff, et
  // sans conséquence puisqu'une ligne de devis n'est référencée nulle part.
  await db.$transaction([
    db.quoteItem.deleteMany({ where: { quoteId } }),
    db.quote.update({
      where: { id: quoteId },
      data: {
        ...readQuoteFields(formData),
        items: { create: items },
      },
    }),
  ]);

  revalidatePath(`/devis/${quoteId}`);
  revalidatePath("/devis");
  redirect(`/devis/${quoteId}`);
}

export async function setQuoteStatus(quoteId: string, status: QuoteStatus) {
  await requireUser();
  if (!STATUSES.includes(status)) return;

  const quote = await db.quote.update({
    where: { id: quoteId },
    data: {
      status,
      sentAt: status === "SENT" ? new Date() : undefined,
      decidedAt: status === "ACCEPTED" || status === "REJECTED" ? new Date() : undefined,
    },
  });

  // Un devis accepté fait avancer le prospect : le pipeline suit la réalité
  // commerciale sans qu'on ait à le mettre à jour à la main.
  if (status === "ACCEPTED") {
    await db.client.updateMany({
      where: { id: quote.clientId, status: "PROSPECT" },
      data: { pipelineStage: "NEGOTIATION" },
    });
  }

  await logActivity({
    entityType: "quote",
    entityId: quoteId,
    clientId: quote.clientId,
    action: status.toLowerCase(),
    summary: `Devis ${quote.number} — ${QUOTE_STATUS[status].label.toLowerCase()} pour`,
  });

  revalidatePath(`/devis/${quoteId}`);
  revalidatePath("/devis");
}

export async function deleteQuote(quoteId: string) {
  await requireUser();
  await db.quote.delete({ where: { id: quoteId } });
  revalidatePath("/devis");
  redirect("/devis");
}
