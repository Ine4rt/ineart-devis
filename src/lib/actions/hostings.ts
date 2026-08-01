"use server";

import type { BillingCycle } from "@prisma/client";
import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";

import { requireUser } from "@/lib/auth";
import { BILLING_CYCLE } from "@/lib/constants";
import { db } from "@/lib/db";
import { bool, date, dec, enumOf, int, logActivity, optStr, str } from "@/lib/actions/helpers";
import { addCycle } from "@/lib/renewals";

const CYCLES = Object.keys(BILLING_CYCLE) as BillingCycle[];

function readHostingFields(formData: FormData) {
  return {
    provider: str(formData, "provider"),
    plan: optStr(formData, "plan"),
    label: optStr(formData, "label"),
    price: dec(formData, "price"),
    billingCycle: enumOf(formData, "billingCycle", CYCLES, "YEARLY"),
    renewsAt: date(formData, "renewsAt"),
    autoRenew: bool(formData, "autoRenew"),
    diskSpaceGb: int(formData, "diskSpaceGb") || null,
    bandwidth: optStr(formData, "bandwidth"),
    serverIp: optStr(formData, "serverIp"),
    region: optStr(formData, "region"),
    phpVersion: optStr(formData, "phpVersion"),
    controlPanelUrl: optStr(formData, "controlPanelUrl"),
    technicalNotes: optStr(formData, "technicalNotes"),
    notes: optStr(formData, "notes"),
    clientId: optStr(formData, "clientId"),
    projectId: optStr(formData, "projectId"),
  };
}

export async function createHosting(formData: FormData) {
  await requireUser();
  const fields = readHostingFields(formData);
  if (!fields.provider) throw new Error("Le fournisseur est obligatoire.");

  const hosting = await db.hosting.create({ data: fields });

  await logActivity({
    entityType: "hosting",
    entityId: hosting.id,
    clientId: hosting.clientId,
    action: "created",
    summary: `Hébergement ${hosting.provider} enregistré pour`,
  });

  revalidatePath("/hebergements");
  redirect(`/hebergements/${hosting.id}`);
}

export async function updateHosting(hostingId: string, formData: FormData) {
  await requireUser();
  await db.hosting.update({ where: { id: hostingId }, data: readHostingFields(formData) });

  revalidatePath(`/hebergements/${hostingId}`);
  revalidatePath("/hebergements");
  redirect(`/hebergements/${hostingId}`);
}

export async function deleteHosting(hostingId: string) {
  await requireUser();
  await db.hosting.delete({ where: { id: hostingId } });
  revalidatePath("/hebergements");
  redirect("/hebergements");
}

/** Décale l'échéance d'un cycle complet après paiement au fournisseur. */
export async function renewHosting(hostingId: string) {
  await requireUser();
  const hosting = await db.hosting.findUnique({ where: { id: hostingId } });
  if (!hosting) return;

  const base = hosting.renewsAt && hosting.renewsAt > new Date() ? hosting.renewsAt : new Date();

  await db.hosting.update({
    where: { id: hostingId },
    data: { renewsAt: addCycle(base, hosting.billingCycle) },
  });

  await logActivity({
    entityType: "hosting",
    entityId: hostingId,
    clientId: hosting.clientId,
    action: "renewed",
    summary: `Hébergement ${hosting.provider} renouvelé —`,
  });

  revalidatePath(`/hebergements/${hostingId}`);
  revalidatePath("/echeancier");
}
