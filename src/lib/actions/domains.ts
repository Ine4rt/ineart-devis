"use server";

import { revalidatePath } from "next/cache";
import { redirect } from "next/navigation";

import { requireUser } from "@/lib/auth";
import { db } from "@/lib/db";
import { bool, date, dec, logActivity, optStr, str } from "@/lib/actions/helpers";
import { addCycle } from "@/lib/renewals";

function readDomainFields(formData: FormData) {
  return {
    name: str(formData, "name").toLowerCase().replace(/^https?:\/\//, "").replace(/\/$/, ""),
    registrar: optStr(formData, "registrar"),
    purchasedAt: date(formData, "purchasedAt"),
    expiresAt: date(formData, "expiresAt"),
    renewalPrice: dec(formData, "renewalPrice"),
    autoRenew: bool(formData, "autoRenew"),
    dnsProvider: optStr(formData, "dnsProvider"),
    nameservers: optStr(formData, "nameservers"),
    sslProvider: optStr(formData, "sslProvider"),
    sslExpiresAt: date(formData, "sslExpiresAt"),
    sslAutoRenew: bool(formData, "sslAutoRenew"),
    notes: optStr(formData, "notes"),
    clientId: optStr(formData, "clientId"),
    projectId: optStr(formData, "projectId"),
  };
}

export async function createDomain(formData: FormData) {
  await requireUser();
  const fields = readDomainFields(formData);
  if (!fields.name) throw new Error("Le nom de domaine est obligatoire.");

  const domain = await db.domain.create({ data: fields });

  await logActivity({
    entityType: "domain",
    entityId: domain.id,
    clientId: domain.clientId,
    action: "created",
    summary: `Domaine ${domain.name} enregistré pour`,
  });

  revalidatePath("/domaines");
  redirect(`/domaines/${domain.id}`);
}

export async function updateDomain(domainId: string, formData: FormData) {
  await requireUser();
  const domain = await db.domain.update({
    where: { id: domainId },
    data: readDomainFields(formData),
  });

  revalidatePath(`/domaines/${domainId}`);
  revalidatePath("/domaines");
  redirect(`/domaines/${domain.id}`);
}

export async function deleteDomain(domainId: string) {
  await requireUser();
  await db.domain.delete({ where: { id: domainId } });
  revalidatePath("/domaines");
  redirect("/domaines");
}

/**
 * Marque un domaine comme renouvelé : décale l'expiration d'un an et journalise.
 * Évite la manipulation de date à la main, source d'erreurs répétitives.
 */
export async function renewDomain(domainId: string) {
  await requireUser();
  const domain = await db.domain.findUnique({ where: { id: domainId } });
  if (!domain) return;

  const base = domain.expiresAt && domain.expiresAt > new Date() ? domain.expiresAt : new Date();

  await db.domain.update({
    where: { id: domainId },
    data: { expiresAt: addCycle(base, "YEARLY") },
  });

  await logActivity({
    entityType: "domain",
    entityId: domainId,
    clientId: domain.clientId,
    action: "renewed",
    summary: `Domaine ${domain.name} renouvelé pour un an —`,
  });

  revalidatePath(`/domaines/${domainId}`);
  revalidatePath("/echeancier");
}

export async function createSubdomain(domainId: string, formData: FormData) {
  await requireUser();
  const name = str(formData, "name");
  if (!name) return;

  await db.subdomain.create({
    data: {
      domainId,
      name,
      target: optStr(formData, "target"),
      notes: optStr(formData, "notes"),
    },
  });

  revalidatePath(`/domaines/${domainId}`);
}

export async function deleteSubdomain(subdomainId: string, domainId: string) {
  await requireUser();
  await db.subdomain.delete({ where: { id: subdomainId } });
  revalidatePath(`/domaines/${domainId}`);
}
