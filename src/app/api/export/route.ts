import { NextResponse } from "next/server";

import { getSessionUser } from "@/lib/auth";
import { db } from "@/lib/db";

/**
 * Export complet des données au format JSON.
 *
 * Filet de sécurité et garantie de réversibilité : vos données ne sont pas
 * captives de cet outil. Les secrets du coffre sont exportés **chiffrés** —
 * l'export seul ne permet pas de les lire sans APP_ENCRYPTION_KEY.
 */
export async function GET() {
  const user = await getSessionUser();
  if (!user) return new NextResponse("Non autorisé", { status: 401 });

  const [
    clients,
    contacts,
    projects,
    milestones,
    domains,
    subdomains,
    hostings,
    subscriptions,
    periods,
    quotes,
    quoteItems,
    tasks,
    notes,
    documents,
    credentials,
    activities,
  ] = await Promise.all([
    db.client.findMany(),
    db.contact.findMany(),
    db.project.findMany(),
    db.milestone.findMany(),
    db.domain.findMany(),
    db.subdomain.findMany(),
    db.hosting.findMany(),
    db.subscription.findMany(),
    db.subscriptionPeriod.findMany(),
    db.quote.findMany(),
    db.quoteItem.findMany(),
    db.task.findMany(),
    db.note.findMany(),
    db.document.findMany(),
    db.credential.findMany(),
    db.activity.findMany(),
  ]);

  const payload = {
    exportedAt: new Date().toISOString(),
    version: 1,
    note: "Les champs `secretEnc` restent chiffrés (AES-256-GCM). Conservez APP_ENCRYPTION_KEY.",
    data: {
      clients,
      contacts,
      projects,
      milestones,
      domains,
      subdomains,
      hostings,
      subscriptions,
      periods,
      quotes,
      quoteItems,
      tasks,
      notes,
      documents,
      credentials,
      activities,
    },
  };

  const stamp = new Date().toISOString().slice(0, 10);

  return new NextResponse(JSON.stringify(payload, null, 2), {
    headers: {
      "Content-Type": "application/json",
      "Content-Disposition": `attachment; filename="ineart-console-${stamp}.json"`,
      "Cache-Control": "no-store",
    },
  });
}
