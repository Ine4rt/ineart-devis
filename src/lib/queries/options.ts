import "server-only";

import { db } from "@/lib/db";

/**
 * Listes déroulantes partagées par tous les formulaires.
 *
 * Centralisées pour garantir le même tri et les mêmes exclusions partout : un
 * client archivé ne doit pas apparaître dans un nouveau contrat, mais doit
 * rester sélectionnable si une fiche existante le référence déjà.
 */

export async function getClientOptions(includeId?: string | null) {
  return db.client.findMany({
    where: {
      OR: [
        { status: { in: ["ACTIVE", "PROSPECT", "PAUSED"] } },
        ...(includeId ? [{ id: includeId }] : []),
      ],
    },
    select: { id: true, company: true },
    orderBy: { company: "asc" },
  });
}

export async function getProjectOptions(clientId?: string | null) {
  return db.project.findMany({
    where: clientId ? { clientId } : undefined,
    select: { id: true, name: true, clientId: true },
    orderBy: { name: "asc" },
  });
}
