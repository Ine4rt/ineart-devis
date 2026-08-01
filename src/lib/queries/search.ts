import "server-only";

import { db } from "@/lib/db";

/**
 * Recherche globale.
 *
 * Une seule requête par entité, exécutées en parallèle, plafonnées en nombre de
 * résultats : la palette doit répondre pendant la frappe. SQLite reste sur du
 * `contains` — sur quelques milliers de lignes c'est instantané, et cela évite
 * d'imposer une extension full-text pour un gain invisible à cette échelle.
 */

export type SearchEntity =
  | "client"
  | "projet"
  | "domaine"
  | "hebergement"
  | "abonnement"
  | "devis"
  | "tache"
  | "note"
  | "document";

export interface SearchResult {
  id: string;
  entity: SearchEntity;
  title: string;
  subtitle?: string;
  href: string;
}

const PER_ENTITY = 5;

export async function globalSearch(rawQuery: string): Promise<SearchResult[]> {
  const query = rawQuery.trim();
  if (query.length < 2) return [];

  const [clients, projects, domains, hostings, subscriptions, quotes, tasks, notes, documents] =
    await Promise.all([
      db.client.findMany({
        where: {
          OR: [
            { company: { contains: query } },
            { firstName: { contains: query } },
            { lastName: { contains: query } },
            { email: { contains: query } },
            { phone: { contains: query } },
            { vatNumber: { contains: query } },
            { reference: { contains: query } },
            { city: { contains: query } },
            { notes: { contains: query } },
          ],
        },
        select: { id: true, company: true, reference: true, email: true, city: true },
        take: PER_ENTITY,
      }),
      db.project.findMany({
        where: { OR: [{ name: { contains: query } }, { url: { contains: query } }] },
        select: { id: true, name: true, url: true, client: { select: { company: true } } },
        take: PER_ENTITY,
      }),
      db.domain.findMany({
        where: { OR: [{ name: { contains: query } }, { registrar: { contains: query } }] },
        select: { id: true, name: true, registrar: true, client: { select: { company: true } } },
        take: PER_ENTITY,
      }),
      db.hosting.findMany({
        where: {
          OR: [
            { provider: { contains: query } },
            { plan: { contains: query } },
            { label: { contains: query } },
            { serverIp: { contains: query } },
          ],
        },
        select: { id: true, provider: true, plan: true, client: { select: { company: true } } },
        take: PER_ENTITY,
      }),
      db.subscription.findMany({
        where: { OR: [{ name: { contains: query } }] },
        select: { id: true, name: true, client: { select: { company: true } } },
        take: PER_ENTITY,
      }),
      db.quote.findMany({
        where: { OR: [{ title: { contains: query } }, { number: { contains: query } }] },
        select: { id: true, title: true, number: true, client: { select: { company: true } } },
        take: PER_ENTITY,
      }),
      db.task.findMany({
        where: {
          status: { not: "DONE" },
          OR: [{ title: { contains: query } }, { description: { contains: query } }],
        },
        select: { id: true, title: true, client: { select: { company: true } } },
        take: PER_ENTITY,
      }),
      db.note.findMany({
        where: { OR: [{ title: { contains: query } }, { body: { contains: query } }] },
        select: { id: true, title: true, body: true, clientId: true },
        take: PER_ENTITY,
      }),
      db.document.findMany({
        where: { OR: [{ name: { contains: query } }, { originalName: { contains: query } }] },
        select: { id: true, name: true, clientId: true, client: { select: { company: true } } },
        take: PER_ENTITY,
      }),
    ]);

  return [
    ...clients.map((client) => ({
      id: client.id,
      entity: "client" as const,
      title: client.company,
      subtitle: [client.reference, client.city, client.email].filter(Boolean).join(" · "),
      href: `/clients/${client.id}`,
    })),
    ...projects.map((project) => ({
      id: project.id,
      entity: "projet" as const,
      title: project.name,
      subtitle: [project.client?.company, project.url].filter(Boolean).join(" · "),
      href: `/projets/${project.id}`,
    })),
    ...domains.map((domain) => ({
      id: domain.id,
      entity: "domaine" as const,
      title: domain.name,
      subtitle: [domain.client?.company, domain.registrar].filter(Boolean).join(" · "),
      href: `/domaines/${domain.id}`,
    })),
    ...hostings.map((hosting) => ({
      id: hosting.id,
      entity: "hebergement" as const,
      title: [hosting.provider, hosting.plan].filter(Boolean).join(" — "),
      subtitle: hosting.client?.company ?? undefined,
      href: `/hebergements/${hosting.id}`,
    })),
    ...subscriptions.map((subscription) => ({
      id: subscription.id,
      entity: "abonnement" as const,
      title: subscription.name,
      subtitle: subscription.client?.company ?? undefined,
      href: `/abonnements/${subscription.id}`,
    })),
    ...quotes.map((quote) => ({
      id: quote.id,
      entity: "devis" as const,
      title: `${quote.number} — ${quote.title}`,
      subtitle: quote.client?.company ?? undefined,
      href: `/devis/${quote.id}`,
    })),
    ...tasks.map((task) => ({
      id: task.id,
      entity: "tache" as const,
      title: task.title,
      subtitle: task.client?.company ?? undefined,
      href: "/taches",
    })),
    ...notes.map((note) => ({
      id: note.id,
      entity: "note" as const,
      title: note.title || note.body.slice(0, 60),
      subtitle: note.body.slice(0, 80),
      href: note.clientId ? `/clients/${note.clientId}/notes` : "/clients",
    })),
    ...documents.map((document) => ({
      id: document.id,
      entity: "document" as const,
      title: document.name,
      subtitle: document.client?.company ?? undefined,
      href: document.clientId ? `/clients/${document.clientId}/documents` : "/documents",
    })),
  ];
}
