import type {
  BillingCycle,
  ClientStatus,
  CredentialKind,
  DocumentCategory,
  PeriodStatus,
  PipelineStage,
  ProjectStatus,
  ProjectType,
  QuoteStatus,
  SubscriptionStatus,
  SubscriptionType,
  TaskKind,
  TaskPriority,
  TaskStatus,
} from "@prisma/client";

/**
 * Vocabulaire de l'application.
 *
 * Chaque enum Prisma est décrit une seule fois ici : libellé français + tonalité
 * visuelle. Les composants (Badge, Select, filtres, graphiques) consomment ces
 * tables — ajouter une valeur à un enum se répercute partout sans chasse aux
 * chaînes en dur.
 */

export type Tone =
  | "neutral"
  | "accent"
  | "success"
  | "info"
  | "warning"
  | "caution"
  | "critical"
  | "danger"
  | "muted";

export interface EnumMeta {
  label: string;
  tone: Tone;
  description?: string;
}

function entries<T extends string>(map: Record<T, EnumMeta>) {
  return Object.entries(map) as [T, EnumMeta][];
}

// --------------------------------------------------------------------------

export const CLIENT_STATUS: Record<ClientStatus, EnumMeta> = {
  PROSPECT: { label: "Prospect", tone: "info", description: "Piste en cours de conversion" },
  ACTIVE: { label: "Actif", tone: "success", description: "Client sous contrat" },
  PAUSED: { label: "En pause", tone: "warning", description: "Relation suspendue" },
  ARCHIVED: { label: "Archivé", tone: "muted", description: "Conservé pour l'historique" },
  LOST: { label: "Perdu", tone: "danger", description: "Opportunité non concrétisée" },
};
export const CLIENT_STATUS_LIST = entries(CLIENT_STATUS);

export const PIPELINE_STAGE: Record<PipelineStage, EnumMeta> = {
  IDENTIFIED: { label: "Identifié", tone: "muted" },
  CONTACTED: { label: "Contacté", tone: "info" },
  MEETING: { label: "Rendez-vous", tone: "accent" },
  QUOTE_SENT: { label: "Devis envoyé", tone: "caution" },
  NEGOTIATION: { label: "Négociation", tone: "warning" },
  WON: { label: "Gagné", tone: "success" },
  LOST: { label: "Perdu", tone: "danger" },
};
export const PIPELINE_STAGE_LIST = entries(PIPELINE_STAGE);

/** Colonnes affichées dans le kanban de prospection (WON/LOST sortent du flux). */
export const PIPELINE_BOARD_STAGES: PipelineStage[] = [
  "IDENTIFIED",
  "CONTACTED",
  "MEETING",
  "QUOTE_SENT",
  "NEGOTIATION",
];

export const PROJECT_STATUS: Record<ProjectStatus, EnumMeta> = {
  DISCOVERY: { label: "Cadrage", tone: "muted" },
  DESIGN: { label: "Design", tone: "info" },
  DEVELOPMENT: { label: "Développement", tone: "accent" },
  REVIEW: { label: "Relecture", tone: "caution" },
  LIVE: { label: "En ligne", tone: "success" },
  MAINTENANCE: { label: "Maintenance", tone: "success" },
  PAUSED: { label: "En pause", tone: "warning" },
  ARCHIVED: { label: "Archivé", tone: "muted" },
};
export const PROJECT_STATUS_LIST = entries(PROJECT_STATUS);

/** Statuts considérés comme « site en production ». */
export const LIVE_PROJECT_STATUSES: ProjectStatus[] = ["LIVE", "MAINTENANCE"];

export const PROJECT_TYPE: Record<ProjectType, EnumMeta> = {
  VITRINE: { label: "Site vitrine", tone: "neutral" },
  ECOMMERCE: { label: "E-commerce", tone: "neutral" },
  LANDING: { label: "Landing page", tone: "neutral" },
  WEBAPP: { label: "Application web", tone: "neutral" },
  REFONTE: { label: "Refonte", tone: "neutral" },
  AUTRE: { label: "Autre", tone: "neutral" },
};
export const PROJECT_TYPE_LIST = entries(PROJECT_TYPE);

export const BILLING_CYCLE: Record<BillingCycle, EnumMeta> = {
  MONTHLY: { label: "Mensuel", tone: "neutral" },
  QUARTERLY: { label: "Trimestriel", tone: "neutral" },
  YEARLY: { label: "Annuel", tone: "neutral" },
  BIENNIAL: { label: "Biennal", tone: "neutral" },
};
export const BILLING_CYCLE_LIST = entries(BILLING_CYCLE);

export const SUBSCRIPTION_STATUS: Record<SubscriptionStatus, EnumMeta> = {
  ACTIVE: { label: "Actif", tone: "success" },
  PENDING: { label: "En attente", tone: "info" },
  PAST_DUE: { label: "Impayé", tone: "danger" },
  PAUSED: { label: "En pause", tone: "warning" },
  CANCELLED: { label: "Résilié", tone: "muted" },
};
export const SUBSCRIPTION_STATUS_LIST = entries(SUBSCRIPTION_STATUS);

export const SUBSCRIPTION_TYPE: Record<SubscriptionType, EnumMeta> = {
  FULL_CARE: { label: "Formule complète", tone: "accent", description: "Domaine + hébergement + maintenance" },
  MAINTENANCE: { label: "Maintenance", tone: "neutral" },
  HOSTING_DOMAIN: { label: "Domaine + hébergement", tone: "neutral" },
  SEO: { label: "Référencement", tone: "neutral" },
  SUPPORT: { label: "Support", tone: "neutral" },
  CUSTOM: { label: "Sur mesure", tone: "neutral" },
};
export const SUBSCRIPTION_TYPE_LIST = entries(SUBSCRIPTION_TYPE);

export const PERIOD_STATUS: Record<PeriodStatus, EnumMeta> = {
  DUE: { label: "À payer", tone: "warning" },
  PAID: { label: "Payé", tone: "success" },
  OVERDUE: { label: "En retard", tone: "danger" },
  CANCELLED: { label: "Annulé", tone: "muted" },
  REFUNDED: { label: "Remboursé", tone: "info" },
};
export const PERIOD_STATUS_LIST = entries(PERIOD_STATUS);

export const QUOTE_STATUS: Record<QuoteStatus, EnumMeta> = {
  DRAFT: { label: "Brouillon", tone: "muted" },
  SENT: { label: "Envoyé", tone: "info" },
  ACCEPTED: { label: "Accepté", tone: "success" },
  REJECTED: { label: "Refusé", tone: "danger" },
  EXPIRED: { label: "Expiré", tone: "warning" },
};
export const QUOTE_STATUS_LIST = entries(QUOTE_STATUS);

export const TASK_STATUS: Record<TaskStatus, EnumMeta> = {
  TODO: { label: "À faire", tone: "neutral" },
  IN_PROGRESS: { label: "En cours", tone: "accent" },
  BLOCKED: { label: "Bloqué", tone: "danger" },
  DONE: { label: "Terminé", tone: "success" },
};
export const TASK_STATUS_LIST = entries(TASK_STATUS);

export const TASK_PRIORITY: Record<TaskPriority, EnumMeta> = {
  LOW: { label: "Basse", tone: "muted" },
  MEDIUM: { label: "Normale", tone: "info" },
  HIGH: { label: "Haute", tone: "warning" },
  URGENT: { label: "Urgente", tone: "danger" },
};
export const TASK_PRIORITY_LIST = entries(TASK_PRIORITY);

export const TASK_KIND: Record<TaskKind, EnumMeta> = {
  TASK: { label: "Tâche", tone: "neutral" },
  REMINDER: { label: "Rappel", tone: "neutral" },
  CALL: { label: "Appel", tone: "neutral" },
  EMAIL: { label: "E-mail", tone: "neutral" },
  MEETING: { label: "Rendez-vous", tone: "neutral" },
  FOLLOW_UP: { label: "Relance", tone: "neutral" },
};
export const TASK_KIND_LIST = entries(TASK_KIND);

export const DOCUMENT_CATEGORY: Record<DocumentCategory, EnumMeta> = {
  CONTRACT: { label: "Contrat", tone: "accent" },
  QUOTE: { label: "Devis", tone: "info" },
  INVOICE: { label: "Facture", tone: "success" },
  DESIGN: { label: "Design", tone: "neutral" },
  SCREENSHOT: { label: "Capture", tone: "neutral" },
  LEGAL: { label: "Juridique", tone: "warning" },
  TECHNICAL: { label: "Technique", tone: "neutral" },
  OTHER: { label: "Autre", tone: "muted" },
};
export const DOCUMENT_CATEGORY_LIST = entries(DOCUMENT_CATEGORY);

export const CREDENTIAL_KIND: Record<CredentialKind, EnumMeta> = {
  FTP: { label: "FTP / SFTP", tone: "neutral" },
  SSH: { label: "SSH", tone: "neutral" },
  DATABASE: { label: "Base de données", tone: "neutral" },
  CPANEL: { label: "Panneau d'administration", tone: "neutral" },
  CMS: { label: "CMS", tone: "neutral" },
  REGISTRAR: { label: "Registrar", tone: "neutral" },
  EMAIL: { label: "Messagerie", tone: "neutral" },
  ANALYTICS: { label: "Analytics", tone: "neutral" },
  OTHER: { label: "Autre", tone: "neutral" },
};
export const CREDENTIAL_KIND_LIST = entries(CREDENTIAL_KIND);

// --------------------------------------------------------------------------

/** Registrars et hébergeurs proposés en autocomplétion (datalist). */
export const COMMON_REGISTRARS = [
  "OVHcloud",
  "Gandi",
  "Namecheap",
  "Cloudflare",
  "Infomaniak",
  "Hostinger",
  "IONOS",
  "Combell",
];

export const COMMON_HOSTS = [
  "Vercel",
  "OVHcloud",
  "Infomaniak",
  "Hostinger",
  "Netlify",
  "Hetzner",
  "Scaleway",
  "o2switch",
  "Combell",
];

export const COMMON_TECHNOLOGIES = [
  "Next.js",
  "React",
  "Astro",
  "Tailwind CSS",
  "TypeScript",
  "WordPress",
  "Shopify",
  "Prisma",
  "PostgreSQL",
  "Node.js",
  "Sanity",
  "Stripe",
];
