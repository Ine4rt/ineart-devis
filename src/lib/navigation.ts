import {
  Banknote,
  Building2,
  CalendarClock,
  FileText,
  FolderKanban,
  Globe,
  LayoutDashboard,
  ListChecks,
  Paperclip,
  Receipt,
  Server,
  Settings,
  Target,
  Waypoints,
} from "lucide-react";
import type { LucideIcon } from "lucide-react";

/**
 * Arborescence de navigation.
 *
 * Regroupée par intention plutôt que par table de la base : on cherche
 * « combien on me doit » (Revenus) ou « qu'est-ce qui expire » (Pilotage),
 * pas « la table subscription_period ». Cinq groupes courts valent mieux
 * qu'une liste de quinze entrées.
 */

export interface NavLink {
  href: string;
  label: string;
  icon: LucideIcon;
  description: string;
  /** Correspondance exacte : réservé au tableau de bord (« / »). */
  exact?: boolean;
}

export interface NavGroup {
  label: string;
  links: NavLink[];
}

export const NAVIGATION: NavGroup[] = [
  {
    label: "Pilotage",
    links: [
      {
        href: "/",
        label: "Tableau de bord",
        icon: LayoutDashboard,
        description: "Vue d'ensemble de l'activité",
        exact: true,
      },
      {
        href: "/echeancier",
        label: "Échéancier",
        icon: CalendarClock,
        description: "Tout ce qui expire, dans l'ordre",
      },
    ],
  },
  {
    label: "Clients",
    links: [
      {
        href: "/clients",
        label: "Clients",
        icon: Building2,
        description: "Fiches complètes et historique",
      },
      {
        href: "/prospection",
        label: "Prospection",
        icon: Target,
        description: "Pipeline des entreprises à convertir",
      },
      {
        href: "/devis",
        label: "Devis",
        icon: FileText,
        description: "Propositions commerciales",
      },
    ],
  },
  {
    label: "Production",
    links: [
      {
        href: "/projets",
        label: "Projets",
        icon: FolderKanban,
        description: "Sites en cours et en ligne",
      },
      {
        href: "/domaines",
        label: "Domaines",
        icon: Globe,
        description: "Noms de domaine, DNS et SSL",
      },
      {
        href: "/hebergements",
        label: "Hébergements",
        icon: Server,
        description: "Serveurs, offres et échéances",
      },
    ],
  },
  {
    label: "Revenus",
    links: [
      {
        href: "/abonnements",
        label: "Abonnements",
        icon: Waypoints,
        description: "Contrats récurrents et renouvellements",
      },
      {
        href: "/paiements",
        label: "Paiements",
        icon: Banknote,
        description: "Qui a payé, qui doit payer",
      },
      {
        href: "/revenus",
        label: "Revenus",
        icon: Receipt,
        description: "MRR, ARR et rentabilité par client",
      },
    ],
  },
  {
    label: "Organisation",
    links: [
      {
        href: "/taches",
        label: "Tâches",
        icon: ListChecks,
        description: "Rappels, relances et échéances",
      },
      {
        href: "/documents",
        label: "Documents",
        icon: Paperclip,
        description: "Contrats et fichiers clients",
      },
    ],
  },
];

export const SETTINGS_LINK: NavLink = {
  href: "/reglages",
  label: "Réglages",
  icon: Settings,
  description: "Compte, apparence et données",
};

/** Toutes les entrées à plat — utilisé par la palette de commandes. */
export const ALL_NAV_LINKS: NavLink[] = [
  ...NAVIGATION.flatMap((group) => group.links),
  SETTINGS_LINK,
];

/** Détermine si un lien doit être marqué actif pour un chemin donné. */
export function isNavLinkActive(link: NavLink, pathname: string): boolean {
  if (link.exact) return pathname === link.href;
  return pathname === link.href || pathname.startsWith(`${link.href}/`);
}
