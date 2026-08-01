import { PrismaClient } from "@prisma/client";
import { createCipheriv, createHash, randomBytes } from "node:crypto";

/**
 * Jeu de données de démonstration.
 *
 * Objectif : que la console soit immédiatement lisible et évaluable — un
 * tableau de bord vide ne permet pas de juger un outil de pilotage. Les
 * données couvrent volontairement tous les cas de figure : un impayé, une
 * échéance imminente, un domaine sans reconduction automatique, un site en
 * ligne non couvert par un abonnement, un prospect à chaque étape du pipeline.
 *
 * Le script est idempotent : il ne fait rien si la base contient déjà des
 * clients.  Pour repartir de zéro :  rm prisma/dev.db && npm run setup
 */

const db = new PrismaClient();

/** Reproduit `encryptSecret` sans importer le module serveur (contexte Node pur). */
function encryptSecret(plain: string): string {
  const secret = process.env.APP_ENCRYPTION_KEY;
  if (!secret) return "";
  const key = createHash("sha256").update(secret).digest();
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(plain, "utf8"), cipher.final()]);
  return [
    iv.toString("base64url"),
    cipher.getAuthTag().toString("base64url"),
    ciphertext.toString("base64url"),
  ].join(".");
}

/** Date relative à aujourd'hui, pour que les échéances restent pertinentes. */
function days(offset: number): Date {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  date.setDate(date.getDate() + offset);
  return date;
}

function months(offset: number): Date {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  date.setMonth(date.getMonth() + offset);
  return date;
}

async function main() {
  if ((await db.client.count()) > 0) {
    console.log("↷ La base contient déjà des données — seed ignoré.");
    return;
  }

  console.log("→ Création du jeu de démonstration…");

  // ---------------------------------------------------------------- Clients

  const boulangerie = await db.client.create({
    data: {
      reference: "CLI-0001",
      company: "Boulangerie Martin",
      firstName: "Sophie",
      lastName: "Martin",
      email: "contact@boulangerie-martin.be",
      phone: "+32 4 78 12 34 56",
      addressLine1: "Rue de la Station 42",
      postalCode: "4000",
      city: "Liège",
      country: "Belgique",
      vatNumber: "BE0712.345.678",
      industry: "Artisanat alimentaire",
      status: "ACTIVE",
      pipelineStage: null,
      source: "Prospection à froid",
      wonAt: months(-14),
      notes:
        "Très réactive par téléphone, moins par e-mail.\nSouhaite mettre en avant les produits du jour.\nPaiement systématiquement à réception de facture.",
    },
  });

  const garage = await db.client.create({
    data: {
      reference: "CLI-0002",
      company: "Garage Dubois & Fils",
      firstName: "Marc",
      lastName: "Dubois",
      email: "info@garagedubois.be",
      phone: "+32 4 87 65 43 21",
      addressLine1: "Chaussée de Tongres 118",
      postalCode: "4600",
      city: "Visé",
      country: "Belgique",
      vatNumber: "BE0698.111.222",
      industry: "Automobile",
      status: "ACTIVE",
      source: "Recommandation",
      wonAt: months(-8),
      notes: "Recommandé par la Boulangerie Martin. Veut ajouter la prise de rendez-vous en ligne.",
    },
  });

  const cabinet = await db.client.create({
    data: {
      reference: "CLI-0003",
      company: "Cabinet Lambert",
      firstName: "Claire",
      lastName: "Lambert",
      email: "c.lambert@cabinet-lambert.be",
      phone: "+32 2 511 22 33",
      addressLine1: "Avenue Louise 210",
      postalCode: "1050",
      city: "Bruxelles",
      country: "Belgique",
      vatNumber: "BE0644.987.654",
      industry: "Services juridiques",
      status: "ACTIVE",
      source: "Bouche-à-oreille",
      wonAt: months(-3),
      notes: "Exigeante sur la confidentialité. Aucun outil de suivi tiers sur le site.",
    },
  });

  const fleuriste = await db.client.create({
    data: {
      reference: "CLI-0004",
      company: "Fleurs & Sens",
      firstName: "Nadia",
      lastName: "Bensalem",
      email: "bonjour@fleursetsens.be",
      phone: "+32 4 96 22 11 00",
      postalCode: "4020",
      city: "Liège",
      country: "Belgique",
      industry: "Commerce de détail",
      status: "PROSPECT",
      pipelineStage: "QUOTE_SENT",
      potentialValue: 540,
      source: "Prospection à froid",
      notes: "Aucun site actuellement, uniquement une page Facebook. Devis envoyé, relance prévue.",
    },
  });

  const menuiserie = await db.client.create({
    data: {
      reference: "CLI-0005",
      company: "Menuiserie Renard",
      firstName: "Thomas",
      lastName: "Renard",
      email: "t.renard@outlook.be",
      phone: "+32 4 71 55 66 77",
      city: "Herstal",
      country: "Belgique",
      industry: "Construction",
      status: "PROSPECT",
      pipelineStage: "MEETING",
      potentialValue: 720,
      source: "Prospection à froid",
      notes: "Rendez-vous fixé. Cherche surtout à être trouvé sur Google localement.",
    },
  });

  const coiffure = await db.client.create({
    data: {
      reference: "CLI-0006",
      company: "Salon Éclat",
      firstName: "Julie",
      lastName: "Peeters",
      email: "salon.eclat@gmail.com",
      city: "Namur",
      country: "Belgique",
      industry: "Beauté & bien-être",
      status: "PROSPECT",
      pipelineStage: "CONTACTED",
      potentialValue: 420,
      source: "Réseaux sociaux",
      notes: "Premier e-mail envoyé, pas encore de réponse.",
    },
  });

  const traiteur = await db.client.create({
    data: {
      reference: "CLI-0007",
      company: "Traiteur Vanderlinden",
      city: "Huy",
      country: "Belgique",
      industry: "Restauration",
      status: "PROSPECT",
      pipelineStage: "IDENTIFIED",
      potentialValue: 480,
      source: "Prospection à froid",
      notes: "Repéré via l'annuaire local. Aucun site web. À contacter.",
    },
  });

  // --------------------------------------------------------------- Projets

  const siteBoulangerie = await db.project.create({
    data: {
      clientId: boulangerie.id,
      name: "Site vitrine — Boulangerie Martin",
      url: "https://boulangerie-martin.be",
      description:
        "Site vitrine sur mesure : présentation, produits du jour, horaires et plan d'accès.",
      type: "VITRINE",
      status: "MAINTENANCE",
      progress: 100,
      isCustom: true,
      techStack: JSON.stringify(["Next.js", "Tailwind CSS", "TypeScript"]),
      startedAt: months(-15),
      launchedAt: months(-14),
      deliveredAt: months(-14),
      notes: "Photos fournies par la cliente. Mise à jour saisonnière des visuels deux fois par an.",
    },
  });

  const siteGarage = await db.project.create({
    data: {
      clientId: garage.id,
      name: "Site vitrine — Garage Dubois",
      url: "https://garagedubois.be",
      description: "Présentation des services, tarifs indicatifs et formulaire de contact.",
      type: "VITRINE",
      status: "LIVE",
      progress: 100,
      isCustom: true,
      techStack: JSON.stringify(["Astro", "Tailwind CSS"]),
      startedAt: months(-9),
      launchedAt: months(-8),
      deliveredAt: months(-8),
    },
  });

  const siteCabinet = await db.project.create({
    data: {
      clientId: cabinet.id,
      name: "Site institutionnel — Cabinet Lambert",
      url: "https://cabinet-lambert.be",
      description: "Site institutionnel sobre, orienté prise de contact qualifiée.",
      type: "VITRINE",
      status: "DEVELOPMENT",
      progress: 65,
      isCustom: true,
      techStack: JSON.stringify(["Next.js", "Tailwind CSS", "Prisma"]),
      startedAt: months(-2),
      notes: "En attente des textes définitifs pour la page « Domaines d'intervention ».",
    },
  });

  // Site en ligne volontairement sans abonnement : déclenche l'alerte
  // « site non couvert » du tableau de bord.
  const siteRdv = await db.project.create({
    data: {
      clientId: garage.id,
      name: "Module de prise de rendez-vous",
      url: "https://rdv.garagedubois.be",
      description: "Application de réservation de créneaux d'atelier.",
      type: "WEBAPP",
      status: "LIVE",
      progress: 100,
      isCustom: true,
      techStack: JSON.stringify(["Next.js", "Prisma", "PostgreSQL"]),
      startedAt: months(-4),
      launchedAt: months(-2),
    },
  });

  await db.milestone.createMany({
    data: [
      { projectId: siteBoulangerie.id, title: "Mise en ligne initiale", happenedAt: months(-14) },
      {
        projectId: siteBoulangerie.id,
        title: "Ajout de la section « Produits du jour »",
        detail: "Demande formulée lors de l'appel du 12 mars.",
        happenedAt: months(-6),
      },
      { projectId: siteBoulangerie.id, title: "Mise à jour des visuels de saison", happenedAt: months(-1) },
      { projectId: siteGarage.id, title: "Mise en ligne initiale", happenedAt: months(-8) },
      { projectId: siteCabinet.id, title: "Validation de la maquette", happenedAt: months(-1) },
    ],
  });

  // -------------------------------------------------------------- Domaines

  const domaineBoulangerie = await db.domain.create({
    data: {
      clientId: boulangerie.id,
      projectId: siteBoulangerie.id,
      name: "boulangerie-martin.be",
      registrar: "OVHcloud",
      purchasedAt: months(-15),
      expiresAt: days(38),
      renewalPrice: 12,
      autoRenew: true,
      dnsProvider: "Cloudflare",
      nameservers: "ns1.cloudflare.com\nns2.cloudflare.com",
      sslProvider: "Let's Encrypt",
      sslExpiresAt: days(24),
      sslAutoRenew: true,
    },
  });

  await db.subdomain.createMany({
    data: [
      { domainId: domaineBoulangerie.id, name: "www", target: "boulangerie-martin.be" },
      { domainId: domaineBoulangerie.id, name: "mail", target: "mx.infomaniak.com" },
    ],
  });

  // Échéance proche + reconduction manuelle : le cas qui doit sauter aux yeux.
  await db.domain.create({
    data: {
      clientId: garage.id,
      projectId: siteGarage.id,
      name: "garagedubois.be",
      registrar: "Gandi",
      purchasedAt: months(-9),
      expiresAt: days(11),
      renewalPrice: 15,
      autoRenew: false,
      dnsProvider: "Gandi",
      sslProvider: "Let's Encrypt",
      sslExpiresAt: days(45),
      notes: "Reconduction automatique désactivée à la demande du client (carte expirée).",
    },
  });

  await db.domain.create({
    data: {
      clientId: cabinet.id,
      projectId: siteCabinet.id,
      name: "cabinet-lambert.be",
      registrar: "OVHcloud",
      purchasedAt: months(-3),
      expiresAt: months(9),
      renewalPrice: 12,
      autoRenew: true,
      dnsProvider: "OVHcloud",
      sslProvider: "Let's Encrypt",
      sslExpiresAt: days(72),
    },
  });

  // ---------------------------------------------------------- Hébergements

  const hebergementMutualise = await db.hosting.create({
    data: {
      clientId: boulangerie.id,
      projectId: siteBoulangerie.id,
      provider: "Infomaniak",
      plan: "Hébergement Web Starter",
      price: 71.4,
      billingCycle: "YEARLY",
      renewsAt: days(52),
      autoRenew: true,
      diskSpaceGb: 10,
      region: "Genève",
      phpVersion: "8.3",
      controlPanelUrl: "https://manager.infomaniak.com",
      technicalNotes: "Sauvegarde quotidienne incluse. Cron de purge du cache à 4 h.",
    },
  });

  await db.hosting.create({
    data: {
      clientId: garage.id,
      projectId: siteGarage.id,
      provider: "Vercel",
      plan: "Pro",
      label: "Déploiement statique",
      price: 20,
      billingCycle: "MONTHLY",
      renewsAt: days(9),
      autoRenew: true,
      region: "Francfort",
      technicalNotes: "Déploiement automatique depuis la branche main.",
    },
  });

  await db.hosting.create({
    data: {
      clientId: cabinet.id,
      projectId: siteCabinet.id,
      provider: "Hetzner",
      plan: "CX22",
      price: 4.51,
      billingCycle: "MONTHLY",
      renewsAt: days(21),
      autoRenew: true,
      diskSpaceGb: 40,
      serverIp: "203.0.113.24",
      region: "Nuremberg",
      technicalNotes: "Docker + Caddy. Sauvegarde nocturne vers un espace de stockage distant.",
    },
  });

  // ----------------------------------------------------------- Identifiants

  if (process.env.APP_ENCRYPTION_KEY) {
    await db.credential.createMany({
      data: [
        {
          label: "FTP production",
          kind: "FTP",
          username: "martin_prod",
          secretEnc: encryptSecret("EXEMPLE-mot-de-passe-ftp"),
          url: "ftp://boulangerie-martin.be",
          clientId: boulangerie.id,
          hostingId: hebergementMutualise.id,
          notes: "Accès en écriture limité au dossier /web.",
        },
        {
          label: "Compte registrar",
          kind: "REGISTRAR",
          username: "sophie.martin@…",
          secretEnc: encryptSecret("EXEMPLE-mot-de-passe-ovh"),
          url: "https://www.ovh.com/manager",
          clientId: boulangerie.id,
          domainId: domaineBoulangerie.id,
        },
      ],
    });
  }

  // ----------------------------------------------------------- Abonnements

  const abonnementBoulangerie = await db.subscription.create({
    data: {
      clientId: boulangerie.id,
      projectId: siteBoulangerie.id,
      name: "Formule complète — site, domaine et maintenance",
      type: "FULL_CARE",
      amount: 480,
      currency: "EUR",
      billingCycle: "YEARLY",
      status: "ACTIVE",
      autoRenew: true,
      startedAt: months(-14),
      currentPeriodStart: months(-2),
      currentPeriodEnd: months(10),
      nextRenewalAt: months(10),
      paymentMethod: "Virement",
    },
  });

  const abonnementGarage = await db.subscription.create({
    data: {
      clientId: garage.id,
      projectId: siteGarage.id,
      name: "Hébergement + maintenance annuelle",
      type: "FULL_CARE",
      amount: 540,
      currency: "EUR",
      billingCycle: "YEARLY",
      status: "ACTIVE",
      autoRenew: true,
      startedAt: months(-8),
      currentPeriodStart: months(-8),
      currentPeriodEnd: days(24),
      nextRenewalAt: days(24),
      paymentMethod: "Virement",
    },
  });

  const abonnementCabinet = await db.subscription.create({
    data: {
      clientId: cabinet.id,
      projectId: siteCabinet.id,
      name: "Maintenance et hébergement",
      type: "FULL_CARE",
      amount: 65,
      currency: "EUR",
      billingCycle: "MONTHLY",
      status: "ACTIVE",
      autoRenew: true,
      startedAt: months(-3),
      currentPeriodStart: days(-12),
      currentPeriodEnd: days(18),
      nextRenewalAt: days(18),
      paymentMethod: "Domiciliation",
    },
  });

  // ------------------------------------------------------------- Échéances

  await db.subscriptionPeriod.createMany({
    data: [
      // Historique réglé.
      {
        subscriptionId: abonnementBoulangerie.id,
        periodStart: months(-14),
        periodEnd: months(-2),
        dueAt: months(-14),
        amount: 480,
        status: "PAID",
        paidAt: months(-14),
        paymentMethod: "Virement",
        reference: "VIR-2024-018",
      },
      {
        subscriptionId: abonnementBoulangerie.id,
        periodStart: months(-2),
        periodEnd: months(10),
        dueAt: months(-2),
        amount: 480,
        status: "PAID",
        paidAt: months(-2),
        paymentMethod: "Virement",
        reference: "VIR-2025-004",
      },
      // Impayé en retard : alimente l'alerte du tableau de bord.
      {
        subscriptionId: abonnementGarage.id,
        periodStart: months(-8),
        periodEnd: days(24),
        dueAt: days(-21),
        amount: 540,
        status: "OVERDUE",
        notes: "Relancé une première fois par e-mail.",
      },
      {
        subscriptionId: abonnementCabinet.id,
        periodStart: months(-3),
        periodEnd: months(-2),
        dueAt: months(-3),
        amount: 65,
        status: "PAID",
        paidAt: months(-3),
        paymentMethod: "Domiciliation",
      },
      {
        subscriptionId: abonnementCabinet.id,
        periodStart: months(-2),
        periodEnd: months(-1),
        dueAt: months(-2),
        amount: 65,
        status: "PAID",
        paidAt: months(-2),
        paymentMethod: "Domiciliation",
      },
      {
        subscriptionId: abonnementCabinet.id,
        periodStart: days(-12),
        periodEnd: days(18),
        dueAt: days(6),
        amount: 65,
        status: "DUE",
      },
    ],
  });

  // ------------------------------------------------------------------ Devis

  const devisFleuriste = await db.quote.create({
    data: {
      number: `DEV-${new Date().getFullYear()}-001`,
      clientId: fleuriste.id,
      title: "Création d'un site vitrine + formule annuelle",
      status: "SENT",
      issuedAt: days(-9),
      validUntil: days(21),
      sentAt: days(-9),
      taxRate: 21,
      discountPct: 0,
      intro:
        "Suite à notre échange, voici ma proposition pour la création de votre site vitrine et son suivi annuel.",
      terms:
        "Acompte de 30 % à la commande, solde à la mise en ligne.\nL'abonnement annuel démarre à la mise en ligne et couvre l'hébergement, le nom de domaine et la maintenance.",
      items: {
        create: [
          {
            label: "Conception et développement du site vitrine",
            description: "Cinq pages sur mesure, responsive, optimisé pour le référencement local.",
            quantity: 1,
            unitPrice: 1450,
            position: 0,
          },
          {
            label: "Séance photo des compositions",
            description: "Demi-journée sur place, retouches incluses.",
            quantity: 1,
            unitPrice: 280,
            position: 1,
          },
          {
            label: "Formule annuelle — domaine, hébergement et maintenance",
            description: "Première année, facturée à la mise en ligne.",
            quantity: 1,
            unitPrice: 540,
            position: 2,
          },
        ],
      },
    },
  });

  await db.quote.create({
    data: {
      number: `DEV-${new Date().getFullYear()}-002`,
      clientId: menuiserie.id,
      title: "Site vitrine — Menuiserie Renard",
      status: "DRAFT",
      issuedAt: days(-2),
      validUntil: days(28),
      taxRate: 21,
      items: {
        create: [
          {
            label: "Site vitrine sur mesure",
            description: "Quatre pages, galerie de réalisations.",
            quantity: 1,
            unitPrice: 1250,
            position: 0,
          },
          {
            label: "Formule annuelle",
            quantity: 1,
            unitPrice: 480,
            position: 1,
          },
        ],
      },
    },
  });

  // ------------------------------------------------------------- Tâches

  await db.task.createMany({
    data: [
      {
        title: "Relancer le Garage Dubois pour la facture impayée",
        description: "Échéance dépassée de trois semaines. Proposer un paiement en deux fois.",
        status: "TODO",
        priority: "URGENT",
        kind: "FOLLOW_UP",
        dueAt: days(-3),
        clientId: garage.id,
      },
      {
        title: "Renouveler garagedubois.be manuellement",
        description: "La reconduction automatique est désactivée chez Gandi.",
        status: "TODO",
        priority: "HIGH",
        kind: "REMINDER",
        dueAt: days(5),
        clientId: garage.id,
      },
      {
        title: "Relancer Fleurs & Sens sur le devis",
        status: "TODO",
        priority: "HIGH",
        kind: "CALL",
        dueAt: days(2),
        clientId: fleuriste.id,
      },
      {
        title: "Récupérer les textes du Cabinet Lambert",
        description: "Page « Domaines d'intervention » toujours en attente.",
        status: "IN_PROGRESS",
        priority: "MEDIUM",
        kind: "EMAIL",
        dueAt: days(7),
        clientId: cabinet.id,
        projectId: siteCabinet.id,
      },
      {
        title: "Proposer un contrat de maintenance pour le module de rendez-vous",
        description: "Le site est en ligne mais n'est couvert par aucun abonnement.",
        status: "TODO",
        priority: "MEDIUM",
        kind: "TASK",
        dueAt: days(14),
        clientId: garage.id,
        projectId: siteRdv.id,
      },
      {
        title: "Rendez-vous Menuiserie Renard",
        status: "TODO",
        priority: "HIGH",
        kind: "MEETING",
        dueAt: days(4),
        clientId: menuiserie.id,
      },
      {
        title: "Mettre à jour les visuels de saison — Boulangerie Martin",
        status: "DONE",
        priority: "LOW",
        kind: "TASK",
        dueAt: months(-1),
        completedAt: months(-1),
        clientId: boulangerie.id,
        projectId: siteBoulangerie.id,
      },
    ],
  });

  // -------------------------------------------------------------- Notes

  await db.note.createMany({
    data: [
      {
        clientId: boulangerie.id,
        title: "Préférences de contact",
        body: "Toujours appeler entre 14 h et 16 h : la boulangerie est fermée l'après-midi et Sophie est disponible.",
        pinned: true,
      },
      {
        clientId: garage.id,
        title: "Situation de paiement",
        body: "Carte bancaire expirée chez Gandi, d'où la reconduction manuelle.\nMarc a promis de régulariser à la fin du mois.",
        pinned: true,
      },
      {
        clientId: cabinet.id,
        body: "Aucun outil de mesure d'audience tiers autorisé sur le site : contrainte de confidentialité du cabinet.",
        pinned: false,
      },
      {
        clientId: fleuriste.id,
        body: "Budget annoncé autour de 1 500 €. Sensible à l'argument du référencement local.",
        pinned: false,
      },
    ],
  });

  // ------------------------------------------------------------ Activité

  await db.activity.createMany({
    data: [
      {
        entityType: "quote",
        entityId: devisFleuriste.id,
        clientId: fleuriste.id,
        action: "sent",
        summary: "Devis envoyé à",
        createdAt: days(-9),
      },
      {
        entityType: "subscription",
        entityId: abonnementCabinet.id,
        clientId: cabinet.id,
        action: "paid",
        summary: "Paiement mensuel encaissé pour",
        createdAt: days(-12),
      },
      {
        entityType: "project",
        entityId: siteRdv.id,
        clientId: garage.id,
        action: "launched",
        summary: "Module de prise de rendez-vous mis en ligne pour",
        createdAt: months(-2),
      },
      {
        entityType: "client",
        entityId: cabinet.id,
        clientId: cabinet.id,
        action: "created",
        summary: "Nouveau client créé :",
        createdAt: months(-3),
      },
    ],
  });

  console.log(`✓ Jeu de démonstration créé.
  7 clients (3 actifs, 4 prospects) · 4 projets · 3 domaines · 3 hébergements
  3 abonnements · 6 échéances · 2 devis · 7 tâches · 4 notes

  Ouvrez http://localhost:3000 — la première visite crée votre compte.`);
}

main()
  .catch((error) => {
    console.error("✗ Échec du seed :", error);
    process.exit(1);
  })
  .finally(() => db.$disconnect());
