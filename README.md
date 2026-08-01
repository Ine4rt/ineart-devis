# Ine4rt Console

Outil interne de gestion pour un développeur web indépendant : clients, sites,
domaines, hébergements, abonnements et encaissements dans une seule interface
privée.

Ce n'est pas un produit destiné aux clients finaux. Il n'y a ni inscription, ni
espace public, ni partage : un compte unique, des données locales, et une
interface pensée pour être ouverte plusieurs heures par jour.

---

## Démarrage

```bash
npm install
cp .env.example .env          # puis renseignez les deux secrets (voir ci-dessous)
npm run setup                 # crée la base SQLite + un jeu de démonstration
npm run dev                   # http://localhost:3000
```

À la première visite, l'écran de connexion bascule automatiquement en mode
installation : vous y créez le compte administrateur. Cette étape n'apparaît
qu'une seule fois.

### Variables d'environnement

| Variable | Rôle |
|---|---|
| `DATABASE_URL` | Chemin de la base SQLite (`file:./dev.db`). |
| `APP_SESSION_SECRET` | Signature du cookie de session. La changer déconnecte toutes les sessions. |
| `APP_ENCRYPTION_KEY` | Chiffrement des identifiants du coffre. **La perdre rend les secrets stockés définitivement illisibles.** |
| `STORAGE_DIR` | *(facultatif)* Dossier des fichiers joints. Défaut : `storage/documents`. |

Générez chaque secret avec `openssl rand -base64 32`.

### Mise en production

```bash
npm run build
npm start
```

Aucune dépendance externe : ni base de données à administrer, ni service tiers,
ni file d'attente. Un `node` et un disque suffisent.

---

## Ce que fait l'outil

### Pilotage

- **Tableau de bord** — revenus récurrents (ARR/MRR), marge nette, encours à
  encaisser, sites en ligne, pipeline, alertes actionnables, trajectoire de
  l'ARR sur douze mois, encaissements mensuels, échéances à venir, tâches.
- **Échéancier** — abonnements, domaines, hébergements et certificats SSL
  fusionnés sur une seule ligne de temps, groupés par palier de rappel
  (retard / 7 / 15 / 30 / 60 / 90 jours).

### Relation client

- **Clients** — fiche complète en six onglets : vue d'ensemble, projets,
  technique, finances, documents, suivi. Coordonnées, TVA, interlocuteurs
  multiples, notes privées, journal d'activité.
- **Prospection** — kanban glisser-déposer des entreprises sans site web.
  Déposer une fiche dans « Gagné » la convertit en client actif.
- **Devis** — éditeur de lignes avec totaux calculés en direct, remise, TVA,
  et aperçu imprimable (Ctrl+P → PDF).

### Production

- **Projets** — état d'avancement, technologies, sur mesure ou CMS, URL de
  production et de préproduction, historique daté des modifications.
- **Domaines** — registrar, expiration, prix de renouvellement, reconduction
  automatique, DNS, SSL suivi séparément, sous-domaines.
- **Hébergements** — fournisseur, offre, coût ramené à l'année, échéance,
  caractéristiques techniques.

### Revenus

- **Abonnements** — contrat, cycle, période couverte, prochain renouvellement,
  reconduction en un clic.
- **Paiements** — qui a payé, qui doit payer, pour quelle période exactement.
  Vues rapides « à encaisser », « en retard », « payés ».
- **Revenus** — MRR, ARR, et surtout la **marge réelle par client** : le prix
  facturé moins le coût des domaines et hébergements que vous payez.

### Organisation

- **Tâches** — rappels, relances, appels et rendez-vous rattachés aux clients.
- **Documents** — contrats, factures, captures, joints à un client ou un projet.
- **Recherche globale** (`⌘K` / `Ctrl+K`) — atteint n'importe quelle fiche ou
  n'importe quel écran sans quitter le clavier.

---

## Décisions d'architecture

Les choix ci-dessous expliquent le *pourquoi* ; le *comment* est commenté dans
le code, au plus près des fonctions concernées.

**Un prospect n'est pas une entité distincte.** C'est un `Client` avec
`status = PROSPECT` et une étape de pipeline. La conversion est un changement de
statut : pas de recopie de données, pas de doublon possible entre une table
« leads » et une table « clients », pas de code dupliqué entre deux écrans.

**Le contrat est séparé de ses échéances.** `Subscription` porte le contrat,
`SubscriptionPeriod` porte chaque période facturable. C'est ce qui permet de
répondre séparément à « quel est le prochain renouvellement ? » et à « qui n'a
pas payé la période de l'an dernier ? » — deux questions qu'un simple champ
« payé » ne saurait couvrir.

**Les montants sont des `Decimal`, jamais des `Float`.** Les calculs de MRR/ARR
agrègent des dizaines de lignes ; les flottants y accumulent des dérives
visibles.

**Un seul moteur d'échéances.** `lib/renewals.ts` convertit une date en palier
d'alerte, et `lib/queries/renewal-feed.ts` fusionne les quatre sources qui
expirent. Le tableau de bord, l'échéancier et les fiches client partagent donc
exactement la même logique et le même vocabulaire visuel.

**Les statuts d'impayé sont recalculés à l'ouverture des écrans concernés**,
pas par une tâche planifiée. L'outil n'a aucun processus de fond à surveiller,
et le résultat est identique : le statut est juste au moment où on le regarde.

**Les secrets ne sont jamais rendus dans le HTML.** Le coffre affiche un
libellé et un identifiant ; le mot de passe n'est déchiffré que sur clic
explicite, via une server action, et s'efface de l'état au bout de 30 secondes.

**Les fichiers vivent hors de `public/`.** Le téléchargement passe par une route
qui vérifie la session : un contrat client n'est pas accessible à qui devine
son URL.

**L'état des listes vit dans l'URL.** Un filtre ou un tri reste partageable,
rechargeable, et présent dans l'historique du navigateur.

**Les graphiques sont écrits à la main en SVG.** Aucune dépendance de
visualisation : les couleurs viennent des variables du thème, donc les
graphiques suivent le mode clair/sombre sans une ligne de JavaScript, et le
rendu reste dans la même langue visuelle que le reste de l'outil. La palette
catégorielle a été validée (bande de clarté, plancher de chroma, séparation
deutan/protan/tritan des paires adjacentes, contraste ≥ 3:1) séparément pour
chaque thème.

### Pile technique

Next.js 16 (App Router, server actions) · TypeScript · Prisma · SQLite ·
Tailwind CSS 4 · lucide-react. Composants d'interface écrits sur mesure —
aucune bibliothèque de composants.

### Organisation du code

```
prisma/schema.prisma        Modèle de données commenté
prisma/seed.ts              Jeu de démonstration

src/app/
  (app)/                    Écrans authentifiés (garde unique dans le layout)
  login/                    Connexion et installation initiale
  api/                      Téléchargement de documents, export, recherche

src/components/
  ui/                       Briques génériques (Button, Card, Modal, Table…)
  domain/                   Briques métier (échéances, coffre, devis, filtres)
  charts/                   Graphiques SVG
  layout/                   Coquille, navigation, thème, palette de commandes

src/lib/
  actions/                  Server actions, par module
  queries/                  Lectures agrégées (tableau de bord, échéances)
  constants.ts              Vocabulaire : libellés et tonalités des enums
  renewals.ts               Moteur d'échéances et conversions de cycles
  crypto.ts                 Chiffrement du coffre, hachage des mots de passe
  format.ts                 Formatage (montants, dates, tailles)
```

Le fichier `src/lib/constants.ts` mérite une mention : chaque enum Prisma y est
décrit une seule fois (libellé français + tonalité visuelle). Les badges, les
filtres, les `<select>` et les graphiques consomment ces tables — ajouter une
valeur à un enum se répercute partout, sans chasse aux chaînes en dur.

---

## Sauvegarde

Trois éléments composent l'état complet de l'outil :

1. `prisma/dev.db` — la base. Une copie de fichier suffit.
2. `storage/documents/` — les fichiers joints.
3. `.env` — sans `APP_ENCRYPTION_KEY`, le coffre est illisible.

L'écran **Réglages → Données & export** propose en plus un export JSON complet,
pour que vos données ne soient jamais captives de l'outil.

---

## Scripts

| Commande | Effet |
|---|---|
| `npm run dev` | Serveur de développement |
| `npm run build` / `npm start` | Build et exécution en production |
| `npm run typecheck` | Vérification TypeScript |
| `npm run setup` | Crée la base et le jeu de démonstration |
| `npm run db:studio` | Explorateur de base Prisma |

---

## Note sur l'existant

Le dépôt contenait auparavant `analyseur_mails.py` et son workflow GitHub
Actions, sans rapport avec cette console. Ils sont conservés en l'état.
