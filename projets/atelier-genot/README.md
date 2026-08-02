# Atelier Genot — site vitrine

Site d'une seule page, en HTML/CSS/JS purs. Aucune dépendance, aucun outil de
build, aucun serveur applicatif : le dossier tel quel **est** le site.

```
atelier-genot/
├── index.html                  toute la page
├── assets/
│   ├── css/style.css           toute la mise en forme
│   ├── js/main.js              devis, navigation, compte à rebours
│   ├── fonts/                  Archivo (2 fichiers, 40 Ko)
│   └── img/                    logo, icônes et photos
├── build-embed.mjs             génère les versions « un seul fichier »
└── README.md                   ce fichier
```

Poids transmis à la première visite : **environ 800 Ko**, dont la quasi-totalité
en photos. Les images des sections basses ne se chargent qu'au défilement, donc
le premier écran arrive en moins d'une seconde même en 4G.

---

## Mise en ligne

| Hébergement | Marche à suivre |
|---|---|
| **Mutualisé** (OVH, Combell, One.com…) | Envoyer le contenu du dossier dans `www/` ou `public_html/` par FTP |
| **Netlify / Cloudflare Pages / Vercel** | Glisser le dossier dans l'interface |
| **VPS** | Copier le dossier dans la racine servie par nginx ou Apache |

Envoyez le **contenu** du dossier, pas le dossier lui-même : `index.html` doit
se retrouver à la racine, avec `assets/` à côté.

Rien à configurer côté serveur : ni PHP, ni base de données, ni variables
d'environnement. Un hébergement statique suffit, ce qui rend le coût de
fonctionnement quasi nul.

**Test en local** avant tout envoi :

```bash
cd atelier-genot
python3 -m http.server 8080
# puis ouvrir http://localhost:8080
```

Ouvrir `index.html` par double-clic fonctionne aussi, mais certains navigateurs
refusent alors de charger les polices. Le serveur local ci-dessus reproduit
fidèlement le comportement en ligne.

### Versions « un seul fichier »

```bash
node build-embed.mjs
```

Produit deux fichiers, utiles pour envoyer un aperçu par e-mail ou le déposer
là où un dossier n'est pas accepté :

- `atelier-genot.embed.html` — page complète autonome, à ouvrir directement.
- `atelier-genot.artifact.html` — le même contenu sans `<html>`/`<head>`/`<body>`,
  pour les hôtes d'aperçu qui fournissent eux-mêmes l'enveloppe.

Les deux embarquent polices et photos en `data:` URI ; les photos y sont
recompressées, car un fichier unique se télécharge d'un bloc. Le dossier reste
la version de référence — ces fichiers n'en sont qu'un tirage.

---

## Le compte à rebours de démonstration

Le site s'annonce comme une proposition et affiche le temps qu'il lui reste.
Passé la date, la page est remplacée par un écran d'expiration.

Tout se règle en tête de `assets/js/main.js` :

```js
var DEMO = {
  actif: true,
  expire: "2026-08-08T20:00:00+02:00",   // heure belge : +02:00 l'été, +01:00 l'hiver
  contact: "info@ineart.be",
};
```

**Pour livrer le site pour de bon**, passez `actif` à `false`. La bannière
disparaît, plus rien ne compte, le reste du site est inchangé.

### Ce que ce compte à rebours est, et ce qu'il n'est pas

Il s'exécute dans le navigateur du visiteur. Quelqu'un qui désactive JavaScript
ou recule l'horloge de sa machine verra la page malgré tout. **C'est une
échéance commerciale, pas une serrure.** Pour la démarche visée — montrer qu'une
proposition a une durée — c'est suffisant et honnête.

Si vous voulez que l'accès se ferme réellement, il faut agir côté serveur. Le
plus simple, sur un hébergement qui donne accès à `cron` :

```bash
# retire le site le 8 août à 20 h
echo 'rm -rf /var/www/genot/*' | at 20:00 2026-08-08
```

Sur un hébergement mutualisé sans `cron`, la solution reste manuelle : notez la
date et supprimez le dossier ce jour-là.

### Pourquoi la bannière ne se contente pas de compter

Elle dit aussi qui a fait la page : *« Proposition de site réalisée par IneWeb
pour l'Atelier Genot »*. Sans cette phrase, une maquette portant le nom, le logo
et le vrai numéro de téléphone d'une entreprise pourrait passer pour son site
officiel — auprès du client comme auprès de quiconque reçoit le lien. La mention
protège les deux parties et ne coûte rien à la démonstration.

---

## Ce qui est renseigné

Les coordonnées proviennent de la **page Facebook officielle de l'atelier**
(relevée en août 2026) :

| Donnée | Valeur | Où elle apparaît |
|---|---|---|
| Téléphone | 0494 13 35 05 | nav, contact, barre mobile, JSON-LD |
| E-mail | ateliergenot@hotmail.com | contact, JSON-LD, `main.js` |
| Adresse | Rue du Frêne 8b, 4570 Marchin | contact, JSON-LD |
| Facebook | facebook.com/ateliergenot | galerie, pied de page, JSON-LD |
| WhatsApp | wa.me/32494133505 | contact |

Le téléphone s'écrit sous deux formes qu'il faut modifier **toutes les deux**
si le numéro change :

```html
<a href="tel:+32494133505">0494 13 35 05</a>
     ↑ format technique       ↑ format affiché
```

L'adresse du formulaire de devis se trouve à un seul endroit, en tête de
`assets/js/main.js` (`DESTINATAIRE`).

---

## Ce qui reste à confirmer avec le client

**1. Le numéro d'entreprise / TVA.** Le pied de page affiche
`TVA BE 0XXX.XXX.XXX`. C'est une mention légale obligatoire en Belgique pour un
site professionnel — à compléter avant la mise en ligne.

**2. Les horaires.** La page Facebook indique « toujours ouvert », ce qui ne
veut rien dire sur un site. J'ai écrit lundi–vendredi 7 h – 18 h, samedi sur
rendez-vous, avec une ligne « urgence ». À valider ou corriger.

**3. Les six prestations.** Facebook cite *escalier, garde-corps, palissade* et
« projets sur mesure ». J'ai complété avec les catégories standard du métier —
structures, réparation/soudure — et ajouté **découpe décorative**, que les
photos rendent évidente. Retirer ce que l'atelier ne fait pas, ajouter ce qui
manque : une prestation affichée mais non assurée fait perdre du temps aux deux
parties.

**4. La zone d'intervention.** Les huit communes listées sont les voisines de
Marchin, pas un rayon vérifié.

**5. Les engagements.** « Devis gratuit », « pose non sous-traitée », « plan
d'atelier soumis avant fabrication » sont des promesses crédibles pour ce
métier, mais elles engagent le client. À confirmer une par une.

---

## Les photos

Deux photos ont été fournies : le portail à découpes florales et la plaque de
maison en corten. Elles alimentent **six emplacements** — accueil, quatre tuiles
de réalisations, section atelier — par recadrages (vue d'ensemble, détail du
lettrage, motif ajouré, vantail en hauteur). C'est honnête, ce sont bien leurs
ouvrages, mais cela reste **deux chantiers**.

Le site respire dès qu'il y en a davantage. `assets/img/README.md` détaille les
formats, les poids et les conseils de prise de vue. Trois ajouts changeraient
tout : un escalier, un garde-corps, et une photo de l'artisan au travail — celle
qui crée la confiance, et celle qui manque presque toujours.

Le **logo** est extrait de la photo de profil Facebook, détouré en PNG
transparent (`assets/img/logo.png`, version sombre pour fond clair dans
`logo-sombre.png`). Il est net à la taille où il est affiché, mais il vient
d'un JPEG : **demandez le fichier vectoriel** au client si son graphiste l'a
encore. Un SVG resterait parfait à toutes les tailles et à l'impression.

---

## Comment le site est construit

### Le formulaire de devis

Il n'envoie rien à un serveur. À la validation, il ouvre le logiciel de
messagerie du visiteur avec un e-mail déjà rédigé — objet, type d'ouvrage,
description, nom et téléphone.

C'est un choix délibéré. Un formulaire classique exigerait un service tiers
(Formspree, EmailJS) ou un back-end : un compte de plus, un abonnement de plus,
un point de panne de plus, et des coordonnées de prospects hébergées ailleurs.
Pour un atelier qui reçoit quelques demandes par semaine, `mailto:` fait le même
travail sans rien de tout cela, et la demande arrive directement dans la boîte
du gérant.

Si le volume augmente au point de justifier un vrai formulaire, le changement
tient dans une seule fonction de `main.js`.

### Le fonctionnement sans JavaScript

Si `main.js` ne se charge pas, la page reste entièrement utilisable : les liens
de navigation sont des ancres HTML, le téléphone un lien `tel:`, l'e-mail un
lien `mailto:`. Disparaissent seulement le pré-remplissage du devis, le
surlignage de la section courante — et la bannière de démonstration, qui n'est
créée que par le script.

### Les polices

Deux fichiers Archivo découpés sur mesure, embarqués dans le dossier :

- `genot-display.woff2` — version resserrée (largeur 84 %), pour les titres.
  C'est la proportion des enseignes d'atelier et des marquages industriels.
- `genot-text.woff2` — version normale, pour le texte courant.

Ils ne contiennent que les caractères latins utiles au français, ce qui les
ramène à 20 Ko chacun. Servis depuis le dossier, sans requête vers Google
Fonts : la page s'affiche plus vite, et aucune donnée du visiteur ne part vers
un tiers — ce qui règle au passage la question RGPD que pose l'usage direct de
Google Fonts.

### Les couleurs

Trois familles, définies en tête de `style.css` :

| Rôle | Valeur | Usage |
|---|---|---|
| Fer | `#101215` → `#1e2227` | fonds sombres, forge |
| Béton | `#edebe8` | fonds clairs, lumière d'atelier |
| Braise | `#c4441c` | accent unique : boutons, liens, détails |

La braise n'a pas été choisie au hasard : c'est la couleur de la plaque en
corten sur la photo d'accueil. Le site et l'ouvrage sont dans le même ton.

Les bandes alternent volontairement sombre / clair / blanc en descendant la
page. Une page uniformément noire fatigue à la lecture et aplatit les photos.
Toutes les combinaisons texte/fond dépassent le seuil de contraste AA (4,5:1).

Pour changer l'accent, modifiez `--ember` et `--ember-bright` : les deux
variables sont utilisées partout, il n'y a aucune autre occurrence de la
couleur dans le code.

### La mosaïque des réalisations

Les tailles sont portées par des classes — `.work--large` (2 colonnes × 2
rangées), `.work--tall` (2 rangées) — et non par des `:nth-child`. Ajouter ou
retirer une photo ne casse donc pas la composition : il suffit de déplacer la
classe. Une grande tuile, une verticale et deux petites remplissent un
rectangle sans laisser de case vide, aussi bien sur quatre colonnes que sur
deux.

### Référencement local

Le bloc `application/ld+json` en tête d'`index.html` décrit l'entreprise au
format `LocalBusiness` de schema.org. C'est ce que lisent les moteurs pour
associer l'atelier à des requêtes comme « ferronnier près de Huy » et pour
alimenter la fiche d'établissement.

Le référencement local repose surtout sur deux choses hors du site : une fiche
Google Business complète et cohérente avec la page, et des avis clients. La
page Facebook affiche « pas encore évalué » — c'est le levier le plus rentable
à activer, avant toute autre chose.

---

## Accessibilité et compatibilité

- Contrastes conformes AA sur l'ensemble de la page.
- Navigation entièrement possible au clavier, avec un contour de focus visible.
- Chaque photo porte un texte alternatif décrivant l'ouvrage réel.
- Les animations respectent `prefers-reduced-motion` : elles disparaissent pour
  les visiteurs qui l'ont activé.
- Vérifié dans Chromium en 390 px et 1440 px, sans aucune erreur de console.
  Le code n'utilise que des fonctionnalités standard supportées par Chrome,
  Firefox et Safari depuis 2023 (`:has()` étant la plus récente) ; un passage
  sur un vrai iPhone reste recommandé avant la mise en ligne.
- La barre fixe en bas d'écran mobile — « Appeler » et « Demander un devis » —
  place les deux actions utiles à portée de pouce en permanence. C'est le
  détail qui transforme le plus de visites en appels sur un site d'artisan.
