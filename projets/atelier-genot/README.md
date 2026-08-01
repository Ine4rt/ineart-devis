# Atelier Genot — site vitrine

Site d'une seule page, en HTML/CSS/JS purs. Aucune dépendance, aucun outil de
build, aucun serveur applicatif : le dossier tel quel **est** le site.

```
atelier-genot/
├── index.html              toute la page
├── assets/
│   ├── css/style.css       toute la mise en forme
│   ├── js/main.js          formulaire de devis + navigation active
│   ├── fonts/              Archivo (2 fichiers, 40 Ko au total)
│   └── img/                les photos — voir assets/img/README.md
└── README.md               ce fichier
```

Poids actuel sans photos : 80 Ko sur le disque, **environ 50 Ko réellement
transmis** (les fichiers texte sont compressés par l'hébergeur, les polices
le sont déjà). Avec huit photos correctement compressées, la page complète
reste sous 2 Mo — de quoi s'afficher en une seconde ou deux sur une connexion
mobile.

---

## Mise en ligne

Le site est un ensemble de fichiers statiques. Il fonctionne partout :

| Hébergement | Marche à suivre |
|---|---|
| **Hébergement mutualisé** (OVH, Combell, One.com…) | Envoyer le contenu du dossier dans `www/` ou `public_html/` par FTP |
| **Netlify / Cloudflare Pages / Vercel** | Glisser le dossier dans l'interface, ou connecter le dépôt |
| **Serveur maison / VPS** | Copier le dossier dans la racine servie par nginx ou Apache |

Il n'y a rien à configurer côté serveur : ni PHP, ni base de données, ni
variables d'environnement. Un simple hébergement statique suffit, ce qui rend
le coût de fonctionnement quasi nul.

**Test en local**, avant tout envoi :

```bash
cd atelier-genot
python3 -m http.server 8080
# puis ouvrir http://localhost:8080
```

Ouvrir `index.html` directement par double-clic fonctionne aussi, mais les
polices peuvent ne pas se charger selon le navigateur. Le petit serveur local
ci-dessus reproduit fidèlement le comportement en ligne.

---

## À remplir avant la mise en ligne

Tous les emplacements sont repérables en cherchant **`À REMPLIR`** dans
`index.html`. Voici la liste exhaustive.

### Coordonnées

| Élément | Occurrences | Où |
|---|---|---|
| **Téléphone** | 4 | bouton « Appeler » de la barre de navigation, bloc Contact, barre mobile fixe (bas d'écran), bloc JSON-LD |
| **E-mail** | 3 | bloc Contact, bloc JSON-LD, et `DESTINATAIRE` dans `assets/js/main.js` |
| **Adresse** | 2 | bloc Contact, bloc JSON-LD |
| **Horaires** | 2 | bloc Contact, bloc JSON-LD |
| **N° d'entreprise / TVA** | 1 | pied de page |

Le téléphone apparaît sous deux formes qu'il faut modifier **toutes les deux** :

```html
<a href="tel:+3240000000">+32 4XX XX XX XX</a>
     ↑ format technique          ↑ format affiché
```

Le format technique s'écrit sans espaces ni zéro initial :
`085 21 34 56` devient `tel:+3285213456`.

L'adresse e-mail du formulaire de devis se trouve à un seul endroit, en haut de
`assets/js/main.js` :

```js
var DESTINATAIRE = "contact@ateliergenot.be";
```

### Logo

La barre de navigation affiche pour l'instant un carré « AG » dessiné en CSS.
Si le client fournit son logo, remplacer dans `index.html` :

```html
<span class="logo__mark" aria-hidden="true">AG</span>
```

par :

```html
<img src="assets/img/logo.svg" alt="" height="34">
```

Un SVG est préférable (net à toutes les tailles, quelques kilo-octets). Un PNG
sur fond transparent d'au moins 200 px de haut convient également.

La **favicon** — l'icône dans l'onglet du navigateur — est actuellement un SVG
écrit directement dans `index.html` (ligne `<link rel="icon">`). Elle peut être
remplacée par un vrai fichier une fois le logo disponible.

### Photos

Voir **`assets/img/README.md`** : la liste des fichiers attendus, les formats,
le poids maximal et les conseils de prise de vue y sont détaillés.

Tant qu'une photo manque, son emplacement s'affiche en pointillés avec une
étiquette descriptive. Dès qu'une `<img>` est insérée, le cadre disparaît tout
seul — c'est géré en CSS, il n'y a rien à supprimer.

---

## Contenu à faire valider par le client

Le site est fonctionnel en l'état, mais **trois blocs contiennent des valeurs
par défaut du métier, pas des informations vérifiées.** Ne pas les mettre en
ligne sans les confirmer.

**1. Les six prestations.** Ce sont les catégories standard d'un
ferronnier-soudeur : garde-corps, portails, escaliers, structures, réparation,
pièces sur mesure. Il faut retirer ce que l'atelier ne fait pas et ajouter ce
qui manque. Une prestation affichée mais non assurée génère des appels perdus
des deux côtés.

**2. La zone d'intervention.** Les communes listées (Marchin, Huy, Modave,
Nandrin, Clavier, Tinlot, Ouffet, Héron) sont les communes voisines de Marchin.
Elles doivent correspondre au rayon que l'atelier accepte réellement.

**3. Les engagements de la section « L'atelier ».** « Devis gratuit »,
« réponse sous deux jours ouvrables », « pose non sous-traitée » sont des
promesses. Elles sont crédibles pour ce métier, mais elles engagent le client :
à confirmer une par une, ou à reformuler.

Les légendes des six réalisations (« Garde-corps de terrasse — acier
thermolaqué, remplissage verre »…) sont des exemples destinés à montrer le
niveau de détail attendu. Elles se remplacent en même temps que les photos.

---

## Comment le site est construit

### Le formulaire de devis

Il n'envoie rien à un serveur. À la validation, il ouvre le logiciel de
messagerie du visiteur avec un e-mail déjà rédigé — objet, description du
projet, nom et téléphone.

C'est un choix délibéré. Un formulaire classique exigerait un service tiers
(Formspree, EmailJS) ou un back-end : un compte de plus, un abonnement de plus,
un point de panne de plus, et des données de prospects hébergées ailleurs. Pour
un atelier qui reçoit quelques demandes par semaine, `mailto:` fait le même
travail sans rien de tout cela, et la demande arrive directement dans la boîte
du gérant.

Si le volume augmente au point de justifier un vrai formulaire, le changement
se fait en remplaçant la seule fonction concernée dans `main.js`.

### Le fonctionnement sans JavaScript

Si `main.js` ne se charge pas, la page reste entièrement utilisable : les liens
de navigation sont des ancres HTML, le téléphone est un lien `tel:`, l'e-mail
un lien `mailto:`. Seuls le pré-remplissage du devis et le surlignage de la
section courante disparaissent.

### Les polices

Deux fichiers Archivo découpés sur mesure, embarqués dans le dossier :

- `genot-display.woff2` — version resserrée (largeur 84 %), pour les titres.
  C'est la proportion des enseignes d'atelier et des marquages industriels.
- `genot-text.woff2` — version normale, pour le texte courant.

Ils ne contiennent que les caractères latins utiles au français, ce qui les
ramène à 19 Ko chacun. Ils sont servis depuis le dossier, sans requête vers
Google Fonts : la page s'affiche plus vite, et aucune donnée du visiteur ne
part vers un tiers — ce qui règle au passage la question RGPD que pose
l'utilisation directe de Google Fonts.

### Les couleurs

Trois familles, définies en haut de `style.css` :

| Rôle | Valeur | Usage |
|---|---|---|
| Fer | `#101215` → `#1e2227` | fonds sombres, forge |
| Béton | `#edebe8` | fonds clairs, lumière d'atelier |
| Braise | `#c4441c` | accent unique : boutons, liens, détails |

Les bandes alternent volontairement sombre / clair / blanc en descendant la
page : une page uniformément noire fatigue à la lecture et rend les photos
plates. Toutes les combinaisons texte/fond dépassent le seuil de contraste
AA (4,5:1).

Pour changer l'accent orange, modifier `--ember` et `--ember-bright` dans
`assets/css/style.css` : les deux variables sont utilisées partout, il n'y a
aucune autre occurrence de la couleur dans le code.

### Référencement local

Le bloc `application/ld+json` en haut d'`index.html` décrit l'entreprise au
format `LocalBusiness` de schema.org. C'est ce que lisent les moteurs de
recherche pour associer l'atelier à des requêtes du type « ferronnier près de
Huy » et pour alimenter la fiche d'établissement.

**Il doit être rempli avec les vraies coordonnées** — un JSON-LD contenant
`"streetAddress": "À REMPLIR"` est pire que pas de JSON-LD du tout.

Le référencement local repose surtout sur deux choses hors du site : une fiche
Google Business complète et cohérente avec la page, et des avis clients. Le
site en est le point d'ancrage, pas le moteur.

---

## Accessibilité et compatibilité

- Contrastes conformes AA sur l'ensemble de la page.
- Navigation entièrement possible au clavier, avec un contour de focus visible.
- Les animations respectent `prefers-reduced-motion` : elles disparaissent pour
  les visiteurs qui l'ont activé.
- Vérifié dans Chromium en 390 px et 1440 px de large, sans aucune erreur de
  console. Le code n'utilise que des fonctionnalités standard supportées par
  Chrome, Firefox et Safari depuis 2023 (`:has()` étant la plus récente) ;
  un passage sur un vrai iPhone reste recommandé avant la mise en ligne.
- La barre fixe en bas d'écran mobile — « Appeler » et « Demander un devis » —
  place les deux actions utiles à portée de pouce en permanence. C'est le
  détail qui transforme le plus de visites en appels sur un site d'artisan.
