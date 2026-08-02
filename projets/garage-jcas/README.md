# Garage J.C.A.S. — site vitrine

Site d'une seule page, en HTML/CSS/JS purs. Aucune dépendance, aucun outil de
build, aucun serveur applicatif : le dossier tel quel **est** le site.

```
garage-jcas/
├── index.html              toute la page
├── assets/
│   ├── css/style.css       toute la mise en forme
│   ├── js/main.js          sélection des besoins, demande, compte à rebours
│   ├── fonts/              Archivo élargi, Archivo, JetBrains Mono (76 Ko)
│   └── img/                photos de l'atelier
├── DEMARCHAGE.md           mails de prospection
└── README.md               ce fichier
```

Poids transmis à la première visite : **environ 150 Ko**. La page s'affiche en
moins d'une seconde, y compris sur un téléphone en 4G.

---

## Mise en ligne

| Hébergement | Marche à suivre |
|---|---|
| **Mutualisé** (One.com, OVH, Combell…) | Envoyer le **contenu** du dossier dans le répertoire du domaine ou du sous-domaine |
| **Netlify / Cloudflare Pages** | Glisser le dossier dans l'interface |
| **VPS** | Copier le dossier dans la racine servie par nginx ou Apache |

`index.html` doit se retrouver à la racine, avec `assets/` à côté — pas dans un
sous-dossier `garage-jcas/`.

**Test en local** avant tout envoi :

```bash
cd garage-jcas
python3 -m http.server 8080
# puis ouvrir http://localhost:8080
```

---

## Ce qui manque, et qui bloque la mise en ligne

Trois informations n'existent nulle part publiquement. Le site fonctionne sans
elles, mais il ne devrait pas être livré ainsi.

**1. L'adresse exacte.** Ni la page Facebook ni la fiche marchin.be ne donnent
la rue et le numéro — seulement « Marchin, 4570 ». Cherchez `À REMPLIR` dans
`index.html` : l'adresse apparaît à trois endroits, dont le bloc JSON-LD.
Sans elle, Google ne peut pas placer le garage sur la carte, ce qui annule une
bonne part de l'intérêt du site.

**2. Les horaires.** Facebook affiche « Fermé », idgarages « horaires non
communiqués par le garage ». J'ai écrit « du lundi au vendredi, sur appel »,
formulation prudente et probablement fausse. À corriger.

**3. Le numéro de TVA.** Mention légale obligatoire en Belgique pour un site
professionnel. Actuellement `TVA BE 0XXX.XXX.XXX` dans le pied de page.

### Autres points à valider avec le garage

- **Les huit prestations** de la grille sont les catégories standard d'un
  garage généraliste. Retirer ce qu'il ne fait pas — la carrosserie notamment
  est souvent sous-traitée.
- **Les communes desservies** sont les voisines de Marchin, pas un rayon
  vérifié.
- **Les engagements** — « devis avant travaux », « pièces remplacées montrées
  sur demande », « prix des pièces d'origine et équivalentes » — sont crédibles
  pour un indépendant, mais ils l'engagent. À confirmer un par un.
- **« Indépendant Renault »** est repris de l'enseigne visible sur sa vitrine.
  Vérifier que le statut est toujours d'actualité, ce type d'accord évoluant.

---

## Les photos

Les deux images viennent d'une **capture d'écran de sa page Facebook** :
c'était la seule source disponible. Elles sont recadrées pour éviter les
éléments d'interface, et la plaque d'immatriculation d'un client a été floutée
— une donnée personnelle de tiers n'a rien à faire sur le site de son garage.

Leur définition est faible. **Demandez de vraies photos** avant la mise en
ligne. Quatre suffisent :

1. La façade avec l'enseigne, en journée.
2. L'atelier avec une voiture sur le pont.
3. Un véhicule d'occasion propre, s'il en vend.
4. Le gérant. C'est la photo qui crée la confiance dans un métier où le client
   se demande à qui il confie sa voiture.

Le site est conçu pour tenir sans photographie : c'est ce qui le distingue d'un
site de ferronnier ou de menuisier, où l'image est l'argument. Ici l'argument
est l'information — mais de bonnes photos le renforceraient.

---

## Le compte à rebours de démonstration

Réglé en tête de `assets/js/main.js` :

```js
var DEMO = {
  actif: true,
  expire: "2026-08-09T20:00:00+02:00",   // heure belge : +02:00 l'été
  contact: "info@ineart.be",
};
```

`actif: false` livre le site pour de bon : la bannière disparaît, plus rien ne
compte, le reste est inchangé.

Le décompte s'exécute dans le navigateur du visiteur. Qui désactive JavaScript
ou recule son horloge verra la page malgré tout : **c'est une échéance
commerciale, pas une serrure.** Pour fermer réellement l'accès, retirez les
fichiers du serveur à la date voulue.

La bannière nomme aussi l'auteur — « proposition de site pour le Garage
J.C.A.S. ». Sans cette mention, une maquette portant le nom et le vrai numéro
d'une entreprise pourrait passer pour son site officiel.

---

## Les partis pris de conception

### Une fiche, pas une brochure

Quand on cherche un garage, on veut savoir quatre choses : est-ce qu'ils font
ma marque, combien ça coûte, quand puis-je venir, quel est le numéro. La page
est donc construite en blocs d'information nets, avec des filets visibles et
des étiquettes en chasse fixe — le registre de la fiche d'atelier.

C'est l'inverse d'un site de ferronnerie, qui vit de ses photos. Deux métiers
voisins ne se vendent pas de la même façon.

### Fond clair

Un atelier propre et éclairé inspire confiance ; un atelier sombre inspire
l'inverse. Le fond clair n'est pas un défaut d'audace, c'est l'argument.

### Le jaune ne s'écrit jamais

`--jaune: #ffc60a` ne sert que d'aplat : bandeau d'engagement, surlignage du
titre, bouton principal, numéros d'étape. Jamais de texte jaune.

Cette règle tient tout le système. Elle interdit d'emblée les jaunes
illisibles, et elle donne à la page une seule façon d'attirer l'œil — comme un
marquage au sol dans un atelier.

### Le numéro en grand

Sur un site de garage, l'appel est l'action principale. Le numéro est donc
affiché en gros dès le premier écran, en chasse fixe, et répété dans la barre
de navigation, dans le pied de page et dans la barre fixe mobile.

### La grille de besoins

Le visiteur ne cherche pas « la liste des prestations », il cherche ce qui ne
va pas sur sa voiture. Il coche donc — Freins, Pneus, Voyant allumé — et la
demande se rédige toute seule dans sa messagerie.

Cela change ce qui arrive au garage : au lieu d'un « bonjour, c'est pour un
devis », il reçoit un message qui nomme le besoin, le véhicule et un numéro de
téléphone. Sans script, ces cases restent des cases à cocher ordinaires dans un
formulaire qui fonctionne.

### Les trois étapes numérotées

Ce sont les seuls numéros de la page, parce que c'est la seule vraie séquence :
on ne répare pas avant d'avoir chiffré. Ailleurs, numéroter serait décoratif.

### La typographie

- **Titres** : Archivo élargi (largeur 118 %), en bas de casse.
- **Texte** : Archivo à largeur normale.
- **Données** : JetBrains Mono — téléphone, horaires, étiquettes, numéros
  d'étape. La chasse fixe aligne les chiffres et donne le ton de la fiche
  technique.

Les trois fichiers sont découpés aux seuls caractères du français, servis
depuis le dossier. Aucune requête vers Google Fonts : page plus rapide, et
question RGPD réglée.

### Le formulaire

Il compose un e-mail plutôt que d'appeler un service tiers. Pas de serveur, pas
d'abonnement, pas de coordonnées de clients hébergées ailleurs — et la demande
arrive directement dans la boîte du garage.

---

## Accessibilité et compatibilité

- Contrastes conformes AA sur l'ensemble de la page.
- Navigation entièrement possible au clavier ; les cartes de besoin sont de
  vraies cases à cocher, focalisables et utilisables sans souris.
- Les animations respectent `prefers-reduced-motion`.
- Vérifié dans Chromium en 390 px et 1440 px, sans erreur de console. Le code
  n'utilise que des fonctionnalités standard supportées par Chrome, Firefox et
  Safari depuis 2023 ; un passage sur un vrai iPhone reste recommandé.
- La barre fixe mobile place « Appeler » et « Demander un RDV » à portée de
  pouce en permanence.
