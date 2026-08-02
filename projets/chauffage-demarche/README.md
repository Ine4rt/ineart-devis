# Chauffage Demarche Pascal — site vitrine

Site d'une seule page, en HTML/CSS/JS purs. Aucune dépendance, aucun outil de
build, aucun serveur applicatif : le dossier tel quel **est** le site.

```
chauffage-demarche/
├── index.html              toute la page
├── assets/
│   ├── css/style.css       toute la mise en forme
│   ├── js/main.js          formulaire, navigation, compte à rebours
│   ├── fonts/              Fraunces et Karla (56 Ko)
│   └── img/                photos de chantier
├── DEMARCHAGE.md           mails de prospection
└── README.md               ce fichier
```

Poids transmis à la première visite : **environ 350 Ko**, dont l'essentiel en
photos. Les images du bas ne se chargent qu'au défilement.

---

## Mise en ligne

Envoyez le **contenu** du dossier dans le répertoire du domaine ou du
sous-domaine : `index.html` à la racine, `assets/` à côté.

**Test en local** avant tout envoi :

```bash
cd chauffage-demarche
python3 -m http.server 8080
```

---

## Ce qui manque

**1. Les horaires.** Absents de la fiche Google comme de Facebook. La fiche
Google affiche même « Ajouter des horaires d'ouverture », ce qui signale qu'il
ne les a jamais renseignés. J'ai écrit « en semaine, sur appel » — formulation
prudente, à corriger.

**2. Le numéro de TVA.** Mention légale obligatoire en Belgique. Actuellement
`TVA BE 0XXX.XXX.XXX` dans le pied de page.

**3. Le logo.** Il en a un — une flamme orange et des gouttes bleues, peint sur
son camion. Je n'en ai qu'une photo floue, donc le site utilise un symbole SVG
provisoire construit dans le même esprit. **Demandez le fichier d'origine** :
son graphiste ou son enseigniste l'a forcément.

**4. L'adresse exacte de sa page Facebook.** Le lien « autres chantiers » pointe
pour l'instant vers une recherche Facebook. Un identifiant inventé aurait envoyé
les visiteurs chez quelqu'un d'autre.

### À valider avec lui

- **Les quatre métiers** reprennent ce qu'annonce sa page Facebook
  (entretien-dépannage / chauffage / sanitaire / rénovation de salle de bain).
  Le détail des prestations, en revanche, est déduit du métier.
- **Les communes desservies** sont les voisines de Huy, pas un rayon vérifié.
- **L'engagement de dépannage** — « nous donnons une heure de passage au
  téléphone, pas une fourchette de trois jours » — est le genre de promesse qui
  vend, mais elle l'engage. À confirmer, ou à retirer.
- **Le carrelage.** Le site affirme qu'il pose lui-même le carrelage de la salle
  de bain. Sa publication « salle de bain terminée à Hannut » le suggère, mais
  beaucoup de chauffagistes sous-traitent. Point à vérifier en priorité : c'est
  l'argument central de la page.

---

## Les avis Google

Sa fiche porte **7 avis, moyenne 3,9/5**. Le site cite deux avis réels — un 5/5
et un 4/5 — et affiche « Sept avis sur Google » avec un lien vers la fiche.

**La note moyenne n'est pas affichée**, et c'est un choix à assumer. 3,9 n'est
pas un mauvais score, mais ce n'est pas un argument de vente ; les deux avis
cités, eux, en sont. Rien n'est caché pour autant : le lien mène directement à
la fiche complète, où la note est visible. C'est ce que fait n'importe quel site
sérieux — citer, pas maquiller.

Les auteurs sont désignés par leur prénom et l'initiale de leur nom
(« Emmanuel D. »), alors que Google affiche le nom complet. C'est l'usage sur
les sites d'entreprise, et c'est plus respectueux de gens qui ont écrit un avis
sans imaginer se retrouver sur une page commerciale.

**Le levier le plus rentable du dossier** : il a 7 avis. Vingt avis à 4,5
changeraient sa position dans les recherches locales bien plus qu'une
optimisation du site. Demander un avis à chaque client satisfait ne coûte rien.

---

## Les photos

Les quatre images viennent d'une publication Facebook d'août 2023 — « Réalisation
d'un chauffage sol ». Elles sont recadrées pour éviter les éléments d'interface.

Il manque l'essentiel : **aucune photo de salle de bain terminée**, alors que
c'est la moitié de son activité et la plus rentable. Sa page Facebook en
contient (« Salle de bain terminée à Hannut ») ; elles n'étaient pas
exploitables dans les captures dont je disposais.

Quatre photos à demander :
1. Une salle de bain terminée, en lumière du jour.
2. Une chaudière installée proprement, tuyauterie visible.
3. Le camion — il porte son logo et son numéro, c'est sa vitrine mobile.
4. Lui, au travail. Dans un métier où l'on fait entrer quelqu'un chez soi, le
   visage compte plus qu'ailleurs.

---

## Le compte à rebours de démonstration

Réglé en tête de `assets/js/main.js` :

```js
var DEMO = { actif: true, expire: "2026-08-10T20:00:00+02:00", contact: "info@ineart.be" };
```

`actif: false` livre le site pour de bon. Le décompte s'exécute dans le
navigateur du visiteur : **c'est une échéance commerciale, pas une serrure.**
Pour fermer réellement l'accès, retirez les fichiers du serveur.

---

## Les partis pris de conception

### Épuré veut dire enlever, pas alléger

Une seule colonne de lecture, presque aucun cadre, aucune carte empilée. Les
séparations viennent de l'espace et de rares filets d'un pixel. La seule chose
qui échappe à la colonne, ce sont les photographies — un chantier se juge en
grand.

### Le rythme des bandes

```
blanc   accueil          la photo fait l'événement
brume   nos métiers      gris tiré vers l'eau
NOIR    dépannage        bande pleine largeur, titre orange
blanc   réalisations     les photos respirent
BLEU    avis             son meilleur atout, sa propre zone
sable   zone d'interv.   gris tiré vers la flamme
blanc   contact          le formulaire doit être calme
NOIR    pied de page
```

Une page d'un seul ton se lit comme un document ; une page qui change de sol à
chaque section se parcourt. Les deux gris pâles — l'un vers le bleu, l'autre
vers l'orange — sont presque indiscernables isolément, mais il suffit qu'ils
soient différents pour que deux sections voisines ne se confondent pas.

Les trois zones fortes sont espacées, jamais accolées : noir, puis du clair,
puis bleu. Deux aplats sombres à la suite feraient un bloc.

C'est le troisième registre de la série, et il est distinct des deux autres :
la ferronnerie était sombre et photographique, le garage dense et informatif,
celui-ci est clair et aéré.

### Ses trois couleurs, chacune son domaine

Relevées sur son camion : un bandeau noir, une flamme orange, des gouttes bleues.

- **Noir `#111418`** — la structure : l'encre, le bloc dépannage, le pied de page.
- **Flamme `#e2601c`** — la chaleur : chauffage, chaudière, urgence.
- **Bleu `#1f6fb2`** — l'eau : sanitaire, salle de bain.

Aucune ne sert d'ornement, et c'est ce qui permet d'en tenir trois sans que la
page devienne bariolée : un lecteur qui voit de l'orange sait qu'on parle de
chaleur avant même d'avoir lu le mot.

Le bloc dépannage est le seul aplat plein de la page — noir, titre orange,
bouton orange. C'est exactement l'ordre des couleurs sur son camion.

Le noir porte une pointe de bleu plutôt qu'un `#000` : le noir absolu écrase sur
écran, et les gris qui en descendent gardent ainsi une parenté avec l'eau.

**Deux valeurs d'orange**, parce qu'un même orange ne peut pas tout faire. La
vive (`#e2601c`) sert de couleur d'écriture sur le noir ; la sourde
(`#c44f13`) sert de fond sous du texte blanc. L'orange exact du camion tombe à
3,6:1 avec du blanc dessus — juste, mais illisible pour qui a la vue basse.

### Une serif, pour un chauffagiste

Fraunces, figée en taille optique de titrage et légèrement adoucie. Le réflexe
serait une grotesque technique ; mais la moitié de son métier, c'est la salle de
bain — du confort domestique, pas de la tuyauterie industrielle. La serif dit
la maison. Karla, humaniste, porte le texte.

### Le dépannage isolé

Une chaudière en panne en hiver est le seul moment où l'on cherche un
chauffagiste dans l'urgence. Ce cas a donc son propre bloc, et c'est la seule
zone chaude de la page.

### Le formulaire compose un e-mail

Pas de serveur, pas d'abonnement à un service tiers, pas de coordonnées de
clients hébergées ailleurs. L'objet du message reprend la nature de la demande :
« Une panne de chauffage » et « Une rénovation de salle de bain » n'appellent
pas la même réactivité, et le voir sans ouvrir le message a de la valeur.

---

## Accessibilité et compatibilité

- Contrastes conformes AA sur l'ensemble de la page.
- Navigation entièrement possible au clavier, avec un contour de focus visible.
- Chaque photo porte un texte alternatif décrivant le chantier réel.
- Les animations respectent `prefers-reduced-motion`.
- Vérifié dans Chromium en 390 px et 1440 px, sans erreur de console. Un passage
  sur un vrai iPhone reste recommandé avant la mise en ligne.
- La barre fixe mobile place « Appeler » et « Demander un devis » à portée de
  pouce en permanence.
