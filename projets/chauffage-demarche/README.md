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

### Ce qui a été refait, et pourquoi

Une première version, blanche et très aérée avec une serif éditoriale, a été
rejetée par le client. Le diagnostic était juste : élégante, mais froide, et
sans rapport avec un artisan qui se déplace en camionnette. La page ressemblait
à une revue de décoration, pas à quelqu'un qu'on appelle un dimanche parce que
la chaudière a lâché.

La direction actuelle vient de la base de styles `ui-ux-pro-max`, qui recommande
pour ce type d'entreprise le registre **Vibrant & Block-based** : blocs pleins,
contraste maximal, typographie énorme. Le site est bâti comme une signalétique
de chantier.

### Le système en une règle

Chaque zone est un bloc plein, et sa couleur de fond décide de la couleur du
texte. Il n'y a rien d'autre à retenir :

| Fond | Texte | Contraste |
|---|---|---|
| noir `#0c0d10` | blanc | 19,4:1 |
| orange `#f2570a` | **noir** | 5,7:1 |
| bleu `#0b57bd` | blanc | 6,8:1 |
| blanc / papier | noir | 17–19:1 |

**L'orange porte du texte noir, pas blanc.** C'est ce qui donne le registre du
panneau de chantier — orange et noir — et c'est aussi le seul choix lisible :
le blanc n'y tient que 3,4:1, sous le seuil d'accessibilité.

Les trois couleurs sont les siennes, relevées sur son camion. Chacune garde son
domaine : l'orange pour la chaleur et l'urgence, le bleu pour l'eau, le noir
pour la structure.

### Les quatre pavés de métier

Le cœur du système. Chaque activité occupe un pavé entier de couleur — orange,
bleu, papier, noir — assez grand pour se lire de loin. Deux colonnes et non
quatre : à quatre, les pavés deviennent des vignettes et l'effet disparaît.

### Le rythme

```
noir     accueil + trois chiffres
blanc    les quatre pavés de métier
ORANGE   dépannage, bande pleine largeur
papier   chantiers
BLEU     avis
blanc    zone d'intervention
noir     contact
noir     pied de page
```

### La typographie

- **Bricolage Grotesque**, figée en taille optique d'affiche (opsz 96), de 600
  à 800. Ses terminaisons coupées tiennent le gros corps sans devenir molles —
  ce qu'il faut pour des titres qui remplissent un bloc de couleur. Interlignes
  à 0,94, interlettrage à −0,035em : le titre doit occuper sa zone.
- **Public Sans** pour le texte. Dessinée pour l'administration américaine,
  faite pour être lue par tout le monde, sans style à défendre. C'est
  exactement le rôle qu'on lui demande à côté de titres aussi présents.

### Les boutons sont rectangulaires

Un rayon arrondi adoucirait un système qui tient justement par la franchise de
ses angles.

### Le mouvement

Deux gestes seulement. Le premier écran monte au chargement ; les quatre pavés
de métier entrent en cascade décalée de 60 ms quand ils arrivent dans le champ,
dans l'ordre de lecture.

Le CSS ne masque les pavés que si l'attribut `data-cascade` vaut `attente`, et
seul le script pose cette valeur — juste avant d'installer l'observateur qui
saura les révéler. Sans JavaScript, ou sans `IntersectionObserver`, les blocs
restent simplement visibles. Une animation d'apparition qui laisse le contenu
invisible en cas d'échec est le défaut le plus courant du procédé.

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
