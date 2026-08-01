# Page de présentation

Page commerciale destinée au démarchage : un lien à envoyer à une entreprise qui
n'a pas encore de site.

Elle est **volontairement séparée de la console**. La console est privée ; cette
page est publique. Aucun code n'est partagé, donc rien de ce que vous modifiez
ici ne peut casser votre outil de gestion.

## Utilisation

`index.html` est un document autonome de ~90 Ko : aucune requête réseau, aucune
dépendance externe. Déposez-le sur n'importe quel hébergement statique et c'est
en ligne.

## Modifier le contenu

Le fichier source est **`page.body.html`**. Après modification :

```bash
node presentation/build.mjs
```

La construction régénère `index.html` (document complet) et `page.embed.html`
(le même contenu sans enveloppe). Les polices sont réinjectées en base64 à
chaque construction.

## Décisions qui gouvernent cette page

Elles viennent de vous et lient toute rédaction future :

- **Aucun prix affiché.** Le tarif se propose par e-mail, au cas par cas, puis se
  valide ensemble. La question « combien ça coûte ? » est traitée dans la FAQ,
  sans chiffre.
- **Aucune mention de rendez-vous, d'appel ou de créneau.** Tout le parcours
  passe par l'e-mail. C'est aussi ce que raconte la section « Tout se passe par
  e-mail » : vous écrivez, je propose, on valide.
- **Aucun numéro de téléphone.** Une seule adresse : `info@ineart.be`.
- **Rien d'inventé** : ni témoignage, ni logo client, ni statistique. Une section
  « réalisations » ne pourra exister qu'une fois de vrais sites en ligne.

## La voix

**IneWeb est une société.** Toute la page dit « nous ». Aucune formulation à la
première personne, aucun « je crée des sites » : le registre est celui d'un
studio qui prend en charge, pas d'un prestataire qui exécute.

## Le monde visuel

Graphite mat, filets d'un pixel, larges respirations. Une seule graisse de
grotesque (Archivo) menée de 300 à 620 : **la hiérarchie passe par l'échelle et
le blanc, jamais par le gras**.

L'orange ne sert qu'à marquer — un carré de 8 px en tête de section, un
soulignement sous l'adresse e-mail, l'indicateur de diapositive active. Aucun
aplat pleine largeur : c'est précisément ce qui alourdissait la version
précédente.

## La bannière

Elle ne montre pas des photos de banque d'images. Elle montre **des sites**,
rendus en HTML dans un cadre navigateur, qui défilent. Trois mises en page —
commerce de proximité, atelier technique, profession libérale — chacune avec sa
propre identité, son vrai texte et sa vraie hiérarchie.

Ce sont des **démonstrations**, annoncées comme telles sous le carrousel et
servies par des adresses en `-exemple.be`. Ce ne sont pas des clients.

Les zones marquées « photo — … » sont les emplacements d'image. Remplacez-les
par de vraies prises de vue dès que vous en aurez : la mention en clair vaut
mieux qu'un rectangle gris qui ferait croire à un site inachevé.

## Ce qui a été coupé

Une passe de distillation a retiré environ la moitié du texte :

- **Deux sections fusionnées.** « Infrastructure » et « Compris dans le suivi
  annuel » énuméraient les mêmes éléments — nom de domaine, hébergement,
  certificat. Il n'en reste qu'une, « Ce que nous prenons en charge ».
- **Les paragraphes descriptifs deviennent des listes courtes.** On scanne une
  liste ; on subit un paragraphe. C'est vrai partout, et décisif sur téléphone.
- **Le héros tient en une phrase**, une promesse et une adresse.
- **Les réponses de la FAQ ont été coupées de moitié.**

La page fait 346 mots et 3 393 px de haut sur téléphone, contre le double
auparavant. Toute rédaction future doit tenir cette contrainte : si une section
demande un paragraphe, c'est qu'elle demande une liste.

## Le mouvement

Deux gestes, pas davantage :

1. l'entrée du héros au chargement, décalée de 80 ms par bloc ;
2. le glissement de la bannière.

Le carrousel utilise le **défilement natif avec accrochage** (`scroll-snap`), pas
une piste translatée en CSS : c'est ce qui le rend **glissable au doigt** sur
téléphone, au trackpad sur portable, et aux flèches au clavier. Les points ne
pilotent rien, ils reflètent la position réelle — il n'y a donc rien à
désynchroniser.

Le défilement automatique (6,5 s) s'arrête dès le premier geste, et se suspend
au survol, au focus clavier et quand l'onglet passe en arrière-plan. Sur
téléphone les flèches disparaissent : le doigt suffit.

Les maquettes se dimensionnent sur la largeur de leur cadre
(`container-type: inline-size` + unités `cqw`). Sans cela, un titre de 18 px
occupe 2,5 % d'un cadre de bureau mais 5 % d'un cadre de téléphone, et la
miniature paraît zoomée. Des `max()` posent un plancher de lisibilité.

Aucune révélation au défilement : le contenu est lisible en permanence. Courbes
et durées suivent la doctrine d'Emil Kowalski — `ease-out` marqué pour les
entrées, `cubic-bezier(.32,.72,0,1)` pour le glissement, propriétés nommées
plutôt que `all`.

## Polices

Archivo (licence SIL Open Font), une seule famille, axe de graisse 300–620,
sous-ensemblée sur un jeu de caractères français complet et intégrée en base64 —
33 Ko. `fonts/sub3.py` régénère le sous-ensemble.
