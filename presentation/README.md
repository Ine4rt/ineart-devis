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

**Base claire, bandes alternées.** C'est le point le plus important du fichier :
un fond continu, clair ou sombre, transforme la page en un seul bloc. Le rythme
vient de l'alternance, pas de la quantité de marge.

L'ordre des bandes, à respecter si vous ajoutez une section :

| Bande | Fond | Rôle |
|---|---|---|
| Accueil | `#f5f5f7` | Le message et les maquettes |
| Prestations | blanc | Le contenu |
| Méthode | `#16181c` | **La ponctuation sombre — une seule dans la page** |
| Questions | blanc | Le contenu |
| Contact | `#f5f5f7` | La sortie |

Une deuxième bande noire annulerait l'effet. Si une section s'ajoute, elle est
blanche ou grise.

L'orange (`#d2570b`) ne sert qu'à l'action : le bouton, les liens, les puces,
les numéros d'étape. Jamais un aplat.

## La typographie

Registre Apple : **l'interlettrage se resserre à mesure que le corps grandit**
(−0,035 em sur le titre, −0,005 em sur le texte courant), et l'interligne
descend avec lui (1,04 sur le titre, 1,55 sur le texte). Une valeur unique pour
toutes les tailles serait fausse quelque part.

Les retours du titre sont **imposés à la main** aux frontières de phrase. Laissé
libre, le navigateur coupait « Sur / mesure » au milieu.

## Le mouvement

Deux gestes : l'entrée à l'ouverture, et le carrousel.

Le carrousel utilise le **défilement natif avec accrochage** (`scroll-snap`).
C'est ce qui donne gratuitement les trois qualités qu'une piste translatée en
CSS ne sait pas produire : suivi au doigt à l'unité près, inertie au
relâchement, et interruption à n'importe quel instant. Les points ne pilotent
rien — ils reflètent la position réelle, donc rien ne peut se désynchroniser.

L'avance automatique (6,5 s) cesse dès le premier contact, et se suspend au
survol, au focus clavier et quand l'onglet passe en arrière-plan.

Le bouton répond **à l'appui** (`:active { scale(.97) }`, 100 ms), pas au
relâchement. Cibles tactiles à 44 px.

Les maquettes se dimensionnent sur la largeur de leur cadre
(`container-type: inline-size` + unités `cqw`). À taille fixe, un titre pèse
deux fois plus lourd dans un cadre de téléphone que de bureau, et la miniature
paraît zoomée.

## Polices

Archivo (licence SIL Open Font), une seule famille, axe de graisse 300–620,
sous-ensemblée sur un jeu de caractères français complet et intégrée en base64 —
33 Ko. `fonts/sub3.py` régénère le sous-ensemble.
