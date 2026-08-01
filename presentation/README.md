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

## Le monde visuel

Signalétique de chantier. Le gris très foncé et l'orange haute visibilité que
vous avez imposés sont natifs de ce monde — ce n'est pas un thème sombre
décoratif. Vos prospects (menuisier, garagiste, artisan) croisent cette
signalétique tous les jours, et « chantier » dit exactement ce que vous vendez :
quelque chose qu'on vous livre fini.

Conséquences tenues dans le code :

- L'orange est posé en **aplats pleins** — le bandeau des prises en charge, le
  bloc de contact — jamais en halo ni en dégradé. Un orange qui ne sert que de
  liseré lumineux, c'est le fond sombre générique des logiciels en ligne.
- Lettrage **Archivo étendu** (axe de largeur poussé à 118) pour les titres :
  c'est le lettrage large des plaques, pas une police d'interface.
- Bords francs, filets de 2 px, aucun verre, aucune ombre colorée.
- **Un seul monde, sombre.** Une identité de marque ne se décline pas en version
  claire. La console, elle, garde ses deux thèmes : on l'ouvre huit heures par
  jour.

## Le mouvement

Un seul moment orchestré au chargement (le titre monte, le bandeau orange se
déploie latéralement), plus une entrée séquentielle sur les quatre étapes, où le
glissement porte l'ordre de lecture. Rien d'autre ne bouge au défilement : une
entrée identique sur chaque section est du bruit.

Courbes et durées suivent la doctrine d'Emil Kowalski : `ease-out` marqué
(`cubic-bezier(.23,1,.32,1)`) pour les entrées, propriétés nommées plutôt que
`all`, retour d'appui sur les boutons.

Le contenu est **visible par défaut** : l'état masqué n'est posé que par le
script, uniquement sur ce qui est sous le pli, et un filet de sécurité de quatre
secondes révèle tout quoi qu'il arrive. Sans JavaScript, ou si l'utilisateur a
demandé moins d'animations, la page reste entièrement lisible.

## Polices

Archivo (licence SIL Open Font) en deux largeurs, sous-ensemblée sur un jeu de
caractères français complet et intégrée en base64 — 58 Ko au total. C'est ce qui
garantit un rendu identique partout, y compris derrière un réseau qui bloquerait
un CDN. `fonts/sub2.py` régénère les sous-ensembles.
