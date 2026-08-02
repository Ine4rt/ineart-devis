# Photos

## Ce qui est en place

| Fichier | Où | Provenance |
|---|---|---|
| `logo.png` | Barre de navigation, écran de fin | Photo de profil Facebook, détourée |
| `logo-sombre.png` | Réserve, pour fond clair | Idem, en anthracite |
| `icone-32.png` / `icone-180.png` | Onglet du navigateur, écran d'accueil iOS | Logo sur fond anthracite |
| `accueil.jpg` | Bandeau d'accueil | Plaque corten, cadrage paysage |
| `realisation-1.jpg` | Grande tuile | Portail à découpes florales, vue d'ensemble |
| `realisation-2.jpg` | Tuile verticale | Plaque corten, cadrage portrait |
| `realisation-3.jpg` | Petite tuile | Lettrage « Wathelet Lamproye » |
| `realisation-4.jpg` | Petite tuile | Motif de chardons du portail |
| `atelier.jpg` | Section « L'atelier » | Vantail droit du portail, en hauteur |

**Six emplacements, deux chantiers.** Toutes ces images viennent de deux
photos : le portail et la plaque. Les recadrages sont honnêtes — ce sont bien
leurs ouvrages — mais le site respire dès qu'il y a davantage de chantiers
différents.

Le logo vient d'un JPEG. Il est net à la taille où il s'affiche, mais
**demandez le fichier vectoriel** si le graphiste du client l'a encore : un SVG
reste parfait à toutes les tailles et à l'impression.

## Ce qui manquerait le plus

Trois photos changeraient l'effet du site :

1. **Un escalier** — c'est la première prestation listée, et rien ne l'illustre.
2. **Un garde-corps posé** — sur sa terrasse ou son balcon, en situation.
3. **L'artisan au travail** — c'est la photo qui crée la confiance, et c'est
   presque toujours celle qui manque.

## Remplacer ou ajouter une image

Un emplacement vide ressemble à ceci :

```html
<div class="slot"><span class="slot__label">Réalisation 5</span></div>
```

Il devient :

```html
<div class="slot"><img src="assets/img/realisation-5.jpg" alt="Garde-corps de terrasse en acier thermolaqué, remplissage verre" loading="lazy"></div>
```

Le style d'emplacement — bordure en pointillés, étiquette — disparaît
automatiquement dès qu'une `<img>` est présente. C'est géré en CSS par
`.slot:has(img)`, rien d'autre à faire.

Pour ajouter une tuile à la mosaïque, copiez un bloc `<figure class="work">`
existant. Les classes `work--large` et `work--tall` donnent les grandes tailles ;
sans classe, la tuile est un carré simple. Voir le README principal pour la
logique de composition.

## Formats et poids

| Emplacement | Format | Largeur utile |
|---|---|---|
| Accueil | Paysage | 1500 px |
| Grande tuile | Carré ou paysage | 1100 px |
| Tuile verticale | **Portrait** | 700 px |
| Petites tuiles | Carré | 700 px |
| Section atelier | **Portrait** | 620 px |

Visez **moins de 250 Ko par image**. Un JPEG de qualité 75 suffit largement à
ces dimensions. Sans compression, une page avec huit photos sorties d'un
appareil pèse 40 Mo et met dix secondes à s'afficher sur un téléphone en 4G —
ce qui annule tout l'intérêt du site.

Le WebP réduit encore de moitié à qualité égale et il est reconnu par tous les
navigateurs actuels.

## Conseils de prise de vue

Un site de ferronnier vit de ses photos. Trois règles suffisent :

1. **Les ouvrages en situation, pas sur l'établi.** Un garde-corps posé sur sa
   terrasse vend ; le même sur un tréteau ne dit rien. Les deux photos
   existantes le font déjà très bien — le portail devant sa maison, la plaque
   sur son mur en pierre.
2. **Lumière du jour, ciel couvert.** Le plein soleil écrase les reliefs du
   métal et creuse des ombres dures.
3. **Un détail pour chaque ensemble.** Une vue large montre l'ouvrage ; un gros
   plan sur une soudure, une découpe ou un assemblage montre le savoir-faire.
   C'est ce second cliché qui justifie le prix.

## Texte alternatif

Remplissez toujours l'attribut `alt` avec une description réelle de l'ouvrage
(« Escalier quart tournant, structure acier et marches en chêne »). C'est ce que
lisent les moteurs de recherche et les lecteurs d'écran — et ce qui s'affiche si
l'image ne charge pas.
