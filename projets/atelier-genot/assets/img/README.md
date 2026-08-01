# Photos

Déposez les images ici, puis remplacez l'emplacement correspondant dans
`index.html`. Un emplacement ressemble à ceci :

```html
<div class="slot"><span class="slot__label">Réalisation 2</span></div>
```

Il devient :

```html
<div class="slot"><img src="assets/img/realisation-2.jpg" alt="Portail coulissant en acier galvanisé" loading="lazy"></div>
```

Le style d'emplacement (bordure en pointillés, étiquette) disparaît
automatiquement dès qu'une `<img>` est présente — c'est géré en CSS par
`.slot:has(img)`, rien d'autre à faire.

## Ce qu'il faut

| Fichier | Où | Format | Largeur mini |
|---|---|---|---|
| `accueil.jpg` | Bandeau d'accueil | Paysage | 2000 px |
| `realisation-1.jpg` | Grande tuile de la mosaïque | Paysage | 1600 px |
| `realisation-2…6.jpg` | Petites tuiles | Paysage ou carré | 1000 px |
| `atelier.jpg` | Section « L'atelier » | **Portrait** | 1000 px |
| `logo.svg` | Barre de navigation | SVG de préférence | — |

## Conseils de prise de vue

Un site de ferronnier vit de ses photos. Trois règles suffisent :

1. **Les ouvrages en situation, pas sur l'établi.** Un garde-corps posé sur sa
   terrasse vend ; le même sur un tréteau ne dit rien.
2. **Lumière du jour, ciel couvert.** Le plein soleil écrase les reliefs du
   métal et crée des ombres dures.
3. **Une photo de l'artisan au travail.** C'est celle qui crée la confiance,
   et c'est presque toujours celle qui manque.

## Poids des fichiers

Visez **moins de 300 Ko par image**. Un JPEG de qualité 80 à 2000 px suffit
largement. Sans compression, une page avec huit photos d'appareil photo pèse
40 Mo et met dix secondes à s'afficher sur un téléphone en 4G — ce qui annule
tout l'intérêt du site.

Le format WebP réduit encore de moitié à qualité égale, et il est reconnu par
tous les navigateurs actuels.

## Texte alternatif

Remplissez toujours l'attribut `alt` avec une description réelle de l'ouvrage
(« Escalier quart tournant, structure acier et marches en chêne »). C'est ce que
lisent les moteurs de recherche et les lecteurs d'écran — et ce qui s'affiche si
l'image ne charge pas.
