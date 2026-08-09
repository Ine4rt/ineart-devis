# VOID RUNNER

Un petit robot d'exploration, une station spatiale abandonnée dont tous les
systèmes de sécurité sont détraqués, et trente salles qui mentent.

Jeu de plateforme 2D mobile, orientation paysage. Contrôles à trois boutons,
salles de 10 à 45 secondes, morts fréquentes et redémarrage en 320 ms.

---

## Démarrer

```bash
cd void-runner
npm start            # http://localhost:8080
npm test             # valide les 30 niveaux (le solveur les joue vraiment)
npm run test:fast    # structure et cohérence seulement (~1 s)
npm run build        # dist/void-runner.html — le jeu entier en un fichier
```

`dist/void-runner.html` (212 Ko) s'ouvre directement dans n'importe quel
navigateur, sans serveur : c'est la version à envoyer pour faire essayer le
jeu. La version de développement reste en modules séparés, c'est celle qu'on
modifie.

Aucune dépendance, aucun asset binaire : le jeu est constitué de modules ES
chargés directement par le navigateur, sans transpilation. Le serveur fourni
n'existe que parce que `file://` interdit les modules ; `npm run build`
n'est nécessaire que pour produire la version en un fichier.

Raccourci de test : `index.html#17` démarre directement la salle 17.

**Clavier** — ← → déplacement · Espace / ↑ saut · Maj / X impulsion · Échap pause.
**Manette** — stick ou croix, A saut, X impulsion.
**Tactile** — deux boutons à gauche, saut (et impulsion) à droite.

## Empaqueter pour iOS / Android

Le projet est une application web autonome ; `capacitor.config.json` est déjà
écrit et pointe sur la racine.

```bash
npm i -D @capacitor/cli @capacitor/core @capacitor/ios @capacitor/android
npx cap add ios && npx cap add android
npx cap sync
npx cap open ios      # ou android
```

En attendant, `manifest.webmanifest` + `sw.js` en font déjà une PWA installable
qui démarre hors ligne.

---

## Ce que contient le prototype

| Demandé | État |
|---|---|
| Menu principal, sélection de niveaux, robots, statistiques, paramètres | ✅ |
| 30 niveaux jouables, 5 chapitres | ✅ (le cahier des charges en demandait 10 au minimum) |
| Déplacement, saut, impulsion (dash) | ✅ |
| Système de mort + redémarrage instantané | ✅ 320 ms, écourtable |
| Types de pièges | ✅ 24 mécaniques (8 demandées) |
| Compteur de morts, écran de victoire | ✅ morts du jour, morts totales, némésis |
| Robots déblocables | ✅ 8, purement cosmétiques |
| Sauvegarde locale | ✅ versionnée et migrable |
| Effets sonores + musique | ✅ 100 % synthétisés, zéro fichier audio |
| Interface mobile, haptique, manette | ✅ |
| Salles secrètes | ✅ 3 |

---

## Architecture

```
src/
  engine/        ← simulation pure : aucune dépendance au DOM, exécutable en Node
    constants.js       toutes les valeurs de game feel, à un seul endroit
    World.js           instance de niveau, signaux, ordre de mise à jour
    Player.js          course, saut variable, coyote time, dash, gravité
    entities/          les 24 mécaniques, une classe chacune
  data/
    levels/            LES NIVEAUX SONT DES DONNÉES (chapitre1.js … chapitre5.js)
    skins.js  quips.js
  core/          ← services : boucle, caméra, audio, entrées, sauvegarde…
  render/        ← dessin : thème, robot vectoriel, rendu de salle
  ui/            ← écrans DOM
tests/
  solver.js  run.js    validation automatique des niveaux
```

La frontière importante est `src/engine/`. Rien n'y touche au navigateur : pas
de `window`, pas de `canvas`, pas d'`Audio`. C'est ce qui permet au solveur de
faire tourner un niveau en Node — et donc de **prouver** qu'il est
franchissable.

### Ajouter un niveau

Ajouter un objet dans un fichier de chapitre. Rien d'autre.

```js
{
  name: 'Sas gelé',
  hint: 'Ça glisse.',
  w: 28, h: 15,
  spawn: { x: 1.5, y: 10 },
  entities: [
    { t: 'solid', x: 0, y: 12, w: 28, h: 3 },
    { t: 'field', x: 8, y: 9, w: 8, h: 3, kind: 'ice' },
    { t: 'laser', x: 18, y: 4, dir: 'down', len: 9, cycle: [0.8, 1.4] },
    { t: 'exit',  x: 25, y: 10 },
  ],
}
```

Coordonnées **en tuiles** (24 px). Les murs de bordure sont ajoutés
automatiquement. Puis `npm test` : le solveur joue le niveau et refuse de le
valider s'il n'est pas franchissable.

### Ajouter une mécanique

Une classe dans `src/engine/entities/`, une ligne dans le registre
`entities/index.js`, un cas dans `render/Renderer.js`. Aucun niveau n'est à
recoder.

---

## Repères de conception (mesurés, pas devinés)

| Grandeur | Valeur |
|---|---|
| Tuile | 24 px |
| Hauteur de saut | **83 px — 3,4 tuiles** (mesurée, pas calculée) |
| Portée d'un saut lancé | **114 px — 4,8 tuiles** |
| Portée avec impulsion | ~160 px — 6,6 tuiles |
| Pas de simulation | 1/60 s fixe, déterministe |
| Coyote time / jump buffer | 90 ms / 120 ms |
| Reprise après la mort | 320 ms |

Ces chiffres sont **mesurés sur le moteur** au lancement de `npm test`, jamais
recopiés d'une formule : la formule continue annonce 87 px d'apex, la
simulation à pas fixe en donne 83. Neuf pour cent d'écart suffisent à rendre un
gouffre infranchissable sans que rien ne se voie sur le plan.

Deux outils complètent le solveur, parce que « franchissable » ne veut pas dire
« jouable » :

```bash
node tools/tolerance.js 1   # de combien peut-on se tromper sur une salle ?
npm test                    # inclut la robustesse : une frame d'erreur est-elle fatale ?
```

`tolerance.js` rejoue la salle pour toutes les combinaisons (instant du saut ×
durée d'appui) et dessine la carte des issues. Repère : une tape de pouce dure
60 à 150 ms — une salle d'introduction dont la marge d'appui est plus courte
que ça est injouable, quoi qu'en dise le solveur. C'est exactement ce qui
clouait la salle 1 : 150 ms de marge.

---

## Monétisation prévue

Rien n'est implémenté (aucun SDK publicitaire dans le dépôt), mais l'ossature
est posée pour ce qui a été décidé :

- **jamais** d'interstitiel pendant une salle ni entre deux tentatives ;
- publicité éventuellement au retour au menu principal, et pas plus d'une
  toutes les quelques minutes ;
- achat unique « supprimer les publicités » ;
- robots premium possibles — à la condition absolue qu'ils restent cosmétiques.

La règle qui commande tout : dans un jeu où l'on meurt mille fois, la moindre
interruption entre deux essais tue le jeu.

---

## Ce qui n'est pas fait

- Pas de sauvegarde distante (iCloud / Google Play Jeux). `SaveManager` isole
  déjà les accès dans `read()` / `write()` pour rendre l'ajout indolore.
- Pas de localisation : l'interface est en français uniquement.
- Pas de classements ni de rejouabilité en ligne.
- Le solveur valide qu'un niveau est franchissable, pas qu'il est *amusant* —
  ça, seul le test auprès de joueurs le dira.
