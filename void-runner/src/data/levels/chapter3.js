/**
 * CHAPITRE 3 — ZONE DE GRAVITÉ (niveaux 13 à 18)
 *
 * Rôle : casser le repère le plus solide du joueur — le bas.
 * Une fois la gravité inversée, tout le vocabulaire appris redevient neuf :
 * un plafond est un sol, une chute est une ascension, un saut descend.
 *
 * Le niveau 15 est le pivot du jeu : il rejoue le niveau 4, à l'identique, en
 * ayant changé les règles. C'est le seul moment où j'utilise la mémoire du
 * joueur comme une arme — au-delà, ce serait de la mauvaise foi.
 */

export const CHAPTER_3 = {
  id: 3,
  name: 'Zone de gravité',
  subtitle: 'Secteur C — champs artificiels',
  palette: 'grav',
  levels: [
    // ---------------------------------------------------------------- 13 ----
    {
      name: 'Haut / bas',
      hint: 'Le plafond est un sol.',
      w: 28, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 9, h: 3 },
        { t: 'solid', x: 20, y: 12, w: 8, h: 3 },
        { t: 'solid', x: 0, y: 0, w: 28, h: 1 },

        { t: 'grav', x: 8.4, y: 8, w: 1.2, h: 4, dir: 'flip', mode: 'flip' },
        { t: 'grav', x: 19.4, y: 1, w: 1.2, h: 4, dir: 'flip', mode: 'flip' },

        // Obstacle « au plafond » : une fois inversé, il faut sauter par-dessus,
        // c'est-à-dire vers le bas de l'écran. Tout est à réapprendre.
        { t: 'zap', x: 13, y: 1, w: 2, h: 0.5, dir: 'down' },
        { t: 'deco', x: 10, y: 1.1, w: 2, h: 0.4, kind: 'pipe' },
        { t: 'sign', x: 4.5, y: 9.4, w: 4, h: 1, text: 'CHAMP INVERSEUR\n↑↓', always: true },
        { t: 'exit', x: 24, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 14 ----
    {
      name: 'Deux sens',
      hint: 'Choisis ton sol.',
      // Leçon : la gravité devient une ressource. Deux couloirs superposés,
      // chacun impraticable seul ; il faut basculer deux fois.
      w: 32, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 32, h: 3 },
        { t: 'solid', x: 0, y: 0, w: 32, h: 1 },
        { t: 'solid', x: 12, y: 5, w: 10, h: 1 },

        { t: 'grav', x: 6.4, y: 8, w: 1.2, h: 4, dir: 'flip', mode: 'flip' },
        { t: 'laser', x: 0.4, y: 3.2, w: 0.5, h: 0.5, dir: 'right', len: 11, cycle: [1.0, 1.4] },
        { t: 'zap', x: 12, y: 4.5, w: 10, h: 0.5, dir: 'down' },
        { t: 'grav', x: 23.4, y: 1, w: 1.2, h: 3.5, dir: 'flip', mode: 'flip' },
        { t: 'crumble', x: 26, y: 12, w: 1, h: 1, delay: 0.4 },
        { t: 'crumble', x: 27, y: 12, w: 1, h: 1, delay: 0.4 },
        { t: 'exit', x: 29, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 15 ----
    {
      name: 'Déjà-vu',
      hint: '',
      // COPIE EXACTE de la géométrie du niveau 4. Ce qui change :
      //  - les deux faisceaux sont déphasés dans l'autre sens ;
      //  - l'îlot central s'effondre (il était solide au niveau 4).
      // Le joueur exécute sa solution mémorisée et meurt. La règle n'a pas
      // triché : c'est lui qui n'a pas regardé.
      w: 26, h: 15,
      spawn: { x: 2, y: 10 },
      dejaVu: 4,
      entities: [
        { t: 'solid', x: 0, y: 12, w: 9, h: 3 },
        { t: 'crumble', x: 12, y: 12, w: 1, h: 1, delay: 0.45, respawn: 2.6 },
        { t: 'crumble', x: 13, y: 12, w: 1, h: 1, delay: 0.45, respawn: 2.6 },
        { t: 'crumble', x: 14, y: 12, w: 1, h: 1, delay: 0.45, respawn: 2.6 },
        { t: 'crumble', x: 15, y: 12, w: 1, h: 1, delay: 0.45, respawn: 2.6 },
        { t: 'solid', x: 19, y: 12, w: 7, h: 3 },
        // Mêmes faisceaux qu'à la salle 4, phases interverties. Rien d'autre
        // ne change dans le rythme : c'est la mémoire du joueur qui le trompe.
        { t: 'laser', x: 10.2, y: 4.5, w: 0.5, h: 0.6, dir: 'down', len: 9, cycle: [0.9, 1.8], phase: 1.3 },
        { t: 'laser', x: 17.2, y: 4.5, w: 0.5, h: 0.6, dir: 'down', len: 9, cycle: [0.9, 1.8], phase: 2.2 },
        { t: 'deco', x: 9.8, y: 3.4, w: 1.4, h: 1.1, kind: 'panel' },
        { t: 'deco', x: 16.8, y: 3.4, w: 1.4, h: 1.1, kind: 'panel' },
        { t: 'sign', x: 4, y: 9.4, w: 4, h: 1, text: 'FAISCEAUX\nINTERMITTENTS', always: true },
        { t: 'exit', x: 22.5, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 16 ----
    {
      name: 'Soufflerie',
      hint: 'Laisse-toi porter.',
      // Leçon : une force externe peut être plus forte que le saut. La colonne
      // d'air est le seul moyen de franchir les cloisons.
      w: 32, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 32, h: 3 },
        { t: 'solid', x: 10, y: 4, w: 1, h: 8 },
        { t: 'solid', x: 21, y: 4, w: 1, h: 8 },

        { t: 'fan', x: 7.5, y: 4, w: 2.4, h: 8, dir: 'up', force: 2650 },
        { t: 'deco', x: 7.5, y: 11.6, w: 2.4, h: 0.5, kind: 'grid' },

        // Deuxième soufflerie coupée : il faut d'abord la rallumer.
        { t: 'fan', x: 18.5, y: 4, w: 2.4, h: 8, dir: 'up', force: 2650, signal: 'air' },
        { t: 'deco', x: 18.5, y: 11.6, w: 2.4, h: 0.5, kind: 'grid' },
        { t: 'button', x: 14, y: 11.6, w: 2, h: 0.4, mode: 'once', emits: 'air' },
        { t: 'sign', x: 12, y: 9.4, w: 3.5, h: 1, text: 'VENTILATION\nSECTEUR 2', always: true },

        // Le courant d'air du haut pousse vers la gauche : rester en altitude
        // n'est pas gratuit.
        { t: 'fan', x: 11, y: 1, w: 10, h: 3, dir: 'left', force: 950 },
        { t: 'exit', x: 28, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 17 ----
    {
      name: 'Champ magnétique',
      hint: 'Quelque chose t\'attire.',
      // Leçon : un champ courbe les trajectoires. Le premier aide, le second
      // tue — et les deux sont dessinés exactement de la même façon, à la
      // couleur près (cyan = attire vers un point sûr, orange = vers la grille).
      w: 32, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 8, h: 3 },
        { t: 'solid', x: 15, y: 12, w: 5, h: 3 },
        { t: 'solid', x: 26, y: 12, w: 6, h: 3 },

        { t: 'field', x: 8, y: 5, w: 7, h: 7, kind: 'magnet', anchor: [16, 11], force: 620 },
        { t: 'field', x: 20, y: 5, w: 6, h: 7, kind: 'magnet', anchor: [23, 13.6], force: 700, hostile: true },
        { t: 'zap', x: 20, y: 13.4, w: 6, h: 0.6, dir: 'up' },

        { t: 'mover', x: 20.5, y: 11, w: 2, h: 0.6, path: [[24, 11]], speed: 78, mode: 'pingpong' },
        { t: 'sign', x: 4, y: 9.4, w: 3.5, h: 1, text: 'CHAMPS ACTIFS', always: true },
        { t: 'exit', x: 29, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 18 ----
    {
      name: 'Faux repère',
      hint: 'Un point de contrôle. Enfin.',
      // Leçon : même l'interface peut mentir. Le premier point de contrôle est
      // vrai, le second est un décor. Un seul faux checkpoint dans tout le jeu :
      // répété, ce ne serait plus une farce, ce serait un impôt.
      w: 42, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 12, h: 3 },
        { t: 'crusher', x: 5, y: 0, w: 3, h: 3, axis: 'y', to: 9, speed: 620, back: 110, hold: 0.3, trigger: 'near', range: 3 },
        { t: 'check', x: 10, y: 10.5, w: 1, h: 1.5 },

        { t: 'vanish', x: 13.5, y: 11, w: 2, h: 0.6, mode: 'onLand', delay: 0.45, respawn: 1.4 },
        { t: 'vanish', x: 17, y: 10, w: 2, h: 0.6, mode: 'onLand', delay: 0.45, respawn: 1.4 },
        { t: 'solid', x: 20, y: 12, w: 8, h: 3 },
        { t: 'check', x: 24, y: 10.5, w: 1, h: 1.5, fake: true },

        { t: 'grav', x: 28.4, y: 8, w: 1.2, h: 4, dir: 'flip', mode: 'flip' },
        { t: 'solid', x: 0, y: 0, w: 42, h: 1 },
        { t: 'zap', x: 32, y: 1, w: 2, h: 0.5, dir: 'down' },
        { t: 'zap', x: 36, y: 1, w: 2, h: 0.5, dir: 'down' },
        { t: 'grav', x: 39.4, y: 1, w: 1.2, h: 4, dir: 'flip', mode: 'flip' },
        { t: 'solid', x: 36, y: 12, w: 6, h: 3 },
        { t: 'exit', x: 39, y: 10 },
      ],
    },
  ],
};
