/**
 * CHAPITRE 4 — CENTRALE ÉNERGÉTIQUE (niveaux 19 à 24)
 *
 * Rôle : donner un pouvoir, puis le rendre indispensable.
 * Le module d'impulsion (dash) arrive ici, au niveau 19 — soit aux deux tiers
 * du jeu. C'est volontaire : offrir une capacité trop tôt dilue le vocabulaire
 * de base ; l'offrir ici relance complètement la lecture des distances.
 *
 * Rappel de portée, module actif : ~165 px ≈ 6,8 tuiles.
 * Tout gouffre de 6 tuiles est donc franchissable, et uniquement en dashant.
 */

export const CHAPTER_4 = {
  id: 4,
  name: 'Centrale énergétique',
  subtitle: 'Secteur D — cœur thermique',
  palette: 'core',
  dash: true,
  levels: [
    // ---------------------------------------------------------------- 19 ----
    {
      name: 'Impulsion',
      hint: 'Module d\'impulsion installé.',
      // Aucun piège. Un niveau entier pour que la nouvelle capacité devienne un
      // réflexe : deux gouffres impossibles sans elle, rien d'autre.
      w: 36, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 8, h: 3 },
        { t: 'solid', x: 14, y: 12, w: 6, h: 3 },
        { t: 'mirage', x: 20, y: 12, w: 1, h: 3 },
        { t: 'solid', x: 21, y: 12, w: 1, h: 3 },
        { t: 'solid', x: 28, y: 12, w: 8, h: 3 },
        { t: 'deco', x: 3, y: 10.6, w: 1.5, h: 1.4, kind: 'core' },
        { t: 'sign', x: 4.5, y: 9.4, w: 4, h: 1, text: 'IMPULSION\nDISPONIBLE', always: true },
        { t: 'exit', x: 32, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 20 ----
    {
      name: 'Lames',
      hint: 'Elles ne s\'arrêtent jamais.',
      // Leçon : un danger mobile régulier est un métronome, pas une menace.
      w: 36, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 36, h: 3 },
        { t: 'mover', x: 7, y: 11, w: 1, h: 1, path: [[12, 11]], speed: 150, mode: 'pingpong', deadly: true, style: 'saw' },
        { t: 'mover', x: 16, y: 5, w: 1, h: 1, path: [[16, 11]], speed: 175, mode: 'pingpong', deadly: true, style: 'saw' },
        { t: 'mover', x: 21, y: 11, w: 1, h: 1, path: [[27, 11]], speed: 205, mode: 'pingpong', deadly: true, style: 'saw' },
        { t: 'mover', x: 29, y: 5, w: 1, h: 1, path: [[29, 11]], speed: 190, mode: 'pingpong', deadly: true, style: 'saw' },
        { t: 'mirage', x: 13, y: 12, w: 1, h: 3 },
        { t: 'deco', x: 0, y: 11.7, w: 36, h: 0.3, kind: 'rail' },
        { t: 'exit', x: 32, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 21 ----
    {
      name: 'Surchauffe',
      hint: 'Le sol n\'a plus d\'adhérence.',
      // Leçon : deux zones qui modifient la physique elle-même. La glace punit
      // l'excès de vitesse, le survolteur l'exige. Le joueur doit changer de
      // registre en deux secondes.
      w: 38, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 13, h: 3 },
        { t: 'field', x: 6, y: 9, w: 7, h: 3, kind: 'ice' },
        // Un mirage sur la glace : impossible de s'arrêter pile devant.
        { t: 'mirage', x: 10, y: 12, w: 1, h: 3 },
        { t: 'solid', x: 16, y: 12, w: 9, h: 3 },
        // Le survolteur couvre AUSSI le gouffre : une zone de vitesse qui
        // s'arrête au bord de la falaise ne sert à rien, la décélération
        // aérienne annule l'élan en un dixième de seconde.
        { t: 'field', x: 16, y: 5, w: 15, h: 7, kind: 'fast', amount: 1.75 },
        { t: 'solid', x: 31, y: 12, w: 7, h: 3 },
        { t: 'zap', x: 13, y: 13.4, w: 3, h: 0.6, dir: 'up' },
        { t: 'zap', x: 25, y: 13.4, w: 6, h: 0.6, dir: 'up' },
        { t: 'sign', x: 3, y: 9.4, w: 3.5, h: 1, text: 'GIVRE →', always: true },
        { t: 'sign', x: 17, y: 8, w: 4, h: 1, text: 'SURVOLTEUR', always: true },
        { t: 'exit', x: 34, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 22 ----
    {
      name: 'Séquence',
      hint: 'Trois verrous.',
      // Leçon : le troisième verrou est DERRIÈRE toi.
      //
      // Le robot apparaît face à la droite ; le verrou 1 est hors champ, à
      // gauche. On presse donc 2 et 3, on arrive devant une porte close, et
      // il faut faire l'aller-retour — sous les 3,8 s de la temporisation.
      // Aller-retour mesuré : 2,5 s, soit 1,3 s de marge. Aucun blocage
      // possible, le verrou se réarme.
      //
      // Première version : les trois verrous alignés de gauche à droite, le
      // temporisé en premier. Le joueur prudent automatique (tools/playtest.js)
      // l'a franchie du premier coup en courant tout droit — la temporisation
      // n'expirait jamais avant la porte. Il n'y avait tout simplement pas
      // d'énigme.
      w: 30, h: 15,
      spawn: { x: 8, y: 10 },
      solve: [[14, 11.5], [19, 11.5], [3, 11.5], [23, 11.5], [27, 11.4]],
      entities: [
        { t: 'solid', x: 0, y: 12, w: 30, h: 3 },
        { t: 'button', x: 2, y: 11.6, w: 2, h: 0.4, mode: 'once', emits: 'v1', autoOff: 3.8 },
        { t: 'sign', x: 1.2, y: 9.4, w: 3.5, h: 1, text: 'VERROU 1\n3,8 s', always: true },
        { t: 'button', x: 13, y: 11.6, w: 2, h: 0.4, mode: 'once', emits: 'v2' },
        { t: 'sign', x: 12, y: 9.4, w: 3, h: 1, text: 'VERROU 2', always: true },
        { t: 'mirage', x: 16, y: 12, w: 1, h: 3 },
        { t: 'button', x: 18, y: 11.6, w: 2, h: 0.4, mode: 'once', emits: 'v3' },
        { t: 'sign', x: 17, y: 9.4, w: 3, h: 1, text: 'VERROU 3', always: true },

        { t: 'logic', op: 'and', in: ['v1', 'v2', 'v3'], emits: 'porte' },
        { t: 'solid', x: 22, y: 3, w: 2, h: 7 },
        { t: 'door', x: 22, y: 10, w: 2, h: 2, slide: 'up', travel: 2, speed: 420, signal: 'porte' },
        { t: 'deco', x: 21, y: 8.2, w: 0.8, h: 0.8, kind: 'lamp', signal: 'porte' },
        { t: 'exit', x: 26.5, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 23 ----
    {
      name: 'Vague',
      hint: 'Ils tirent l\'un après l\'autre.',
      // Leçon : lire un motif plutôt que quatre pièges. Les quatre faisceaux
      // sont déphasés d'un quart de période : ils forment une vague qui se
      // remonte à contre-temps. Une fois vu, le niveau se traverse d'un trait.
      w: 32, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 32, h: 3 },
        { t: 'laser', x: 7.2, y: 3.5, w: 0.5, h: 0.6, dir: 'down', len: 10, cycle: [0.7, 1.7], phase: 0 },
        { t: 'laser', x: 12.2, y: 3.5, w: 0.5, h: 0.6, dir: 'down', len: 10, cycle: [0.7, 1.7], phase: 0.6 },
        { t: 'laser', x: 17.2, y: 3.5, w: 0.5, h: 0.6, dir: 'down', len: 10, cycle: [0.7, 1.7], phase: 1.2 },
        { t: 'laser', x: 22.2, y: 3.5, w: 0.5, h: 0.6, dir: 'down', len: 10, cycle: [0.7, 1.7], phase: 1.8 },
        { t: 'mirage', x: 14, y: 12, w: 1, h: 3 },
        { t: 'mirage', x: 25, y: 12, w: 1, h: 3 },
        { t: 'deco', x: 6.8, y: 2.4, w: 1.4, h: 1.1, kind: 'panel' },
        { t: 'deco', x: 11.8, y: 2.4, w: 1.4, h: 1.1, kind: 'panel' },
        { t: 'deco', x: 16.8, y: 2.4, w: 1.4, h: 1.1, kind: 'panel' },
        { t: 'deco', x: 21.8, y: 2.4, w: 1.4, h: 1.1, kind: 'panel' },
        { t: 'exit', x: 28, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 24 ----
    {
      name: 'Cœur du réacteur',
      hint: 'Ne t\'arrête pas au milieu.',
      // Final de chapitre : impulsion obligatoire, lame mobile, dalles fatiguées
      // et faisceau final. Long, mais sans temps mort — un final doit être une
      // démonstration de maîtrise, pas une épreuve d'endurance.
      w: 46, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 7, h: 3 },
        { t: 'solid', x: 13, y: 12, w: 5, h: 3 },
        { t: 'mover', x: 15, y: 11, w: 1, h: 1, path: [[15, 6]], speed: 190, mode: 'pingpong', deadly: true, style: 'saw' },

        { t: 'solid', x: 22, y: 11, w: 4, h: 4 },
        { t: 'crumble', x: 26, y: 11, w: 1, h: 1, delay: 0.4 },
        { t: 'crumble', x: 27, y: 11, w: 1, h: 1, delay: 0.4 },
        { t: 'crumble', x: 28, y: 11, w: 1, h: 1, delay: 0.4 },
        { t: 'solid', x: 29, y: 11, w: 2, h: 4 },
        { t: 'mirage', x: 31, y: 11, w: 1, h: 4 },

        { t: 'mover', x: 33, y: 10, w: 3, h: 1, path: [[33, 6], [33, 10]], speed: 74, mode: 'loop', wait: 0.5 },
        { t: 'solid', x: 38, y: 6, w: 8, h: 1 },
        { t: 'laser', x: 44.4, y: 4.9, w: 0.5, h: 0.5, dir: 'left', len: 6, cycle: [0.75, 1.5] },
        { t: 'zap', x: 32, y: 14.4, w: 14, h: 0.6, dir: 'up' },
        { t: 'exit', x: 42, y: 4 },

        // Alcôve secrète : sous la passerelle du réacteur.
        { t: 'zone', x: 19.5, y: 13, w: 1.6, h: 1, mode: 'once', remember: 'secret', emits: 'secret' },
        { t: 'solid', x: 18, y: 14, w: 4, h: 1 },
        { t: 'deco', x: 19.5, y: 13, w: 1.6, h: 1, kind: 'gift' },
      ],
    },
  ],
};
