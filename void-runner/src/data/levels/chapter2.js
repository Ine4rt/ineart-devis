/**
 * CHAPITRE 2 — LABORATOIRE (niveaux 7 à 12)
 *
 * Rôle : apprendre à se méfier de l'interface elle-même.
 * Le chapitre 1 mentait avec le décor. Celui-ci ment avec les SYMBOLES :
 * une porte de sortie, un panneau directionnel, un téléporteur, un voyant.
 * Règle que je m'impose ici : chaque mensonge est démasquable en une seconde
 * une fois qu'on sait où regarder (voyant, étiquette, halo instable).
 */

export const CHAPTER_2 = {
  id: 2,
  name: 'Laboratoire',
  subtitle: 'Secteur B — recherche appliquée',
  palette: 'lab',
  levels: [
    // ---------------------------------------------------------------- 07 ----
    {
      name: 'Sortie évidente',
      hint: 'Trop facile ?',
      // Leçon : une sortie peut être un mensonge complet.
      // La fausse sortie renvoie au départ — et, ce faisant, révèle les
      // plateformes fantômes du vrai chemin. La punition EST l'indice.
      w: 30, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 30, h: 3 },
        { t: 'sign', x: 18, y: 9.4, w: 4, h: 1, text: 'SORTIE  →', always: true },
        { t: 'exit', x: 25, y: 10, fake: 'warp', to: [1.5, 10], emits: 'vu' },

        { t: 'ghost', x: 13, y: 10, w: 2, h: 0.6, reveal: 'signal', signal: 'vu' },
        { t: 'ghost', x: 16.5, y: 8, w: 2, h: 0.6, reveal: 'signal', signal: 'vu' },
        { t: 'ghost', x: 20, y: 6, w: 2, h: 0.6, reveal: 'signal', signal: 'vu' },
        { t: 'solid', x: 23, y: 4, w: 7, h: 1 },
        { t: 'exit', x: 26, y: 2 },
        { t: 'deco', x: 23.5, y: 2.6, w: 1.2, h: 1.2, kind: 'vent' },
        { t: 'deco', x: 6, y: 10.6, w: 1.5, h: 1.4, kind: 'crate' },
      ],
    },

    // ---------------------------------------------------------------- 08 ----
    {
      name: 'Deux portes',
      hint: 'Le scanner sait laquelle.',
      // Leçon : un bouton fait toujours DEUX choses.
      // Il désigne la bonne porte (voyant vert) et lance la fermeture du
      // couloir. Information contre temps : c'est ça, le marché.
      w: 30, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 30, h: 3 },
        { t: 'button', x: 4, y: 11.6, w: 2, h: 0.4, mode: 'once', emits: 'scan' },
        { t: 'sign', x: 2, y: 9.4, w: 3.5, h: 1, text: 'SCANNER\nDE PORTES', always: true },

        { t: 'crusher', x: 11, y: 0, w: 4, h: 2, axis: 'y', to: 10, speed: 95, back: 50, hold: 1.4, trigger: 'signal', signal: 'scan' },
        { t: 'deco', x: 11, y: 2.1, w: 4, h: 0.4, kind: 'stripe' },

        { t: 'exit', x: 20, y: 10, fake: 'trap' },
        { t: 'deco', x: 20.2, y: 8.6, w: 0.8, h: 0.8, kind: 'lamp' },
        { t: 'exit', x: 25, y: 10 },
        { t: 'deco', x: 25.2, y: 8.6, w: 0.8, h: 0.8, kind: 'lamp', signal: 'scan' },
      ],
    },

    // ---------------------------------------------------------------- 09 ----
    {
      name: 'Téléportation',
      hint: 'Lis les étiquettes.',
      // Leçon : l'information est écrite, encore faut-il la lire.
      // TP-1 est marqué « DÉFAUT » et renvoie au départ. TP-2 fonctionne.
      // Aucune mort ici : juste une perte de temps. Un chapitre doit respirer.
      w: 32, h: 15,
      spawn: { x: 1.5, y: 10 },
      solve: [[10.5, 8.5], [21, 11.4], [24, 9.4], [27.5, 8.4], [30.5, 7.4]],
      entities: [
        { t: 'solid', x: 0, y: 12, w: 14, h: 3 },
        { t: 'solid', x: 18, y: 12, w: 14, h: 3 },
        // Mur de confinement : infranchissable, il n'y a pas de « bon » saut.
        { t: 'laser', x: 15.6, y: 0.2, w: 0.5, h: 0.5, dir: 'down', len: 13 },
        { t: 'deco', x: 15.2, y: -0.6, w: 1.4, h: 1, kind: 'panel' },

        // TP-2 est en hauteur : il faut le vouloir. TP-1 est sur le chemin
        // naturel, contre le mur du fond — on lui rentre dedans, et il renvoie
        // au départ. Aucune mort, juste douze secondes perdues et une leçon :
        // dans cette station, les étiquettes sont fiables. Pas les évidences.
        { t: 'solid', x: 9, y: 9, w: 3, h: 0.6 },
        { t: 'tele', x: 9.5, y: 8.4, w: 2, h: 0.6, to: [20, 10.6] },
        { t: 'sign', x: 8.6, y: 7.2, w: 3.5, h: 1, text: 'TP-2\nOK', always: true },
        { t: 'tele', x: 13.4, y: 9.5, w: 1, h: 2.5, to: [1.5, 10], broken: true },
        { t: 'sign', x: 12, y: 8, w: 3.5, h: 1, text: 'TP-1\nDÉFAUT', always: true },

        { t: 'vanish', x: 23, y: 10, w: 2, h: 0.6, mode: 'cycle', period: 2.4, duty: 0.7 },
        { t: 'vanish', x: 26.5, y: 9, w: 2, h: 0.6, mode: 'cycle', period: 2.4, duty: 0.7, phase: 0.6 },
        { t: 'solid', x: 29, y: 8, w: 3, h: 1 },
        { t: 'exit', x: 29.6, y: 6 },
      ],
    },

    // ---------------------------------------------------------------- 10 ----
    {
      name: 'Presses',
      hint: 'Avance par à-coups.',
      // Leçon : le rythme de progression. Trois presses déclenchées par la
      // proximité, trois abris. Elles frappent vite mais remontent lentement :
      // la bonne stratégie est d'attendre, pas de courir.
      w: 34, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 34, h: 3 },
        { t: 'crusher', x: 7, y: 0, w: 3, h: 3, axis: 'y', to: 9, speed: 620, back: 110, hold: 0.35, trigger: 'near', range: 3 },
        { t: 'crusher', x: 14, y: 0, w: 3, h: 3, axis: 'y', to: 9, speed: 620, back: 110, hold: 0.35, trigger: 'near', range: 3 },
        { t: 'crusher', x: 21, y: 0, w: 3, h: 3, axis: 'y', to: 9, speed: 620, back: 110, hold: 0.35, trigger: 'near', range: 3 },
        { t: 'deco', x: 7, y: 3.1, w: 3, h: 0.4, kind: 'stripe' },
        { t: 'deco', x: 14, y: 3.1, w: 3, h: 0.4, kind: 'stripe' },
        { t: 'deco', x: 21, y: 3.1, w: 3, h: 0.4, kind: 'stripe' },
        { t: 'sign', x: 2.5, y: 9.4, w: 4, h: 1, text: 'PRESSES\nHYDRAULIQUES', always: true },

        // Alcôve secrète, sous la dernière presse : il faut oser s'arrêter là.
        { t: 'zone', x: 25.4, y: 11, w: 1.6, h: 1, mode: 'once', remember: 'secret', emits: 'secret' },
        { t: 'deco', x: 25.4, y: 11, w: 1.6, h: 1, kind: 'gift' },
        { t: 'exit', x: 30, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 11 ----
    {
      name: 'Vitres',
      hint: 'Compte les temps.',
      // Leçon : la synchronisation. Deux séries de dalles en opposition de
      // phase : il existe un unique tempo qui traverse tout le niveau.
      w: 30, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 7, h: 3 },
        // Onde progressive : chaque dalle est décalée d'un quart de période, ce
        // qui laisse ~1,1 s de recouvrement entre deux dalles voisines. Un
        // déphasage d'une demi-période (mon premier essai) ne laissait que
        // 0,26 s : le solveur l'a immédiatement déclaré infranchissable.
        { t: 'vanish', x: 8, y: 11, w: 2, h: 0.6, mode: 'cycle', period: 2.4, duty: 0.7 },
        { t: 'vanish', x: 12, y: 10, w: 2, h: 0.6, mode: 'cycle', period: 2.4, duty: 0.7, phase: 0.6 },
        { t: 'vanish', x: 16, y: 11, w: 2, h: 0.6, mode: 'cycle', period: 2.4, duty: 0.7, phase: 1.2 },
        { t: 'vanish', x: 20, y: 10, w: 2, h: 0.6, mode: 'cycle', period: 2.4, duty: 0.7, phase: 1.8 },
        { t: 'solid', x: 24, y: 12, w: 6, h: 3 },
        { t: 'deco', x: 24, y: 11.6, w: 6, h: 0.4, kind: 'pipe' },
        { t: 'exit', x: 27, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 12 ----
    {
      name: 'Sortie de secours',
      hint: 'Elle ne veut pas de toi.',
      // Final de chapitre : la sortie FUIT. Elle grimpe vers une passerelle,
      // et le seul moyen de la rattraper est le chemin qu'elle vient de
      // révéler. Une poursuite, pas une punition.
      w: 40, h: 15,
      spawn: { x: 1.5, y: 10 },
      solve: [[16, 11.4], [21.5, 8.5], [25.5, 7.0], [29.5, 7.5], [29.5, 3.5], [33, 4.5]],
      entities: [
        { t: 'solid', x: 0, y: 12, w: 40, h: 3 },
        { t: 'crusher', x: 6, y: 0, w: 3, h: 3, axis: 'y', to: 9, speed: 600, back: 120, hold: 0.3, trigger: 'near', range: 3 },
        { t: 'deco', x: 6, y: 3.1, w: 3, h: 0.4, kind: 'stripe' },

        { t: 'exit', x: 15, y: 10, fake: 'flee', to: [32, 3], range: 4 },

        { t: 'solid', x: 20, y: 9, w: 3, h: 1 },
        { t: 'vanish', x: 24.5, y: 7.5, w: 2, h: 0.6, mode: 'onLand', delay: 0.5, respawn: 1.6 },
        { t: 'mover', x: 28, y: 8, w: 3, h: 1, path: [[28, 4]], speed: 66, mode: 'pingpong', wait: 0.7 },
        { t: 'solid', x: 31, y: 5, w: 9, h: 1 },
        // Le faisceau garde la passerelle finale, jamais la trajectoire de
        // l'ascenseur : on ne tue jamais un joueur qui ne peut plus rien faire.
        { t: 'laser', x: 38.4, y: 3.9, w: 0.5, h: 0.5, dir: 'left', len: 7.5, cycle: [0.8, 1.6] },
        { t: 'sign', x: 17, y: 9.4, w: 3.5, h: 1, text: 'NIVEAU 2\n↑' },
      ],
    },
  ],
};
