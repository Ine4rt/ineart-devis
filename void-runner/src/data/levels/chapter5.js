/**
 * CHAPITRE 5 — SECTEUR INTERDIT (niveaux 25 à 30)
 *
 * Rôle : retourner contre le joueur tout ce qu'il a appris, y compris
 * lui-même. C'est ici qu'arrivent l'écho (son propre trajet devient mortel),
 * la mémoire persistante entre deux tentatives, et le renversement complet du
 * niveau 1 : cette fois, le piège est la sortie.
 */

export const CHAPTER_5 = {
  id: 5,
  name: 'Secteur interdit',
  subtitle: 'Secteur ??? — accès non répertorié',
  palette: 'void',
  dash: true,
  levels: [
    // ---------------------------------------------------------------- 25 ----
    {
      name: 'Écho',
      hint: 'Quelque chose te suit. C\'est toi.',
      // Leçon : l'hésitation devient mortelle. L'écho rejoue le trajet du robot
      // avec 1,6 s de retard. Attendre devant un faisceau, c'est se faire
      // rattraper par la version de soi qui hésitait déjà.
      w: 38, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 13, h: 3 },
        { t: 'mirage', x: 13, y: 12, w: 1, h: 3 },
        { t: 'solid', x: 14, y: 12, w: 7, h: 3 },
        { t: 'mirage', x: 21, y: 12, w: 1, h: 3 },
        { t: 'solid', x: 22, y: 12, w: 16, h: 3 },
        { t: 'clone', x: 1.5, y: 10, delay: 1.6, deadly: true, source: 'live' },
        { t: 'laser', x: 10.2, y: 4, w: 0.5, h: 0.6, dir: 'down', len: 9, cycle: [0.6, 1.3] },
        { t: 'laser', x: 18.2, y: 4, w: 0.5, h: 0.6, dir: 'down', len: 9, cycle: [0.6, 1.3], phase: 0.9 },
        { t: 'laser', x: 26.2, y: 4, w: 0.5, h: 0.6, dir: 'down', len: 9, cycle: [0.6, 1.3], phase: 1.6 },
        { t: 'sign', x: 3.5, y: 9.4, w: 4, h: 1, text: 'NE T\'ARRÊTE PAS', always: true },
        { t: 'exit', x: 34, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 26 ----
    {
      name: 'Code inconnu',
      hint: 'ACCÈS REFUSÉ.',
      // Leçon : la mémoire du jeu dépasse la tentative en cours. Le code trouvé
      // en haut à gauche reste acquis pour toutes les tentatives suivantes —
      // mourir après l'avoir trouvé ne le fait pas perdre.
      w: 34, h: 15,
      spawn: { x: 1.5, y: 10 },
      solve: [[3.5, 8.5], [6.5, 6.9], [3.7, 4.9], [12, 11.4], [23, 11.4], [31, 11.4]],
      entities: [
        { t: 'solid', x: 0, y: 12, w: 34, h: 3 },

        { t: 'solid', x: 2.5, y: 9, w: 2, h: 0.6 },
        { t: 'solid', x: 5.5, y: 7.4, w: 2, h: 0.6 },
        { t: 'solid', x: 2.5, y: 5.4, w: 2.5, h: 0.6 },
        { t: 'zone', x: 3, y: 4.4, w: 1.6, h: 1, mode: 'once', remember: 'code', emits: 'codeVu' },
        { t: 'deco', x: 3, y: 4.3, w: 1.6, h: 1.1, kind: 'core' },

        { t: 'memgate', key: 'code', gte: 1, emits: 'codeMem' },
        { t: 'logic', op: 'or', in: ['codeVu', 'codeMem'], emits: 'porte' },

        // Fausse piste assumée : ce bouton n'ouvre rien du tout.
        { t: 'mirage', x: 10, y: 12, w: 1, h: 3 },
        { t: 'button', x: 13, y: 11.6, w: 2, h: 0.4, mode: 'toggle', emits: 'rien' },
        { t: 'sign', x: 12, y: 9.4, w: 3.5, h: 1, text: 'DIAGNOSTIC', always: true },

        { t: 'solid', x: 20, y: 3, w: 2, h: 7 },
        { t: 'door', x: 20, y: 10, w: 2, h: 2, slide: 'up', travel: 2, speed: 380, signal: 'porte' },
        { t: 'sign', x: 16.5, y: 9.4, w: 3.5, h: 1, text: 'ACCÈS REFUSÉ\nCODE INCONNU', always: true },
        { t: 'deco', x: 22.4, y: 8.2, w: 0.8, h: 0.8, kind: 'lamp', signal: 'porte' },
        { t: 'exit', x: 30, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 27 ----
    {
      name: 'Structure absente',
      hint: 'La poussière ne ment pas.',
      // Leçon : faire confiance à un indice ténu. Les dalles n'apparaissent
      // qu'à 2,5 tuiles — mais la poussière en suspension les trahit de loin.
      // Sans cet indice, ce serait une devinette ; avec, c'est de l'observation.
      w: 36, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 8, h: 3 },
        { t: 'ghost', x: 10, y: 11, w: 2, h: 0.6, reveal: 'near' },
        { t: 'ghost', x: 14, y: 10, w: 2, h: 0.6, reveal: 'near' },
        { t: 'ghost', x: 18, y: 11, w: 2, h: 0.6, reveal: 'near' },
        { t: 'ghost', x: 22, y: 10, w: 2, h: 0.6, reveal: 'near' },
        { t: 'ghost', x: 25.5, y: 11, w: 2, h: 0.6, reveal: 'near' },
        { t: 'solid', x: 28, y: 12, w: 8, h: 3 },
        // Le fond du gouffre paraît vide. Il l'est — jusqu'à ce que le robot
        // s'engage au-dessus : la grille jaillit alors sous ses pieds. Se
        // laisser tomber n'est plus un plan de secours.
        { t: 'zone', x: 9, y: 9, w: 18, h: 4, mode: 'once', emits: 'grille' },
        { t: 'zap', x: 8, y: 14.4, w: 20, h: 0.6, dir: 'up', signal: 'grille' },
        { t: 'deco', x: 10, y: 10.2, w: 2, h: 0.8, kind: 'dust' },
        { t: 'deco', x: 14, y: 9.2, w: 2, h: 0.8, kind: 'dust' },
        { t: 'deco', x: 18, y: 10.2, w: 2, h: 0.8, kind: 'dust' },
        { t: 'deco', x: 22, y: 9.2, w: 2, h: 0.8, kind: 'dust' },
        { t: 'deco', x: 25.5, y: 10.2, w: 2, h: 0.8, kind: 'dust' },
        { t: 'exit', x: 32, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 28 ----
    {
      name: 'Inversion totale',
      hint: 'Tout à l\'envers, et une lame.',
      w: 38, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 38, h: 3 },
        { t: 'solid', x: 0, y: 0, w: 38, h: 1 },
        { t: 'grav', x: 6.4, y: 8, w: 1.2, h: 4, dir: 'flip', mode: 'flip' },

        { t: 'mover', x: 13, y: 1, w: 1, h: 1, path: [[19, 1]], speed: 165, mode: 'pingpong', deadly: true, style: 'saw' },
        // Soufflerie dirigée vers le bas de l'écran : sous gravité inversée,
        // elle arrache le robot du plafond. Il faut la traverser vite.
        { t: 'fan', x: 22, y: 1, w: 3, h: 6, dir: 'down', force: 1500 },
        { t: 'deco', x: 22, y: 0.9, w: 3, h: 0.4, kind: 'grid' },
        { t: 'zap', x: 27, y: 1, w: 2, h: 0.5, dir: 'down' },
        { t: 'mirage', x: 33, y: 12, w: 1, h: 3 },

        { t: 'grav', x: 31.4, y: 1, w: 1.2, h: 4, dir: 'flip', mode: 'flip' },
        { t: 'exit', x: 34.5, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 29 ----
    {
      name: 'Le piège parfait',
      hint: '',
      // Le renversement complet du niveau 1, à la tuile près.
      // Même salle, même trappe, même émetteur mural. Sauf que cette fois :
      //  - la porte du fond est un piège mortel ;
      //  - la vraie sortie est DANS le trou.
      // Le joueur doit accepter de faire exactement ce qui l'a tué au premier
      // niveau du jeu. C'est le point d'orgue de tout le propos.
      w: 26, h: 22,
      spawn: { x: 2, y: 10 },
      dejaVu: 1,
      deathHints: { burn: 'Cette porte-là ne s\'ouvre pas. Regarde en bas.' },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 9, h: 1 },
        { t: 'mirage', x: 9, y: 12, w: 1, h: 1 },
        { t: 'solid', x: 10, y: 12, w: 6, h: 1 },
        { t: 'solid', x: 18, y: 12, w: 8, h: 1 },
        { t: 'door', x: 16, y: 12, w: 1, h: 1, slide: 'left', travel: 1, speed: 260, signal: 'trappe' },
        { t: 'door', x: 17, y: 12, w: 1, h: 1, slide: 'right', travel: 1, speed: 260, signal: 'trappe' },
        { t: 'zone', x: 11, y: 9, w: 3, h: 3, mode: 'once', emits: 'trappe' },

        { t: 'laser', x: 0.3, y: 8.9, w: 0.6, h: 0.5, dir: 'right', len: 26, armSignal: 'altitude' },
        { t: 'zone', x: 14.2, y: 7.1, w: 5, h: 1.6, mode: 'once', emits: 'altitude' },

        { t: 'sign', x: 3.5, y: 9.4, w: 3, h: 1, text: 'SORTIE  →', always: true },
        { t: 'exit', x: 22.5, y: 10, fake: 'trap' },

        // Chambre inférieure : le sol est électrifié, sauf une corniche à
        // droite. Il faut donc tomber ET diriger la chute.
        { t: 'zap', x: 0, y: 20.4, w: 26, h: 0.6, dir: 'up' },
        { t: 'solid', x: 18, y: 19, w: 6, h: 1.4 },
        { t: 'exit', x: 20, y: 17 },
        { t: 'deco', x: 18.5, y: 18.2, w: 1, h: 0.8, kind: 'lamp', signal: 'trappe' },
      ],
    },

    // ---------------------------------------------------------------- 30 ----
    {
      name: 'Sortie',
      hint: 'Dernière salle.',
      // Final : impulsion, lame, dalles fatiguées, soufflerie, faisceau, et une
      // dernière blague — la porte esquisse une fuite d'une tuile, puis
      // renonce. Après trente niveaux, le jeu a le droit de faire une farce
      // qui ne tue personne.
      w: 54, h: 15,
      spawn: { x: 1.5, y: 10 },
      solve: [
        [16, 11.4], [25, 11.4], [29, 7], [34, 4.5], [37, 1.5], [44, 1.5],
        [52, 1.5], [52, 11.4],
      ],
      entities: [
        { t: 'solid', x: 0, y: 12, w: 8, h: 3 },
        { t: 'solid', x: 14, y: 12, w: 6, h: 3 },
        { t: 'mover', x: 17, y: 11, w: 1, h: 1, path: [[17, 6]], speed: 195, mode: 'pingpong', deadly: true, style: 'saw' },

        { t: 'crumble', x: 20, y: 12, w: 1, h: 1, delay: 0.38 },
        { t: 'crumble', x: 21, y: 12, w: 1, h: 1, delay: 0.38 },
        { t: 'crumble', x: 22, y: 12, w: 1, h: 1, delay: 0.38 },
        { t: 'crumble', x: 23, y: 12, w: 1, h: 1, delay: 0.38 },
        { t: 'solid', x: 24, y: 12, w: 3, h: 3 },
        { t: 'zap', x: 27, y: 14.4, w: 12, h: 0.6, dir: 'up' },

        { t: 'fan', x: 28, y: 5, w: 2.6, h: 9, dir: 'up', force: 2700 },
        { t: 'deco', x: 28, y: 13.9, w: 2.6, h: 0.5, kind: 'grid' },
        { t: 'solid', x: 31, y: 5, w: 6, h: 1 },
        { t: 'laser', x: 36.4, y: 3.9, w: 0.5, h: 0.5, dir: 'left', len: 5, cycle: [0.7, 1.5] },

        { t: 'solid', x: 0, y: 0, w: 54, h: 1 },
        { t: 'grav', x: 36.2, y: 1.5, w: 1.2, h: 3.5, dir: 'flip', mode: 'flip' },
        { t: 'zap', x: 41, y: 1, w: 2, h: 0.5, dir: 'down' },
        // Lame plongeante plutôt que lame coulissante : une lame qui balaie un
        // couloir de plafond de trois tuiles ne laisse aucune fenêtre lisible.
        // Celle-ci descend, remonte, et dégage franchement le passage.
        { t: 'mover', x: 46, y: 1, w: 1, h: 1, path: [[46, 4]], speed: 130, mode: 'pingpong', wait: 0.4, deadly: true, style: 'saw' },
        { t: 'grav', x: 50.2, y: 1, w: 1.2, h: 3.5, dir: 'flip', mode: 'flip' },

        { t: 'solid', x: 44, y: 12, w: 3, h: 3 },
        { t: 'mirage', x: 47, y: 12, w: 1, h: 3 },
        { t: 'solid', x: 48, y: 12, w: 6, h: 3 },
        { t: 'exit', x: 49, y: 10, fake: 'flee', to: [50.5, 10], range: 2.5 },
      ],
    },
  ],
};
