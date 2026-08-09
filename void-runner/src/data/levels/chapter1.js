/**
 * CHAPITRE 1 — STATION ABANDONNÉE (niveaux 1 à 6)
 *
 * Rôle : apprendre à jouer sans jamais afficher un tutoriel.
 * Chaque niveau enseigne UNE chose, et la première trahison arrive dès le
 * niveau 1 — parce que la promesse du jeu doit être tenue immédiatement.
 *
 * Repères de level design — MESURÉS sur le moteur, affichés par `npm test` :
 *   - hauteur de saut maximale ....... 83 px  ≈ 3,4 tuiles
 *   - portée horizontale maximale .... 114 px ≈ 4,8 tuiles
 *   - surface de sol standard ........ y = 12
 * (La formule continue annonce 87 px : elle ment de 9 %, le pas fixe applique
 * la gravité dès la frame du saut.)
 *
 * Toute plateforme au-delà de ces valeurs est un bug de conception, pas un
 * défi — et `npm test` le signale. Mais « franchissable » ne suffit pas dans
 * un chapitre d'introduction : `node tools/tolerance.js N` mesure la marge
 * d'erreur laissée au joueur, et `node tools/playtest.js` vérifie qu'un joueur
 * prudent — qui court, s'arrête au bord, attend et saute — s'en sort.
 */

export const CHAPTER_1 = {
  id: 1,
  name: 'Station abandonnée',
  subtitle: 'Secteur A — maintenance',
  palette: 'station',
  levels: [
    // ---------------------------------------------------------------- 01 ----
    {
      name: 'Premier contact',
      hint: 'Rejoins la porte.',
      // La mort par laser est incompréhensible sans un mot : le joueur ne peut
      // pas deviner que c'est SA hauteur de saut qui l'a déclenché.
      deathHints: { fall: 'Il fallait sauter.', burn: 'Saut trop haut. Une tape suffit.' },
      // Leçon : « ce que tu vois n'est pas le niveau ». Le sol s'ouvre, puis
      // le réflexe naturel (sauter haut) est puni. La bonne réponse est un
      // petit saut — ce qui enseigne du même coup le saut à hauteur variable.
      w: 26, h: 15,
      spawn: { x: 2, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 16, h: 3 },
        { t: 'solid', x: 18, y: 12, w: 8, h: 3 },
        // La trappe : deux panneaux qui s'escamotent LATÉRALEMENT dans le sol.
        // (Première version : un panneau descendant — le robot descendait avec
        // et atterrissait tranquillement dessus. Une trappe doit lâcher, pas
        // accompagner.)
        { t: 'door', x: 16, y: 12, w: 1, h: 1, slide: 'left', travel: 1, speed: 260, signal: 'trappe' },
        { t: 'door', x: 17, y: 12, w: 1, h: 1, slide: 'right', travel: 1, speed: 260, signal: 'trappe' },
        { t: 'zone', x: 11, y: 9, w: 3, h: 3, mode: 'once', emits: 'trappe' },

        // Émetteur mural. Décoratif… jusqu'à ce qu'il ne le soit plus.
        { t: 'laser', x: 0.3, y: 8.2, w: 0.6, h: 0.5, dir: 'right', len: 26, armSignal: 'altitude' },
        // Le détecteur d'altitude est calé à 74 px de montée, sur les 79,5 px
        // réellement atteignables : une tape passe dessous, un appui maintenu
        // le franchit. C'est bien la DURÉE d'appui qui décide, pas l'adresse.
        //
        // Bornes mesurées (`node tools/tolerance.js 1`), non estimées :
        // appui sûr jusqu'à ~215 ms, mortel à partir de ~295 ms. La première
        // version se déclenchait dès 57 px : la marge tombait à 150 ms, soit
        // moins qu'une tape de pouce. Injouable, et c'est ce qui a bloqué.
        { t: 'zone', x: 14.2, y: 5.5, w: 5, h: 2.5, mode: 'once', emits: 'altitude' },

        { t: 'deco', x: 4, y: 10.6, w: 1.5, h: 1.4, kind: 'crate' },
        { t: 'sign', x: 3.5, y: 9.4, w: 3, h: 1, text: 'SORTIE  →', always: true },
        { t: 'exit', x: 22.5, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 02 ----
    {
      name: 'Trois plateformes',
      hint: 'Ne t\'attarde pas.',
      deathHints: { fall: 'La deuxième plateforme ne tient pas. Repars aussitôt.' },
      // Leçon : un appui n'est pas un acquis. La plateforme centrale décroche
      // au contact. Elle descend assez lentement pour laisser une fenêtre
      // d'environ 0,8 s — largement suffisant, mais seulement si on réagit.
      w: 26, h: 15,
      spawn: { x: 2, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 5, h: 3 },
        { t: 'mover', x: 8, y: 12, w: 3, h: 1, path: [[8, 20]], speed: 105, mode: 'once', trigger: 'stand', style: 'loose' },
        { t: 'solid', x: 14, y: 12, w: 12, h: 3 },
        { t: 'deco', x: 8.2, y: 11.3, w: 0.6, h: 0.6, kind: 'bolt' },
        { t: 'deco', x: 10.2, y: 11.3, w: 0.6, h: 0.6, kind: 'bolt' },
        { t: 'sign', x: 1.5, y: 9.4, w: 3, h: 1, text: 'PLANCHER : 92 %', always: true },
        { t: 'exit', x: 22.5, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 03 ----
    {
      name: 'Sas de sécurité',
      hint: 'Le bouton ouvre le sas.',
      deathHints: { crush: 'Le verrou n\'était pas stabilisé. Attends le voyant.' },
      // Leçon : la patience est une commande de jeu.
      // Le sas s'ouvre immédiatement mais son verrou magnétique met 2,6 s à se
      // stabiliser. Franchir avant = écrasement. Le panneau le dit — personne
      // ne le lit au premier essai, tout le monde le lit au deuxième.
      w: 26, h: 15,
      spawn: { x: 2, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 26, h: 3 },
        { t: 'solid', x: 13, y: 3, w: 2, h: 7 },

        { t: 'button', x: 6, y: 11.6, w: 2, h: 0.4, mode: 'once', emits: 'cmd' },
        { t: 'door', x: 13, y: 10, w: 2, h: 2, slide: 'up', travel: 2, speed: 420, signal: 'ouvert' },

        // Chaîne logique : ouvert = cmd ET NON claque ; claque = dansSas ET NON stable
        { t: 'timer', from: 'cmd', delay: 2.6, emits: 'stable' },
        { t: 'zone', x: 12.8, y: 10, w: 2.4, h: 2, mode: 'stay', emits: 'dansSas' },
        { t: 'logic', op: 'not', in: ['stable'], emits: 'instable' },
        { t: 'logic', op: 'and', in: ['dansSas', 'instable'], emits: 'claque' },
        { t: 'logic', op: 'not', in: ['claque'], emits: 'nonClaque' },
        { t: 'logic', op: 'and', in: ['cmd', 'nonClaque'], emits: 'ouvert' },

        { t: 'deco', x: 15.3, y: 8.2, w: 0.8, h: 0.8, kind: 'lamp', signal: 'stable' },
        { t: 'deco', x: 12, y: 8.2, w: 0.8, h: 0.8, kind: 'lamp', signal: 'stable' },
        { t: 'sign', x: 9, y: 9.4, w: 4, h: 1, text: 'SAS 1\nVERROU : 2,6 s', always: true },
        { t: 'exit', x: 22.5, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 04 ----
    {
      name: 'Maintenance',
      hint: 'Le faisceau a un rythme.',
      deathHints: { burn: 'Attends qu\'il s\'éteigne. Il a un cycle régulier.' },
      // Leçon : lire un cycle. Aucun mensonge ici — c'est une respiration.
      // Un jeu qui trahit à chaque niveau devient du bruit ; il faut des
      // niveaux honnêtes pour que la trahison garde sa valeur.
      w: 26, h: 15,
      spawn: { x: 2, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 9, h: 3 },
        { t: 'solid', x: 12, y: 12, w: 4, h: 3 },
        { t: 'solid', x: 19, y: 12, w: 7, h: 3 },
        // Deux faisceaux déphasés : on ne peut pas franchir les deux d'un trait,
        // il faut s'arrêter sur l'îlot. Le joueur choisit QUAND il s'expose —
        // c'est la condition pour qu'un piège de timing reste juste.
        //
        // Les phases sont calées pour que le premier faisceau soit FRANCHEMENT
        // allumé à l'instant où le robot arrive (t ≈ 1,0 s), puis franchement
        // éteint 0,4 s plus tard. Version initiale : il s'éteignait pile à
        // l'arrivée, si bien qu'on passait ou qu'on mourait à trois frames
        // près, sans rien pouvoir en apprendre. Une mort doit enseigner ;
        // celle-là ne faisait que trancher.
        { t: 'laser', x: 10.2, y: 4.5, w: 0.5, h: 0.6, dir: 'down', len: 9, cycle: [0.9, 1.8], phase: 2.2 },
        { t: 'laser', x: 17.2, y: 4.5, w: 0.5, h: 0.6, dir: 'down', len: 9, cycle: [0.9, 1.8], phase: 1.3 },
        { t: 'deco', x: 9.8, y: 3.4, w: 1.4, h: 1.1, kind: 'panel' },
        { t: 'deco', x: 16.8, y: 3.4, w: 1.4, h: 1.1, kind: 'panel' },
        { t: 'sign', x: 4, y: 9.4, w: 4, h: 1, text: 'FAISCEAUX\nINTERMITTENTS', always: true },
        { t: 'exit', x: 22.5, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 05 ----
    {
      name: 'Sol fatigué',
      hint: 'Continue d\'avancer.',
      // Leçon : la vitesse comme outil. Un couloir de dalles qui s'effondrent,
      // puis un saut. Le piège : la dernière dalle avant le vide est SOLIDE —
      // beaucoup de joueurs sautent trop tôt par excès de prudence.
      w: 30, h: 15,
      spawn: { x: 1.5, y: 10 },
      entities: [
        { t: 'solid', x: 0, y: 12, w: 5, h: 3 },
        ...Array.from({ length: 11 }, (_, i) => ({
          t: 'crumble', x: 5 + i, y: 12, w: 1, h: 1, delay: 0.38, respawn: 2.4,
        })),
        { t: 'solid', x: 16, y: 12, w: 1, h: 3 },
        { t: 'solid', x: 20, y: 12, w: 10, h: 3 },
        { t: 'sign', x: 2, y: 9.4, w: 3.5, h: 1, text: 'DALLES : FATIGUE 87 %', always: true },
        { t: 'deco', x: 17.5, y: 13.5, w: 2, h: 0.4, kind: 'pipe' },
        { t: 'exit', x: 26.5, y: 10 },
      ],
    },

    // ---------------------------------------------------------------- 06 ----
    {
      name: 'Secteur A — sortie',
      hint: 'Tout ce que tu as appris.',
      // Final de chapitre : trappe + dalle mobile + cycle laser + patience.
      // Aucune mécanique nouvelle : on vérifie que le vocabulaire est acquis.
      w: 34, h: 15,
      spawn: { x: 1.5, y: 10 },
      // `solve` décrit l'itinéraire prévu par le concepteur. Il ne sert QU'aux
      // tests : le solveur doit prouver que cet itinéraire est physiquement
      // réalisable. Il n'apparaît jamais dans le jeu.
      solve: [[11.5, 11.4], [15.5, 11.4], [15.5, 7.4], [21, 7.4], [27, 7.4]],
      entities: [
        { t: 'solid', x: 0, y: 12, w: 8, h: 3 },
        { t: 'door', x: 8, y: 12, w: 1, h: 1, slide: 'left', travel: 1, speed: 260, signal: 'trappe' },
        { t: 'door', x: 9, y: 12, w: 1, h: 1, slide: 'right', travel: 1, speed: 260, signal: 'trappe' },
        { t: 'zone', x: 5, y: 9, w: 2, h: 3, mode: 'once', emits: 'trappe' },
        { t: 'solid', x: 10, y: 12, w: 3, h: 3 },

        { t: 'mover', x: 14, y: 11, w: 3, h: 1, path: [[14, 7], [14, 11]], speed: 62, mode: 'loop', wait: 0.6 },
        { t: 'solid', x: 19, y: 7, w: 4, h: 1 },
        { t: 'laser', x: 25.4, y: 6.3, w: 0.5, h: 0.5, dir: 'left', len: 4, cycle: [0.9, 1.4] },

        { t: 'crumble', x: 23, y: 7, w: 1, h: 1, delay: 0.4 },
        { t: 'crumble', x: 24, y: 7, w: 1, h: 1, delay: 0.4 },
        { t: 'solid', x: 25, y: 7, w: 9, h: 8 },
        { t: 'exit', x: 30, y: 5 },
        { t: 'sign', x: 10.4, y: 9.6, w: 2.4, h: 1, text: 'ASCENSEUR' },

        // Salle secrète : une alcôve sous la passerelle, invisible d'en haut.
        { t: 'solid', x: 17, y: 13.4, w: 6, h: 0.4 },
        { t: 'zone', x: 19, y: 12.6, w: 2, h: 0.8, mode: 'once', remember: 'secret', emits: 'secret' },
        { t: 'deco', x: 19, y: 12.5, w: 2, h: 0.9, kind: 'gift' },
      ],
    },
  ],
};
