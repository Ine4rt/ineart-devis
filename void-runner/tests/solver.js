import { World } from '../src/engine/World.js';
import { TILE, FIXED_DT } from '../src/engine/constants.js';

/**
 * SOLVEUR DE VALIDATION
 *
 * Pourquoi ce fichier existe : dans un jeu qui repose entièrement sur la mort,
 * la pire faute de conception possible est un niveau infranchissable. On ne
 * peut pas le détecter à l'œil — un gouffre d'une demi-tuile de trop ne se
 * voit pas sur un plan.
 *
 * Le solveur joue donc réellement chaque niveau, avec le moteur de production,
 * par recherche en faisceau (beam search) sur les actions possibles. S'il
 * termine, le niveau est franchissable ; s'il échoue, c'est un bug de level
 * design, pas un défi.
 *
 * Il ne prétend pas jouer bien — juste prouver qu'un chemin existe.
 */

const ACTIONS = [
  { left: false, right: true, jump: false },
  { left: false, right: true, jump: true },
  { left: true, right: false, jump: false },
  { left: true, right: false, jump: true },
  { left: false, right: false, jump: false },
  { left: false, right: false, jump: true },
];
const DASH_ACTIONS = [
  { left: false, right: true, jump: true, dash: true },
  { left: true, right: false, jump: true, dash: true },
];

function dist(p, wp) {
  return Math.hypot(p.cx - wp[0], p.cy - wp[1]);
}

/**
 * Jalons de l'itinéraire, puis la sortie.
 *
 * La sortie est suivie par INDICE et non par position : une sortie qui fuit
 * (niveaux 12 et 30) se déplace pendant la partie, et viser sa position de
 * départ ferait tourner le solveur en rond juste devant elle.
 */
function makeTargets(level, world) {
  const pts = (level.solve || []).map((p) => [p[0] * TILE, p[1] * TILE]);
  let exitIdx = -1;
  world.entities.forEach((e, i) => {
    if (e.type === 'exit' && (!e.fake || e.fake === 'flee')) exitIdx = i;
  });
  const count = pts.length + (exitIdx >= 0 ? 1 : 0);
  return {
    count,
    at(w, k) {
      if (k < pts.length) return pts[k];
      const e = w.entities[exitIdx];
      return [e.x + e.w / 2, e.y + e.h / 2];
    },
  };
}

export function solve(level, opts = {}) {
  const beamWidth = opts.beam ?? 64;
  const maxPerBucket = opts.maxPerBucket ?? 6;
  const stride = opts.stride ?? 8; // frames par décision
  const maxSeconds = opts.maxSeconds ?? 55;
  const maxDepth = Math.ceil((maxSeconds / FIXED_DT) / stride);

  const root = new World(level, { dash: level.dash === true });
  const wps = makeTargets(level, root);
  if (!wps.count) return { ok: false, reason: 'aucune sortie atteignable déclarée' };

  const actions = level.dash ? [...ACTIONS, ...DASH_ACTIONS] : ACTIONS;

  let beam = [{ w: root, k: 0, best: dist(root.player, wps.at(root, 0)), path: [] }];
  let expansions = 0;

  for (let depth = 0; depth < maxDepth; depth++) {
    const next = [];
    for (const node of beam) {
      for (const a of actions) {
        const w = node.w.clone();
        let k = node.k;
        let dead = false;
        for (let f = 0; f < stride; f++) {
          const input = {
            left: a.left,
            right: a.right,
            jump: a.jump,
            // Le front montant n'existe qu'à la première frame du palier :
            // sinon le solveur « re-sauterait » en boucle.
            jumpPressed: a.jump && f === 0,
            dashPressed: a.dash === true && f === 0,
          };
          w.step(input);
          expansions++;
          if (w.state === 'win') {
            return {
              ok: true, time: w.time, expansions,
              path: [...node.path, a], stride,
            };
          }
          if (w.state === 'dead') { dead = true; break; }
          while (k < wps.count - 1 && dist(w.player, wps.at(w, k)) < TILE * 1.4) k++;
        }
        if (dead) continue;
        const d = dist(w.player, wps.at(w, k));
        next.push({ w, k, best: d, path: [...node.path, a] });
      }
    }
    if (!next.length) return { ok: false, reason: 'toutes les branches meurent', depth, expansions };

    // Tri : jalon atteint d'abord, puis proximité de la cible.
    next.sort((a, b) => (b.k - a.k) || (a.best - b.best));

    // Déduplication + quota par zone.
    //
    // Le quota est essentiel. Sans lui, le faisceau se remplit uniquement des
    // états les plus avancés — et devant une presse hydraulique, « le plus
    // avancé » signifie « celui qui meurt dans trois dixièmes de seconde ».
    // Le faisceau s'effondrait alors d'un coup et déclarait à tort le niveau
    // infranchissable. Garder quelques états en retrait, c'est autoriser le
    // solveur à faire ce que fait un joueur humain : reculer et attendre.
    const seen = new Set();
    const bucket = new Map();
    beam = [];
    for (const n of next) {
      const p = n.w.player;
      const key = `${n.k}|${Math.round(p.x / 6)}|${Math.round(p.y / 6)}|${Math.round(p.vx / 60)}|${Math.round(p.vy / 90)}|${p.gravDir}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const bk = `${n.k}|${Math.floor(p.x / (TILE * 2))}`;
      const used = bucket.get(bk) || 0;
      if (used >= maxPerBucket) continue;
      bucket.set(bk, used + 1);
      beam.push(n);
      if (beam.length >= beamWidth) break;
    }
  }
  const b = beam[0];
  return {
    ok: false,
    reason: 'temps imparti dépassé',
    reached: b ? `jalon ${b.k + 1}/${wps.count}, à ${(b.best / TILE).toFixed(1)} tuiles` : '—',
    expansions,
  };
}

/**
 * ROBUSTESSE : une erreur d'une frame est-elle fatale ?
 *
 * Le solveur prouve qu'un chemin existe. Ça ne dit rien de ce qui compte pour
 * un joueur : la marge d'erreur. On reprend donc la solution trouvée et on la
 * décale d'une frame — en retard, puis en avance — à chaque point de décision,
 * et on compte combien de ces variantes franchissent quand même la salle.
 *
 * Une salle à 100 % pardonne un décalage d'une frame n'importe où. Une salle
 * sous 50 % exige une exécution à la frame près à un moment donné : c'est
 * acceptable en fin de jeu, jamais dans un chapitre d'introduction.
 */
export function robustness(level, path, stride) {
  const frames = [];
  for (const a of path) {
    for (let f = 0; f < stride; f++) frames.push({ a, first: f === 0 });
  }

  const run = (list) => {
    const w = new World(level, { dash: level.dash === true });
    for (const { a, first } of list) {
      w.step({
        left: a.left, right: a.right, jump: a.jump,
        jumpPressed: a.jump && first,
        dashPressed: a.dash === true && first,
      });
      if (w.state !== 'run') break;
    }
    return w.state === 'win';
  };

  let tried = 0;
  let survived = 0;
  for (let i = 0; i < path.length; i++) {
    const at = i * stride;
    // Une frame de retard : le joueur a réagi un poil trop tard.
    const late = frames.slice(0, at).concat([frames[at]], frames.slice(at));
    // Une frame d'avance : il a réagi un poil trop tôt.
    const early = frames.slice(0, at).concat(frames.slice(at + 1));
    for (const v of [late, early]) { tried++; if (run(v)) survived++; }
  }
  return tried ? survived / tried : 1;
}

/** Rejoue une solution trouvée (utile pour vérifier le déterminisme). */
export function replay(level, path, stride) {
  const w = new World(level, { dash: level.dash === true });
  for (const a of path) {
    for (let f = 0; f < stride; f++) {
      w.step({
        left: a.left, right: a.right, jump: a.jump,
        jumpPressed: a.jump && f === 0,
        dashPressed: a.dash === true && f === 0,
      });
      if (w.state !== 'run') return w.state;
    }
  }
  return w.state;
}
