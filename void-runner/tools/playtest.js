import { World } from '../src/engine/World.js';
import { LEVELS, getLevel } from '../src/data/levels/index.js';

/**
 * JOUEUR PRUDENT AUTOMATIQUE.
 *
 * Le solveur cherche *une* solution, aussi acrobatique soit-elle. Il valide
 * donc des salles qu'aucun humain ne franchirait de la même façon.
 *
 * Ce robot-ci joue comme quelqu'un de raisonnable qui découvre la salle :
 * il court à droite, s'arrête au bord de chaque vide, patiente, puis saute.
 * Deux paramètres seulement — le temps d'attente au bord et la durée d'appui —
 * balayés exhaustivement.
 *
 * Ce qu'on lit dans le résultat : le nombre de temps d'attente qui mènent à la
 * sortie. Zéro sur une salle d'introduction = la salle exige autre chose que
 * « regarder, attendre, sauter », et il faut savoir si c'est délibéré.
 *
 *   node tools/playtest.js        (les 30 salles)
 *   node tools/playtest.js 4      (une seule, en détail)
 */

const MAX_FRAMES = 60 * 40;

/** @param wait frames d'attente au bord · @param hold durée d'appui du saut */
function play(level, wait, hold) {
  const w = new World(level, { dash: level.dash === true });
  let waited = 0;
  let jumpAt = -1;
  let lastX = -1;
  let stuck = 0;

  for (let f = 0; f < MAX_FRAMES; f++) {
    const p = w.player;
    const g = p.gravDir;
    // Y a-t-il encore du sol juste devant ? (sondé un demi-corps plus loin)
    const ahead = w.overlapsSolid(p.x + p.w * 0.6, p.y + g * 3, p.w, p.h);
    const atLedge = p.grounded && !ahead;

    let right = true;
    let jumping = jumpAt >= 0 && f < jumpAt + hold;

    if (atLedge && jumpAt < 0) {
      if (waited < wait) { right = false; waited++; }
      else { jumpAt = f; jumping = true; waited = 0; }
    }
    if (jumpAt >= 0 && f >= jumpAt + hold + 6) jumpAt = -1;

    w.step({
      left: false, right, jump: jumping,
      jumpPressed: f === jumpAt, dashPressed: false,
    });
    if (w.state === 'win') return { ok: true, t: w.time };
    if (w.state === 'dead') return { ok: false, why: w.deathKind };

    // Détection d'immobilité : inutile de simuler 40 s pour rien.
    if (Math.abs(p.x - lastX) < 0.3) { if (++stuck > 60 * 12) break; } else stuck = 0;
    lastX = p.x;
  }
  return { ok: false, why: 'temps écoulé' };
}

function scan(level) {
  const waits = [];
  let best = null;
  for (let wait = 0; wait <= 200; wait += 2) {
    for (const hold of [6, 10, 16, 24]) {
      const r = play(level, wait, hold);
      if (r.ok) { waits.push(wait); if (!best || r.t < best.t) best = r; break; }
    }
  }
  return { waits, best };
}

const only = process.argv.slice(2).filter((a) => /^\d+$/.test(a)).map(Number);
const list = only.length ? only.map(getLevel) : LEVELS;

console.log('\n\x1b[1mVOID RUNNER — passage par un joueur prudent\x1b[0m');
console.log('  Court à droite, s\'arrête à chaque rebord, patiente, saute.\n');

for (const lv of list) {
  const { waits, best } = scan(lv);
  const ms = (waits.length * 2 / 60 * 1000).toFixed(0);
  const tag = waits.length === 0 ? '\x1b[33m—  \x1b[0m' : '\x1b[32m ok\x1b[0m';
  const detail = waits.length
    ? `${waits.length} temps d'attente viables (${ms} ms cumulés), meilleur passage ${best.t.toFixed(1)} s`
    : 'aucun passage par ce schéma simple — la salle demande autre chose';
  console.log(`  ${tag} ${String(lv.num).padStart(2)} ${lv.name.padEnd(22)} ${detail}`);
}
console.log('');
