import { World } from '../src/engine/World.js';
import { getLevel } from '../src/data/levels/index.js';

/**
 * MESURE DE LA FENÊTRE DE TOLÉRANCE.
 *
 * Le solveur répond à « existe-t-il un chemin ? ». Cet outil répond à la
 * question qui compte vraiment pour le joueur : « de combien puis-je me
 * tromper ? »
 *
 * Il joue la politique la plus naturelle du monde — courir à droite, et taper
 * une fois sur saut — pour toutes les combinaisons (instant du saut × durée
 * d'appui), et dessine la carte des issues.
 *
 * Repère : sur mobile, une tape dure typiquement 60 à 150 ms, un appui
 * volontaire dépasse 250 ms. Une salle d'introduction dont la fenêtre d'appui
 * est plus étroite que 150 ms est injouable, quoi qu'en dise le solveur.
 *
 *   node tools/tolerance.js 1
 */

const num = Number(process.argv[2] || 1);
const level = getLevel(num);
if (!level) { console.error(`Niveau ${num} inconnu.`); process.exit(1); }

const MAX_FRAMES = 60 * 12;
const SYM = { win: '\x1b[32m·\x1b[0m', dead: '\x1b[31m×\x1b[0m', run: '\x1b[90m-\x1b[0m' };

/** Court à droite, saute à `at`, relâche après `hold` frames. */
function attempt(at, hold) {
  const w = new World(level, { dash: level.dash === true });
  for (let f = 0; f < MAX_FRAMES; f++) {
    const jump = f >= at && f < at + hold;
    w.step({
      left: false, right: true, jump,
      jumpPressed: f === at, dashPressed: false,
    });
    if (w.state !== 'run') break;
  }
  return { state: w.state, kind: w.deathKind, x: w.player.x };
}

const jumpFrames = [];
for (let f = 0; f <= 150; f++) jumpFrames.push(f);
const holds = [];
for (let h = 1; h <= 30; h++) holds.push(h);

const rows = [];
let best = null;
for (const at of jumpFrames) {
  let row = '';
  let wins = 0;
  let firstWin = null, lastWin = null;
  for (const h of holds) {
    const r = attempt(at, h);
    row += SYM[r.state] || '?';
    if (r.state === 'win') {
      wins++;
      if (firstWin === null) firstWin = h;
      lastWin = h;
    }
  }
  if (wins) {
    rows.push({ at, row, wins, firstWin, lastWin });
    if (!best || wins > best.wins) best = { at, wins, firstWin, lastWin };
  }
}

console.log(`\n\x1b[1mSalle ${num} — ${level.name}\x1b[0m`);
console.log('  Politique testée : courir à droite + une seule tape sur saut.');
console.log('  Colonnes = durée d\'appui (1 à 30 frames, soit 17 à 500 ms).\n');

if (!rows.length) {
  console.log('  \x1b[31mAUCUNE combinaison ne franchit la salle avec ce schéma simple.\x1b[0m');
  console.log('  (Normal si la salle demande autre chose qu\'un saut : attendre, reculer…)\n');
  process.exit(0);
}

for (const r of rows) {
  const ms = (r.at / 60 * 1000).toFixed(0).padStart(4);
  console.log(`  saut à ${ms} ms  ${r.row}`);
}

const holdMs = (f) => (f / 60 * 1000);
const winMs = holdMs(best.lastWin) - holdMs(best.firstWin - 1);
const posFrames = rows.length;
console.log(`
  \x1b[1mFenêtres au meilleur instant de saut (${(best.at / 60 * 1000).toFixed(0)} ms) :\x1b[0m
    durée d'appui tolérée : ${holdMs(best.firstWin).toFixed(0)} à ${holdMs(best.lastWin).toFixed(0)} ms  → \x1b[1m${winMs.toFixed(0)} ms\x1b[0m de marge
    instants de saut qui passent : ${posFrames} frames → \x1b[1m${(posFrames / 60 * 1000).toFixed(0)} ms\x1b[0m de marge

  Objectif pour une salle de tutoriel : >= 200 ms d'appui ET >= 200 ms d'instant.
`);
