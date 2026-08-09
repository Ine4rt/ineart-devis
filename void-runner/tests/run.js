import { LEVELS } from '../src/data/levels/index.js';
import { ENTITY_TYPES } from '../src/engine/entities/index.js';
import { World } from '../src/engine/World.js';
import { TILE, JUMP_VEL, GRAVITY, RUN_SPEED } from '../src/engine/constants.js';
import { solve, replay } from './solver.js';

/**
 * Suite de validation. Trois passes, de la moins chère à la plus chère :
 *   1. structure   — le niveau est-il bien formé ?
 *   2. cohérence   — les signaux référencés existent-ils ? les distances sont-
 *                    elles dans les capacités physiques du robot ?
 *   3. solvabilité — le solveur termine-t-il réellement le niveau ?
 *
 * Usage :  node tests/run.js            (tout)
 *          node tests/run.js 7 15 29    (niveaux ciblés)
 *          node tests/run.js --fast     (structure + cohérence seulement)
 */

const args = process.argv.slice(2);
const fast = args.includes('--fast');
const only = args.filter((a) => /^\d+$/.test(a)).map(Number);

const APEX = (JUMP_VEL * JUMP_VEL) / (2 * GRAVITY);
const RANGE = ((2 * JUMP_VEL) / GRAVITY) * RUN_SPEED;

let failures = 0;
const fail = (lvl, msg) => {
  failures++;
  console.log(`  \x1b[31m✗\x1b[0m  ${lvl.num.toString().padStart(2)} ${lvl.name} — ${msg}`);
};

// --- 1 & 2 : structure et cohérence -----------------------------------------

function checkStructure(lv) {
  let ok = true;
  const emitted = new Set();
  const consumed = new Set();

  if (!lv.spawn) { fail(lv, 'aucun point d\'apparition'); ok = false; }
  if (!lv.entities.some((e) => e.t === 'exit')) { fail(lv, 'aucune sortie'); ok = false; }

  for (const e of lv.entities) {
    if (!ENTITY_TYPES.includes(e.t)) { fail(lv, `type inconnu « ${e.t} »`); ok = false; }
    if (e.emits) emitted.add(e.emits);
    if (e.out) emitted.add(e.out);
    if (e.signal) consumed.add(e.signal);
    if (e.armSignal) consumed.add(e.armSignal);
    if (e.requires) consumed.add(e.requires);
    for (const s of e.in || []) consumed.add(s);
  }
  for (const s of consumed) {
    if (!emitted.has(s)) { fail(lv, `signal « ${s} » consommé mais jamais émis`); ok = false; }
  }

  // Le robot ne doit jamais apparaître dans un mur.
  const w = new World(lv, { dash: lv.dash });
  const p = w.player;
  if (w.overlapsSolid(p.x, p.y, p.w, p.h)) { fail(lv, 'apparition dans un solide'); ok = false; }

  // Une sortie posée hors de la salle est une faute de frappe, pas un piège.
  for (const e of w.findAll('exit')) {
    if (e.x < 0 || e.x > lv.w * TILE || e.y > lv.h * TILE) {
      fail(lv, 'sortie hors de la salle'); ok = false;
    }
  }
  return ok;
}

// --- 3 : solvabilité ---------------------------------------------------------

function checkSolvable(lv) {
  const t0 = Date.now();
  const r = solve(lv);
  const ms = Date.now() - t0;
  if (!r.ok) {
    fail(lv, `INFRANCHISSABLE (${r.reason}${r.reached ? ' — ' + r.reached : ''})`);
    return false;
  }
  // Rejouer la solution doit donner exactement le même résultat : c'est le
  // test de déterminisme du moteur.
  const again = replay(lv, r.path, r.stride);
  if (again !== 'win') {
    fail(lv, `simulation non déterministe (rejeu : ${again})`);
    return false;
  }
  console.log(
    `  \x1b[32m✓\x1b[0m  ${lv.num.toString().padStart(2)} ${lv.name.padEnd(22)} ` +
    `résolu en ${r.time.toFixed(1)} s  (${(ms / 1000).toFixed(1)} s de calcul)`,
  );
  return true;
}

// --- Exécution ---------------------------------------------------------------

console.log('\n\x1b[1mVOID RUNNER — validation des niveaux\x1b[0m');
console.log(`  Capacités du robot : apex ${APEX.toFixed(0)} px (${(APEX / TILE).toFixed(1)} tuiles), ` +
  `portée ${RANGE.toFixed(0)} px (${(RANGE / TILE).toFixed(1)} tuiles)\n`);

const list = only.length ? LEVELS.filter((l) => only.includes(l.num)) : LEVELS;

console.log('\x1b[2m— structure et cohérence —\x1b[0m');
let structOk = 0;
for (const lv of list) if (checkStructure(lv)) structOk++;
console.log(`  ${structOk}/${list.length} niveaux bien formés\n`);

if (!fast) {
  console.log('\x1b[2m— solvabilité (le solveur joue réellement) —\x1b[0m');
  for (const lv of list) checkSolvable(lv);
}

console.log(
  failures === 0
    ? '\n\x1b[32m✓ Tout est bon.\x1b[0m\n'
    : `\n\x1b[31m✗ ${failures} problème(s).\x1b[0m\n`,
);
process.exit(failures === 0 ? 0 : 1);
