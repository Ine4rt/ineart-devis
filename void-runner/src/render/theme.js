/**
 * Charte graphique de VOID RUNNER — sci-fi cartoon, formes pleines, contours
 * nets, lumière additive. Volontairement à l'opposé du pixel-art rétro.
 *
 * La couleur est un contrat passé avec le joueur, et il n'est jamais rompu :
 *   cyan   → sûr, interactif, la sortie ;
 *   violet → structure, décor inerte ;
 *   ORANGE → ça tue, maintenant ;
 *   ROUGE  → ça va tuer, dans un instant (pré-signal) ;
 *   blanc  → le robot, et lui seul.
 *
 * Un joueur qui meurt sans avoir vu d'orange ni de rouge est un joueur qu'on a
 * trahi : c'est le seul cas où le jeu est en tort.
 */

export const C = {
  night: '#0b0d20',
  night2: '#151a3a',
  deep: '#070818',
  wall: '#232a55',
  wallLit: '#39407a',
  wallEdge: '#4b53a0',
  violet: '#7b5cff',
  cyan: '#3fe8ff',
  cyanDim: '#1c8ea8',
  white: '#eaf6ff',
  danger: '#ff8a3d',
  dangerDeep: '#c94a10',
  alert: '#ff3b5c',
  gold: '#ffd166',
  ink: '#05060f',
};

/** Variantes d'ambiance par chapitre — même charte, autre température. */
export const PALETTES = {
  station: { bg0: '#0b0d20', bg1: '#171b3d', grid: 'rgba(123,92,255,0.10)', haze: 'rgba(63,232,255,0.05)' },
  lab: { bg0: '#0a1226', bg1: '#123049', grid: 'rgba(63,232,255,0.10)', haze: 'rgba(63,232,255,0.07)' },
  grav: { bg0: '#140b26', bg1: '#2a1650', grid: 'rgba(163,120,255,0.12)', haze: 'rgba(163,120,255,0.06)' },
  core: { bg0: '#1a0d16', bg1: '#3a1526', grid: 'rgba(255,138,61,0.09)', haze: 'rgba(255,138,61,0.06)' },
  void: { bg0: '#06060e', bg1: '#12102a', grid: 'rgba(255,255,255,0.05)', haze: 'rgba(123,92,255,0.05)' },
};

export function palette(name) {
  return PALETTES[name] || PALETTES.station;
}

/** Rectangle arrondi — la primitive de base de tout le rendu. */
export function roundRect(ctx, x, y, w, h, r) {
  const rr = Math.min(r, Math.abs(w) / 2, Math.abs(h) / 2);
  ctx.beginPath();
  ctx.moveTo(x + rr, y);
  ctx.arcTo(x + w, y, x + w, y + h, rr);
  ctx.arcTo(x + w, y + h, x, y + h, rr);
  ctx.arcTo(x, y + h, x, y, rr);
  ctx.arcTo(x, y, x + w, y, rr);
  ctx.closePath();
}

/**
 * Halo lumineux. `shadowBlur` est coûteux sur mobile : on l'active par petites
 * zones seulement, jamais sur un remplissage plein écran.
 */
export function glow(ctx, color, blur, draw) {
  ctx.save();
  ctx.shadowColor = color;
  ctx.shadowBlur = blur;
  draw();
  ctx.restore();
}
