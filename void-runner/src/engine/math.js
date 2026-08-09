/** Petites aides mathématiques — sans dépendance, sans allocation superflue. */

export const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);
export const lerp = (a, b, t) => a + (b - a) * t;
export const sign = (v) => (v > 0 ? 1 : v < 0 ? -1 : 0);
export const approach = (v, target, delta) =>
  v < target ? Math.min(v + delta, target) : Math.max(v - delta, target);

/** Amortisseur exponentiel indépendant du framerate (caméra, UI). */
export const damp = (a, b, lambda, dt) => lerp(a, b, 1 - Math.exp(-lambda * dt));

/** Recouvrement de deux AABB (bords jointifs = pas de collision). */
export function aabb(ax, ay, aw, ah, bx, by, bw, bh) {
  return ax < bx + bw && ax + aw > bx && ay < by + bh && ay + ah > by;
}

/** Recouvrement avec marge (utile pour les zones de détection tolérantes). */
export function aabbPad(a, b, pad = 0) {
  return aabb(a.x - pad, a.y - pad, a.w + pad * 2, a.h + pad * 2, b.x, b.y, b.w, b.h);
}

/**
 * Générateur pseudo-aléatoire déterministe (mulberry32).
 * Utilisé UNIQUEMENT pour le décor et les particules, jamais pour le gameplay.
 */
export function makeRng(seed = 1) {
  let a = seed >>> 0;
  return function rng() {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Formatage humain d'un chrono : 21.43 -> "21.43 s" */
export function formatTime(sec) {
  if (sec == null || !isFinite(sec)) return '—';
  if (sec >= 60) {
    const m = Math.floor(sec / 60);
    return `${m}:${(sec % 60).toFixed(2).padStart(5, '0')}`;
  }
  return `${sec.toFixed(2)} s`;
}

/** 1284 -> "1 284" (séparateur français) */
export function formatInt(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}
