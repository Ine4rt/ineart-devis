/**
 * Retour haptique.
 *
 * Volontairement pauvre : quatre motifs, pas un de plus. Une vibration à
 * chaque évènement transforme le téléphone en jouet cassé et vide la batterie.
 * On vibre pour la mort, l'impact lourd, la réussite et le mécanisme majeur —
 * c'est-à-dire pour ce que le joueur doit sentir même en jouant sans le son.
 */
const PATTERNS = {
  death: [18, 30, 45],
  impact: [12],
  win: [10, 40, 10, 40, 60],
  mech: [8],
};

export class Haptics {
  constructor() {
    this.enabled = true;
    this.ok = typeof navigator !== 'undefined' && 'vibrate' in navigator;
    this.last = 0;
  }

  fire(kind) {
    if (!this.enabled || !this.ok) return;
    const p = PATTERNS[kind];
    if (!p) return;
    // Anti-spam : jamais deux vibrations à moins de 60 ms.
    const now = Date.now();
    if (now - this.last < 60) return;
    this.last = now;
    try { navigator.vibrate(p); } catch { /* refusé par la plateforme */ }
  }
}
