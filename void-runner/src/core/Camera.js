import { clamp, damp } from '../engine/math.js';

/**
 * Caméra 2D.
 *
 * Règles tenues, dans l'ordre de priorité :
 *  1. jamais de secousse involontaire — l'amortissement est exponentiel et
 *     indépendant du framerate ;
 *  2. anticipation douce dans le sens de la course (on voit venir le danger) ;
 *  3. une salle qui tient à l'écran est CENTRÉE, elle ne défile pas ;
 *  4. le tremblement d'impact est bref, borné, et jamais pendant un saut
 *     précis.
 */
export class Camera {
  constructor() {
    this.x = 0;
    this.y = 0;
    this.shakeX = 0;
    this.shakeY = 0;
    this.shake = 0;
    this.lookX = 0;
    this.seed = 1;
  }

  /** Recadre instantanément (début de niveau, téléportation). */
  snap(world, view) {
    const t = this.target(world, view);
    this.x = t.x;
    this.y = t.y;
    this.lookX = 0;
    this.shake = 0;
  }

  target(world, view) {
    const p = world.player;
    const maxX = Math.max(0, world.w - view.w);
    const maxY = Math.max(0, world.h - view.h);
    let x = p.cx - view.w / 2 + this.lookX;
    let y = p.cy - view.h * 0.55;
    if (world.w <= view.w) x = (world.w - view.w) / 2;
    else x = clamp(x, 0, maxX);
    if (world.h <= view.h) y = (world.h - view.h) / 2;
    else y = clamp(y, 0, maxY);
    return { x, y };
  }

  addShake(power) {
    this.shake = Math.min(1, this.shake + power);
  }

  update(world, view, dt) {
    const p = world.player;
    // Anticipation : la caméra regarde légèrement devant, proportionnellement
    // à la vitesse réelle. Pas d'à-coup quand on change de sens.
    const want = clamp(p.vx / 190, -1, 1) * 46;
    this.lookX = damp(this.lookX, want, 3.2, dt);

    const t = this.target(world, view);
    // Suivi plus vif verticalement : en gravité inversée, tarder à recadrer
    // rend le niveau illisible.
    this.x = damp(this.x, t.x, 9, dt);
    this.y = damp(this.y, t.y, 11, dt);

    if (this.shake > 0.001) {
      this.shake = Math.max(0, this.shake - dt * 3.4);
      const s = this.shake * this.shake * 7;
      this.seed = (this.seed * 1103515245 + 12345) & 0x7fffffff;
      const a = (this.seed / 0x7fffffff) * Math.PI * 2;
      this.shakeX = Math.cos(a) * s;
      this.shakeY = Math.sin(a) * s;
    } else {
      this.shakeX = 0;
      this.shakeY = 0;
    }
  }
}
