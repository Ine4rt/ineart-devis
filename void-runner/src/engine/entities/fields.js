import { Entity } from './base.js';
import { TILE } from '../constants.js';

/**
 * Ventilateur / soufflerie.
 * Applique une accélération constante tant que le robot est dans le flux.
 * Le flux est arrêté par les solides : on peut donc se mettre à l'abri.
 */
export class Fan extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.dir = spec.dir || 'up';
    this.force = spec.force ?? 2600;
    this.style = 'fan';
    this.spin = 0;
  }

  preStep(world, dt) {
    this.on = this.gateOpen(world);
    this.spin += this.on ? dt * 14 : dt * 1.5;
  }

  postStep(world) {
    if (!this.on || !this.touchesPlayer(world)) return;
    const p = world.player;
    switch (this.dir) {
      case 'down': p.windY += this.force; break;
      case 'left': p.windX -= this.force; break;
      case 'right': p.windX += this.force; break;
      // Une soufflerie ascendante doit pouvoir compenser la gravité, sinon
      // elle n'est qu'un ralentisseur — d'où une force par défaut > GRAVITY.
      default: p.windY -= this.force;
    }
  }
}

/**
 * Champ technique : ralentissement, accélération, magnétisme.
 * kind : 'slow' | 'fast' | 'magnet' | 'ice'
 */
export class Field extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.kind = spec.kind || 'slow';
    this.amount = spec.amount ?? (this.kind === 'slow' ? 0.45 : 1.75);
    this.force = spec.force ?? 900;
    this.anchor = spec.anchor || null; // [x,y] en tuiles, pour le magnétisme
    this.style = 'field';
  }

  postStep(world) {
    if (!this.gateOpen(world) || !this.touchesPlayer(world)) return;
    const p = world.player;
    if (this.kind === 'slow' || this.kind === 'fast') {
      p.speedMul = this.amount;
    } else if (this.kind === 'ice') {
      p.frictionMul = 0.06;
    } else if (this.kind === 'magnet') {
      const ax = (this.anchor ? this.anchor[0] * TILE : this.x + this.w / 2);
      const ay = (this.anchor ? this.anchor[1] * TILE : this.y + this.h / 2);
      const dx = ax - p.cx;
      const dy = ay - p.cy;
      const d = Math.hypot(dx, dy) || 1;
      p.windX += (dx / d) * this.force;
      p.windY += (dy / d) * this.force;
    }
  }
}
