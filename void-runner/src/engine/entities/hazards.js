import { Entity } from './base.js';
import { TILE, DEATH } from '../constants.js';

/**
 * Surface dangereuse fixe : sol électrifié, arête chauffée, grille haute
 * tension. Toujours ORANGE — la charte couleur est une promesse faite au
 * joueur : orange = ça tue, rouge = ça va tuer.
 */
export class Zap extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.deadly = true;
    this.deathKind = spec.deathKind || DEATH.ZAP;
    this.dir = spec.dir || 'up'; // orientation des picots (rendu)
    this.style = spec.style || 'zap';
    // Cycle optionnel : [durée allumé, durée éteint]. Apprenable, donc juste.
    this.cycle = spec.cycle || null;
    this.phase = spec.phase ?? 0;
  }

  preStep(world) {
    let on = this.gateOpen(world);
    if (on && this.cycle) {
      const period = this.cycle[0] + this.cycle[1];
      const t = (world.time + this.phase) % period;
      on = t < this.cycle[0];
    }
    this.deadly = on;
    this.alpha = on ? 1 : 0.25;
  }
}

/**
 * Laser.
 *
 * Un émetteur projette un faisceau jusqu'au premier solide rencontré : le
 * faisceau est donc bloqué par une plateforme mobile, ce qui ouvre du level
 * design (« se cacher derrière la caisse »).
 *
 * dir : 'right' | 'left' | 'up' | 'down'
 * cycle : [allumé, éteint] — un laser permanent est indiqué par cycle absent.
 */
export class Laser extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.dir = spec.dir || 'right';
    this.maxLen = (spec.len ?? 20) * TILE;
    this.cycle = spec.cycle || null;
    this.phase = spec.phase ?? 0;
    this.warn = spec.warn ?? 0.35; // pré-signal rouge avant l'allumage
    this.thick = spec.thick ?? 5;
    this.deathKind = DEATH.BURN;
    this.on = false;
    this.charging = 0;
    this.beam = { x: 0, y: 0, w: 0, h: 0 };
    this.w = TILE * 0.5;
    this.h = TILE * 0.5;
    this.deadly = false;
    // Un laser peut n'exister qu'après un déclencheur (niveau 1 : le petit
    // laser n'apparaît que si le joueur saute).
    this.armed = spec.armSignal ? false : true;
    this.armSignal = spec.armSignal || null;
  }

  computeBeam(world) {
    const cx = this.x + this.w / 2;
    const cy = this.y + this.h / 2;
    let len = this.maxLen;
    const horiz = this.dir === 'left' || this.dir === 'right';
    const s = this.dir === 'left' || this.dir === 'up' ? -1 : 1;
    const stepPx = 4;
    for (let d = 0; d < this.maxLen; d += stepPx) {
      const px = horiz ? cx + s * d : cx;
      const py = horiz ? cy : cy + s * d;
      if (world.overlapsSolid(px - 1, py - 1, 2, 2, this)) { len = d; break; }
    }
    if (horiz) {
      this.beam.x = s > 0 ? cx : cx - len;
      this.beam.y = cy - this.thick / 2;
      this.beam.w = len;
      this.beam.h = this.thick;
    } else {
      this.beam.x = cx - this.thick / 2;
      this.beam.y = s > 0 ? cy : cy - len;
      this.beam.w = this.thick;
      this.beam.h = len;
    }
  }

  preStep(world, dt) {
    if (this.armSignal && !this.armed && world.isOn(this.armSignal)) {
      this.armed = true;
      this.chargeT = 0;
      world.fx('charge', { x: this.x, y: this.y });
    }
    if (!this.armed) { this.deadly = false; this.on = false; return; }

    let want = this.gateOpen(world);
    if (want && this.cycle) {
      const period = this.cycle[0] + this.cycle[1];
      const t = (world.time + this.phase) % period;
      want = t < this.cycle[0];
      // Pré-signal : la lueur rouge annonce toujours le tir.
      this.charging = !want && t > period - this.warn ? 1 : 0;
    }
    const wasOn = this.on;
    this.on = want;
    if (this.on && !wasOn) world.fx('laser', { x: this.x, y: this.y });
    this.computeBeam(world);
    this.deadly = this.on;
  }

  hazardBox() {
    return this.on ? this.beam : null;
  }
}
