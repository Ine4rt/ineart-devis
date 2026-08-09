import { Entity, SolidBody } from './base.js';
import { TILE, DEATH } from '../constants.js';
import { clamp } from '../math.js';

/** Bloc fixe. La brique de base de toute salle. */
export class Solid extends SolidBody {
  constructor(spec, world) {
    super(spec, world);
    this.style = spec.style || 'wall'; // wall | metal | glass | pipe
  }
}

/**
 * Sol qui s'effondre.
 * Le joueur pose le pied, un grondement, puis plus rien.
 * `delay` doit rester généreux (>= 0.25 s) : le piège doit surprendre, pas
 * être impossible à quitter.
 */
export class Crumble extends SolidBody {
  constructor(spec, world) {
    super(spec, world);
    this.delay = spec.delay ?? 0.4;
    this.respawn = spec.respawn ?? 2.2;
    this.t = -1; // -1 = intact
    this.gone = 0;
    this.style = 'crumble';
    this.shake = 0;
  }

  onStand(world) {
    if (this.t < 0) {
      this.t = this.delay;
      world.fx('creak', { x: this.x + this.w / 2, y: this.y });
    }
  }

  preStep(world, dt) {
    if (this.t > 0) {
      this.t -= dt;
      this.shake = clamp(1 - this.t / this.delay, 0, 1);
      if (this.t <= 0) {
        this.active = false;
        this.gone = this.respawn;
        world.fx('crumble', { x: this.x + this.w / 2, y: this.y + this.h / 2, w: this.w });
      }
    } else if (!this.active) {
      this.gone -= dt;
      if (this.respawn > 0 && this.gone <= 0) {
        // On ne réapparaît jamais dans le joueur : ce serait une mort injuste.
        const p = world.player;
        const inside = p.x < this.x + this.w && p.x + p.w > this.x &&
          p.y < this.y + this.h && p.y + p.h > this.y;
        if (!inside) {
          this.active = true;
          this.t = -1;
          this.shake = 0;
        }
      }
    }
  }
}

/**
 * Plateforme fantôme : elle disparaît selon un déclencheur.
 *  - onLand  : elle s'évapore dès que le robot se pose dessus ;
 *  - onLeave : elle disparaît quand il la quitte (piège au retour) ;
 *  - cycle   : elle clignote selon une période fixe (donc apprenable).
 */
export class Vanish extends SolidBody {
  constructor(spec, world) {
    super(spec, world);
    this.mode = spec.mode || 'onLand';
    this.delay = spec.delay ?? 0.12;
    this.respawn = spec.respawn ?? 1.6;
    this.period = spec.period ?? 1.6;
    this.phase = spec.phase ?? 0;
    this.duty = spec.duty ?? 0.5;
    this.t = -1;
    this.gone = 0;
    this.wasStood = false;
    this.style = 'vanish';
  }

  onStand() {
    if (this.mode === 'onLand' && this.t < 0) this.t = this.delay;
    this.wasStood = true;
  }

  preStep(world, dt) {
    if (this.mode === 'cycle') {
      const tt = (world.time + this.phase) % this.period;
      const on = tt < this.period * this.duty;
      this.active = on;
      this.alpha = on ? 1 : 0.18;
      return;
    }
    if (this.mode === 'onLeave') {
      const still = world.player.groundEnt === this;
      if (this.wasStood && !still && this.t < 0) this.t = this.delay;
      this.wasStood = still || this.wasStood;
    }
    if (this.t > 0) {
      this.t -= dt;
      this.alpha = 0.35 + 0.65 * (this.t / Math.max(0.001, this.delay));
      if (this.t <= 0) {
        this.active = false;
        this.alpha = 0;
        this.gone = this.respawn;
        world.fx('vanish', { x: this.x + this.w / 2, y: this.y + this.h / 2 });
      }
    } else if (!this.active && this.respawn > 0) {
      this.gone -= dt;
      if (this.gone <= 0) {
        const p = world.player;
        const inside = p.x < this.x + this.w && p.x + p.w > this.x &&
          p.y < this.y + this.h && p.y + p.h > this.y;
        if (!inside) {
          this.active = true;
          this.alpha = 1;
          this.t = -1;
          this.wasStood = false;
        }
      }
    }
  }
}

/**
 * MIRAGE — l'exact contraire de la plateforme fantôme.
 *
 * Dessiné rigoureusement comme un sol plein : même dégradé, même arête
 * lumineuse, aucune nuance. Sauf qu'il n'existe pas. On court dessus, on passe
 * au travers.
 *
 * La règle de justice du jeu tient dans une seule ligne de ce fichier : dès
 * que le robot l'a traversé une fois, le mirage est marqué en mémoire de
 * salle, et le rendu le trahit d'un léger scintillement pour toutes les
 * tentatives suivantes. On ne meurt donc qu'UNE fois par mirage — après, c'est
 * de l'observation, plus de la malchance.
 */
export class Mirage extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.solid = false;
    this.style = 'mirage';
    // Apparence empruntée au sol voisin : 'wall' (défaut) ou 'crumble'.
    this.look = spec.look || 'wall';
    this.key = `mirage:${spec.x},${spec.y}`;
    this.seen = false;
  }

  init(world) {
    this.seen = (world.memory.get(this.key) || 0) > 0;
  }

  postStep(world) {
    if (this.seen || !this.touchesPlayer(world)) return;
    this.seen = true;
    world.memory.set(this.key, 1);
    world.fx('mirage', { x: this.x + this.w / 2, y: this.y });
  }
}

/**
 * Plateforme invisible.
 * Solide en permanence, mais dessinée seulement à proximité (ou jamais).
 * Piège honnête : le joueur voit toujours de fines particules trahir sa
 * présence — il ne peut donc pas dire que rien ne l'indiquait.
 */
export class Ghost extends SolidBody {
  constructor(spec, world) {
    super(spec, world);
    this.reveal = spec.reveal || 'near'; // near | signal | never
    this.style = 'ghost';
    this.vis = 0;
  }

  preStep(world, dt) {
    let target = 0;
    if (this.reveal === 'near') {
      const p = world.player;
      const dx = Math.max(this.x - (p.x + p.w), p.x - (this.x + this.w), 0);
      const dy = Math.max(this.y - (p.y + p.h), p.y - (this.y + this.h), 0);
      target = Math.hypot(dx, dy) < TILE * 2.5 ? 1 : 0;
    } else if (this.reveal === 'signal') {
      target = this.gateOpen(world) ? 1 : 0;
    }
    this.vis += (target - this.vis) * Math.min(1, dt * 12);
  }
}

/**
 * Plateforme mobile / scie / ascenseur / mur qui avance.
 *
 * Un seul composant couvre tout ça, parce que c'est exactement la même chose :
 * un solide qui suit une trajectoire. Ce qui change, c'est le déclencheur et le
 * caractère mortel.
 *
 * trigger : 'auto' | 'signal' | 'stand' | 'near' | 'once'
 * mode    : 'loop' | 'pingpong' | 'once'
 */
export class Mover extends SolidBody {
  constructor(spec, world) {
    super(spec, world);
    this.speed = spec.speed ?? 60;
    this.mode = spec.mode || 'pingpong';
    this.trigger = spec.trigger || 'auto';
    this.wait = spec.wait ?? 0;
    this.deadly = spec.deadly === true;
    this.deathKind = spec.deathKind || (spec.deadly ? DEATH.CRUSH : DEATH.BOOM);
    this.style = spec.style || (spec.deadly ? 'saw' : 'platform');
    this.nearRange = (spec.range ?? 4) * TILE;

    const pts = [[this.x, this.y]];
    for (const p of spec.path ?? []) pts.push([p[0] * TILE, p[1] * TILE]);
    this.pts = pts;
    this.i = 0;
    this.dir = 1;
    this.t = 0; // 0..1 le long du segment courant
    this.waitT = spec.delay ?? 0;
    this.running = this.trigger === 'auto';
    this.done = false;
    // Décalage de phase : permet de désynchroniser deux plateformes jumelles
    if (spec.phase) this.advance(spec.phase * this.speed, null);
  }

  onStand(world) {
    if (this.trigger === 'stand' && !this.running) {
      this.running = true;
      world.fx('mech', { x: this.x + this.w / 2, y: this.y });
    }
  }

  preStep(world, dt) {
    if (this.trigger === 'signal') {
      this.running = this.gateOpen(world);
    } else if (this.trigger === 'near' && !this.running) {
      const p = world.player;
      const dx = Math.abs(p.x + p.w / 2 - (this.x + this.w / 2));
      const dy = Math.abs(p.y + p.h / 2 - (this.y + this.h / 2));
      if (dx < this.nearRange && dy < this.nearRange * 1.5) {
        this.running = true;
        world.fx('mech', { x: this.x + this.w / 2, y: this.y });
      }
    }
    if (!this.running || this.done || this.pts.length < 2) return;

    if (this.waitT > 0) { this.waitT -= dt; return; }
    this.advance(this.speed * dt, world);
  }

  advance(dist, world) {
    let remain = dist;
    let guard = 0;
    while (remain > 0 && guard++ < 8) {
      const a = this.pts[this.i];
      const b = this.pts[this.i + this.dir];
      if (!b) { this.done = true; return; }
      const segLen = Math.hypot(b[0] - a[0], b[1] - a[1]) || 1;
      const need = (1 - this.t) * segLen;
      if (remain < need) {
        this.t += remain / segLen;
        remain = 0;
      } else {
        remain -= need;
        this.t = 0;
        this.i += this.dir;
        // Fin de trajectoire : on boucle, on rebrousse ou on s'arrête.
        if (this.dir > 0 && this.i >= this.pts.length - 1) {
          if (this.mode === 'pingpong') { this.dir = -1; this.waitT = this.wait; }
          else if (this.mode === 'loop') { this.i = 0; this.waitT = this.wait; }
          else { this.done = true; this.i = this.pts.length - 1; break; }
        } else if (this.dir < 0 && this.i <= 0) {
          if (this.mode === 'pingpong') { this.dir = 1; this.waitT = this.wait; }
          else { this.done = true; break; }
        }
        if (this.waitT > 0) break;
      }
    }
    const a = this.pts[this.i];
    const b = this.pts[this.i + this.dir] || a;
    const nx = a[0] + (b[0] - a[0]) * this.t;
    const ny = a[1] + (b[1] - a[1]) * this.t;
    if (world) this.moveBy(nx - this.x, ny - this.y, world);
    else { this.x = nx; this.y = ny; }
  }
}

/**
 * Porte coulissante.
 * Elle glisse vraiment (elle n'est pas juste « désactivée ») : une porte qui se
 * referme sur le robot doit l'écraser pour de bon. C'est tout le niveau 3.
 */
export class Door extends SolidBody {
  constructor(spec, world) {
    super(spec, world);
    this.speed = spec.speed ?? 260;
    this.slide = spec.slide || 'up';
    this.travel = (spec.travel ?? spec.h ?? 1) * TILE;
    this.openAmt = spec.startOpen ? 1 : 0;
    this.style = 'door';
    this.baseX = this.x;
    this.baseY = this.y;
    this.deathKind = DEATH.CRUSH;
    this.applyOffset(null);
  }

  offset() {
    const d = this.openAmt * this.travel;
    switch (this.slide) {
      case 'down': return [0, d];
      case 'left': return [-d, 0];
      case 'right': return [d, 0];
      default: return [0, -d];
    }
  }

  applyOffset(world) {
    const [ox, oy] = this.offset();
    const tx = this.baseX + ox;
    const ty = this.baseY + oy;
    if (world) this.moveBy(tx - this.x, ty - this.y, world);
    else { this.x = tx; this.y = ty; }
  }

  preStep(world, dt) {
    const target = this.gateOpen(world) ? 1 : 0;
    if (this.openAmt !== target) {
      const step = (this.speed / this.travel) * dt;
      this.openAmt += Math.sign(target - this.openAmt) * step;
      this.openAmt = clamp(this.openAmt, 0, 1);
      this.applyOffset(world);
      this.flash = 0.4;
    }
  }
}

/**
 * Presse : plafond, mur ou piston qui fonce sur une position cible.
 * `hold` = temps d'attente en bout de course, `back` = vitesse de retour
 * (0 = ne revient pas).
 */
export class Crusher extends SolidBody {
  constructor(spec, world) {
    super(spec, world);
    this.speed = spec.speed ?? 420;
    this.back = spec.back ?? 90;
    this.hold = spec.hold ?? 0.35;
    this.trigger = spec.trigger || 'near';
    this.range = (spec.range ?? 3) * TILE;
    this.axis = spec.axis || 'y';
    this.to = (spec.to ?? 3) * TILE; // déplacement relatif, en tuiles
    this.baseX = this.x;
    this.baseY = this.y;
    this.p = 0; // progression 0..1
    this.phase = 'idle'; // idle | strike | hold | back
    this.holdT = 0;
    this.deadly = false; // c'est l'écrasement (moveBy) qui tue, pas le contact
    this.style = 'crusher';
    this.warn = 0;
  }

  preStep(world, dt) {
    const p = world.player;
    if (this.phase === 'idle') {
      let fire = false;
      if (this.trigger === 'signal') fire = this.gateOpen(world);
      else if (this.trigger === 'near') {
        const cx = this.baseX + this.w / 2;
        const cy = this.baseY + this.h / 2;
        fire = this.axis === 'y'
          ? Math.abs(p.x + p.w / 2 - cx) < this.range
          : Math.abs(p.y + p.h / 2 - cy) < this.range;
      }
      if (fire) {
        this.phase = 'strike';
        world.fx('alarm', { x: this.baseX + this.w / 2, y: this.baseY });
      }
    }

    if (this.phase === 'strike') {
      this.warn = Math.min(1, this.warn + dt * 6);
      this.p = Math.min(1, this.p + (this.speed / Math.abs(this.to)) * dt);
      if (this.p >= 1) { this.phase = 'hold'; this.holdT = this.hold; world.fx('slam', { x: this.baseX + this.w / 2, y: this.baseY }); }
    } else if (this.phase === 'hold') {
      this.holdT -= dt;
      if (this.holdT <= 0) this.phase = this.back > 0 ? 'back' : 'stuck';
    } else if (this.phase === 'back') {
      this.p = Math.max(0, this.p - (this.back / Math.abs(this.to)) * dt);
      if (this.p <= 0) { this.phase = 'idle'; this.warn = 0; }
    }

    const d = this.p * this.to;
    const tx = this.axis === 'x' ? this.baseX + d : this.baseX;
    const ty = this.axis === 'y' ? this.baseY + d : this.baseY;
    this.moveBy(tx - this.x, ty - this.y, world);
  }
}
