import { aabb } from '../math.js';
import { DEATH, TILE } from '../constants.js';

/**
 * Entité de base.
 *
 * Toute mécanique du jeu est une entité. Une entité peut être :
 *  - solide (le joueur s'y appuie) ;
 *  - mortelle (`deadly`) ;
 *  - logique (bouton, zone, mémoire — ni solide ni mortelle).
 *
 * Une entité ne dessine rien : le renderer lit ses champs publics. Cela garde
 * le moteur exécutable hors navigateur.
 */
export class Entity {
  constructor(spec, world) {
    this.type = spec.t;
    this.id = spec.id ?? null;
    this.spec = spec;
    // Les niveaux sont écrits en TUILES (lisible à la main), convertis ici en
    // pixels une seule fois, au chargement.
    this.x = (spec.x ?? 0) * TILE;
    this.y = (spec.y ?? 0) * TILE;
    this.w = (spec.w ?? 1) * TILE;
    this.h = (spec.h ?? 1) * TILE;
    this.solid = false;
    this.deadly = false;
    this.active = true;
    this.deathKind = DEATH.BOOM;
    this.vx = 0;
    this.vy = 0;
    // Champs purement cosmétiques lus par le renderer
    this.flash = 0;
    this.alpha = 1;
  }

  /** Le signal (optionnel) qui conditionne l'activité de l'entité. */
  gateOpen(world) {
    const s = this.spec.signal;
    if (!s) return true;
    const on = world.isOn(s);
    return this.spec.invert ? !on : on;
  }

  touchesPlayer(world, pad = 0) {
    const p = world.player;
    return aabb(
      this.x - pad, this.y - pad, this.w + pad * 2, this.h + pad * 2,
      p.x, p.y, p.w, p.h,
    );
  }
}

/**
 * Corps solide mobile.
 *
 * Le déplacement d'un solide est la partie délicate d'un platformer : il doit
 * PORTER le joueur posé dessus et le POUSSER s'il le percute — et provoquer un
 * écrasement franc s'il le coince contre un autre solide. Sans cela, les
 * plateformes mobiles et les plafonds descendants seraient injustes.
 */
export class SolidBody extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.solid = true;
  }

  isRider(p) {
    if (p.dead) return false;
    const g = p.gravDir;
    const xo = p.x < this.x + this.w - 0.5 && p.x + p.w > this.x + 0.5;
    if (!xo) return false;
    if (g > 0) return Math.abs(p.y + p.h - this.y) <= 2.5;
    return Math.abs(p.y - (this.y + this.h)) <= 2.5;
  }

  /** Déplace le solide de (dx, dy) en gérant portage, poussée et écrasement. */
  moveBy(dx, dy, world) {
    if (dx === 0 && dy === 0) return;
    const p = world.player;
    const riding = this.isRider(p);

    this.x += dx;
    this.y += dy;
    this.vx = dx;
    this.vy = dy;

    if (p.dead || world.state !== 'run') return;

    if (aabb(this.x, this.y, this.w, this.h, p.x, p.y, p.w, p.h)) {
      // Poussée : on éjecte le joueur dans le sens du mouvement.
      if (dy !== 0) {
        p.y = dy > 0 ? this.y + this.h : this.y - p.h;
        if (dy > 0 && p.vy < 0) p.vy = 0;
        if (dy < 0) p.vy = Math.min(p.vy, dy / 0.0166);
      }
      if (dx !== 0) {
        p.x = dx > 0 ? this.x + this.w : this.x - p.w;
      }
      if (world.overlapsSolid(p.x + 1.5, p.y + 1.5, p.w - 3, p.h - 3, this)) {
        world.kill(DEATH.CRUSH, this);
        return;
      }
    } else if (riding) {
      // Portage : le joueur suit exactement la plateforme.
      p.x += dx;
      p.y += dy;
      if (world.overlapsSolid(p.x + 1.5, p.y + 1.5, p.w - 3, p.h - 3, this)) {
        world.kill(DEATH.CRUSH, this);
      }
    }
  }
}
