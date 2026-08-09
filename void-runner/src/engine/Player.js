import {
  PLAYER_W, PLAYER_H, RUN_SPEED, ACCEL_GROUND, ACCEL_AIR, FRICTION_GROUND,
  FRICTION_AIR, GRAVITY, JUMP_VEL, JUMP_CUT, MAX_FALL, COYOTE_TIME,
  JUMP_BUFFER, DASH_SPEED, DASH_TIME, DASH_ENDLAG, CORNER_CORRECT, DEATH,
} from './constants.js';
import { approach, sign } from './math.js';

/**
 * Le robot.
 *
 * Tout ce qui rend un platformer agréable est ici, et rien d'autre :
 *  - accélération/friction séparées sol et air ;
 *  - saut à hauteur variable (relâcher = saut court) ;
 *  - coyote time (saut toléré juste après le vide) ;
 *  - jump buffer (saut mémorisé juste avant l'atterrissage) ;
 *  - correction de coin (on n'accroche jamais un angle de plateforme).
 *
 * Ces quatre derniers points ne sont pas du confort : sans eux, le joueur
 * accuse les contrôles au lieu d'accuser sa décision. Or c'est la règle
 * absolue du jeu — difficile, mais juste.
 */
export class Player {
  constructor(x, y, world) {
    this.world = world;
    this.w = PLAYER_W;
    this.h = PLAYER_H;
    this.x = x;
    this.y = y;
    this.vx = 0;
    this.vy = 0;

    this.gravDir = 1; // +1 = gravité vers le bas, -1 = inversée
    this.grounded = false;
    this.groundEnt = null;
    this.facing = 1;
    this.dead = false;

    this.coyote = 0;
    this.buffer = 0;
    this.jumpHeld = false;
    this.jumping = false;

    this.canDash = world.opts?.dash === true;
    this.dashLeft = 0; // temps de dash restant
    this.dashReady = true;
    this.endlag = 0;

    // Modificateurs appliqués par les champs (ventilateurs, aimants, zones)
    this.windX = 0;
    this.windY = 0;
    this.speedMul = 1;
    this.frictionMul = 1;
    this.noControl = 0; // secondes de perte de contrôle (téléporteur, choc)

    this.anim = 'idle';
    this.animT = 0;
    this.squash = 0; // -1 aplati, +1 étiré (feedback visuel)
    this.wasGrounded = false;
    this.landImpact = 0;
  }

  get cx() { return this.x + this.w / 2; }
  get cy() { return this.y + this.h / 2; }
  get bottom() { return this.y + this.h; }
  get right() { return this.x + this.w; }

  /** Le sol se trouve dans la direction de la gravité. */
  probeGround() {
    const g = this.gravDir;
    const s = this.world.overlapsSolid(this.x + 1, this.y + g, this.w - 2, this.h);
    return s;
  }

  step(input, dt) {
    const W = this.world;
    const g = this.gravDir;

    this.animT += dt;
    if (this.noControl > 0) this.noControl -= dt;
    const controlled = this.noControl <= 0;

    // --- Entrées ------------------------------------------------------------
    let dir = 0;
    if (controlled) dir = (input.right ? 1 : 0) - (input.left ? 1 : 0);
    if (dir !== 0) this.facing = dir;

    if (input.jumpPressed) this.buffer = JUMP_BUFFER;
    else this.buffer = Math.max(0, this.buffer - dt);

    // Relâcher le saut coupe l'élan : c'est ce qui permet le "petit saut"
    // exigé par le niveau 1 (sauter, mais pas trop haut).
    if (this.jumpHeld && !input.jump && this.jumping) {
      if (this.vy * g < 0) this.vy *= JUMP_CUT;
      this.jumping = false;
    }
    this.jumpHeld = input.jump;

    // --- Dash ---------------------------------------------------------------
    if (this.dashLeft > 0) {
      this.dashLeft -= dt;
      this.vx = DASH_SPEED * this.facing;
      this.vy = 0;
      if (this.dashLeft <= 0) {
        this.vx *= 0.55;
        this.endlag = DASH_ENDLAG;
      }
    } else if (
      controlled && this.canDash && input.dashPressed && this.dashReady
    ) {
      this.dashLeft = DASH_TIME;
      this.dashReady = false;
      this.vy = 0;
      W.fx('dash', { x: this.cx, y: this.cy, dir: this.facing });
    }

    const dashing = this.dashLeft > 0;
    if (this.endlag > 0) this.endlag -= dt;

    // --- Déplacement horizontal --------------------------------------------
    if (!dashing) {
      const maxSpeed = RUN_SPEED * this.speedMul;
      if (dir !== 0) {
        const a = (this.grounded ? ACCEL_GROUND : ACCEL_AIR) * this.speedMul;
        // Demi-tour plus vif que l'accélération pure : réactivité.
        const turn = sign(this.vx) !== 0 && sign(this.vx) !== dir ? 1.8 : 1;
        this.vx = approach(this.vx, maxSpeed * dir, a * turn * dt);
      } else {
        const f = (this.grounded ? FRICTION_GROUND : FRICTION_AIR) * this.frictionMul;
        this.vx = approach(this.vx, 0, f * dt);
      }
      this.vx += this.windX * dt;
    }

    // --- Saut ---------------------------------------------------------------
    if (this.grounded) this.coyote = COYOTE_TIME;
    else this.coyote = Math.max(0, this.coyote - dt);

    if (this.buffer > 0 && this.coyote > 0 && !dashing) {
      this.vy = -JUMP_VEL * g;
      this.buffer = 0;
      this.coyote = 0;
      this.jumping = true;
      this.grounded = false;
      this.squash = 0.6;
      W.fx('jump', { x: this.cx, y: this.bottom });
    }

    // --- Gravité ------------------------------------------------------------
    if (!dashing) {
      this.vy += (GRAVITY * g + this.windY) * dt;
      if (this.vy * g > MAX_FALL) this.vy = MAX_FALL * g;
    }

    // --- Intégration + collisions ------------------------------------------
    this.moveX(this.vx * dt);
    this.moveY(this.vy * dt);

    // --- Sol ----------------------------------------------------------------
    const ground = this.probeGround();
    this.wasGrounded = this.grounded;
    this.grounded = !!ground && this.vy * g >= -0.001;
    this.groundEnt = this.grounded ? ground : null;
    if (this.grounded) {
      this.dashReady = true;
      this.jumping = false;
      if (!this.wasGrounded) {
        const impact = Math.min(1, Math.abs(this.vy) / MAX_FALL);
        this.landImpact = impact;
        this.squash = -0.55 * impact - 0.15;
        W.fx('land', { x: this.cx, y: this.bottom, power: impact });
      }
      ground.onStand?.(W, this);
    }

    // --- Écrasement ---------------------------------------------------------
    if (W.overlapsSolid(this.x + 2, this.y + 2, this.w - 4, this.h - 4)) {
      W.kill(DEATH.CRUSH);
    }

    // --- Animation ----------------------------------------------------------
    this.squash = approach(this.squash, 0, dt * 4.5);
    if (dashing) this.anim = 'dash';
    else if (!this.grounded) this.anim = this.vy * g < 0 ? 'jump' : 'fall';
    else if (Math.abs(this.vx) > 12) this.anim = 'run';
    else this.anim = 'idle';

    // Les modificateurs de champ sont réinitialisés : chaque zone les repose.
    this.windX = 0;
    this.windY = 0;
    this.speedMul = 1;
    this.frictionMul = 1;
  }

  /** Déplacement horizontal, sous-échantillonné pour ne jamais traverser. */
  moveX(d) {
    const steps = Math.max(1, Math.ceil(Math.abs(d) / 6));
    const inc = d / steps;
    for (let i = 0; i < steps; i++) {
      const nx = this.x + inc;
      const hit = this.world.overlapsSolid(nx, this.y, this.w, this.h);
      if (hit) {
        this.x = inc > 0 ? hit.x - this.w : hit.x + hit.w;
        this.vx = 0;
        hit.onSideHit?.(this.world, this);
        return;
      }
      this.x = nx;
    }
  }

  moveY(d) {
    const steps = Math.max(1, Math.ceil(Math.abs(d) / 6));
    const inc = d / steps;
    for (let i = 0; i < steps; i++) {
      const ny = this.y + inc;
      const hit = this.world.overlapsSolid(this.x, ny, this.w, this.h);
      if (hit) {
        // Correction de coin : si seul un petit bout de la tête accroche,
        // on décale latéralement plutôt que de stopper net le saut.
        if (inc * this.gravDir < 0) {
          const over = this.correctCorner(hit, ny);
          if (over) continue;
        }
        this.y = inc > 0 ? hit.y - this.h : hit.y + hit.h;
        this.vy = 0;
        hit.onHit?.(this.world, this);
        return;
      }
      this.y = ny;
    }
  }

  correctCorner(hit, ny) {
    for (const s of [1, -1]) {
      for (let px = 1; px <= CORNER_CORRECT; px++) {
        const tx = this.x + px * s;
        if (!this.world.overlapsSolid(tx, ny, this.w, this.h)) {
          this.x = tx;
          this.y = ny;
          return true;
        }
      }
    }
    return false;
  }

  /** Téléportation propre (téléporteurs, sortie-piège, inversion de gravité). */
  warp(x, y, keepVel = false) {
    this.x = x - this.w / 2;
    this.y = y - this.h / 2;
    if (!keepVel) { this.vx = 0; this.vy = 0; }
    this.noControl = 0.08;
    this.world.fx('warp', { x, y });
  }
}
