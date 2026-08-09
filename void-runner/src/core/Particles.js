import { C } from '../render/theme.js';

/**
 * Système de particules à pool fixe.
 *
 * Aucune allocation pendant le jeu : 400 particules sont créées une fois pour
 * toutes et recyclées. Sur mobile, le ramasse-miettes est la première cause de
 * micro-saccades — et une saccade dans un jeu au pixel près, c'est une mort
 * imméritée.
 */
const MAX = 400;

export class Particles {
  constructor() {
    this.p = new Array(MAX);
    for (let i = 0; i < MAX; i++) {
      this.p[i] = { on: false, x: 0, y: 0, vx: 0, vy: 0, life: 0, max: 1, size: 2, col: '#fff', grav: 0, fade: 1 };
    }
    this.head = 0;
    this.seed = 987;
  }

  rnd() {
    this.seed = (this.seed * 1664525 + 1013904223) & 0x7fffffff;
    return this.seed / 0x7fffffff;
  }

  spawn(x, y, opts) {
    const q = this.p[this.head];
    this.head = (this.head + 1) % MAX;
    q.on = true;
    q.x = x; q.y = y;
    q.vx = opts.vx ?? 0;
    q.vy = opts.vy ?? 0;
    q.life = q.max = opts.life ?? 0.5;
    q.size = opts.size ?? 2;
    q.col = opts.col ?? C.cyan;
    q.grav = opts.grav ?? 0;
    q.fade = opts.fade ?? 1;
  }

  burst(x, y, n, opts = {}) {
    for (let i = 0; i < n; i++) {
      const a = this.rnd() * Math.PI * 2;
      const s = (opts.speed ?? 90) * (0.35 + this.rnd() * 0.65);
      this.spawn(x, y, {
        vx: Math.cos(a) * s,
        vy: Math.sin(a) * s - (opts.lift ?? 0),
        life: (opts.life ?? 0.5) * (0.6 + this.rnd() * 0.8),
        size: opts.size ?? 2.4,
        col: opts.col ?? C.cyan,
        grav: opts.grav ?? 420,
      });
    }
  }

  /** Réaction visuelle aux évènements du monde. Un effet, un sens. */
  handle(ev) {
    switch (ev.type) {
      case 'death':
        this.burst(ev.x, ev.y, 26, { speed: 200, col: C.white, size: 2.8, life: 0.55 });
        this.burst(ev.x, ev.y, 16, {
          speed: 150, life: 0.7, size: 3.2,
          col: ev.kind === 'burn' ? C.danger : ev.kind === 'zap' ? C.gold : C.violet,
        });
        break;
      case 'win':
        this.burst(ev.x, ev.y, 30, { speed: 130, col: C.cyan, size: 2.6, life: 0.9, grav: 120 });
        break;
      case 'land':
        if (ev.power > 0.18) {
          for (let i = 0; i < 5; i++) {
            this.spawn(ev.x + (this.rnd() - 0.5) * 14, ev.y, {
              vx: (this.rnd() - 0.5) * 90, vy: -20 - this.rnd() * 30,
              life: 0.28, size: 1.8, col: 'rgba(180,200,255,0.9)', grav: 300,
            });
          }
        }
        break;
      case 'jump':
        for (let i = 0; i < 3; i++) {
          this.spawn(ev.x + (this.rnd() - 0.5) * 10, ev.y, {
            vx: (this.rnd() - 0.5) * 50, vy: 20, life: 0.22, size: 1.6,
            col: 'rgba(160,190,255,0.8)', grav: 120,
          });
        }
        break;
      case 'dash':
        for (let i = 0; i < 8; i++) {
          this.spawn(ev.x, ev.y + (this.rnd() - 0.5) * 14, {
            vx: -ev.dir * (80 + this.rnd() * 140), vy: (this.rnd() - 0.5) * 40,
            life: 0.3, size: 2.2, col: C.cyan, grav: 0,
          });
        }
        break;
      case 'crumble':
        for (let i = 0; i < 10; i++) {
          this.spawn(ev.x + (this.rnd() - 0.5) * (ev.w || 24), ev.y, {
            vx: (this.rnd() - 0.5) * 60, vy: 10, life: 0.8, size: 2.6,
            col: '#3a3f7a', grav: 620,
          });
        }
        break;
      case 'vanish':
        this.burst(ev.x, ev.y, 8, { speed: 60, col: C.cyan, size: 2, life: 0.4, grav: 0 });
        break;
      case 'warp':
      case 'fakeExit':
        this.burst(ev.x, ev.y, 18, { speed: 160, col: C.violet, size: 2.6, life: 0.5, grav: 0 });
        break;
      case 'gravity':
        this.burst(ev.x, ev.y, 14, { speed: 110, col: '#a378ff', size: 2.4, life: 0.5, grav: 0 });
        break;
      case 'checkpoint':
        this.burst(ev.x, ev.y, 14, { speed: 90, col: C.cyan, size: 2.2, life: 0.6, grav: -60 });
        break;
      case 'slam':
        this.burst(ev.x, ev.y, 12, { speed: 130, col: '#6a5a90', size: 2.4, life: 0.4, grav: 500 });
        break;
      default: break;
    }
  }

  update(dt) {
    for (const q of this.p) {
      if (!q.on) continue;
      q.life -= dt;
      if (q.life <= 0) { q.on = false; continue; }
      q.vy += q.grav * dt;
      q.x += q.vx * dt;
      q.y += q.vy * dt;
    }
  }

  draw(ctx) {
    for (const q of this.p) {
      if (!q.on) continue;
      const a = Math.max(0, q.life / q.max) ** q.fade;
      ctx.globalAlpha = a;
      ctx.fillStyle = q.col;
      const s = q.size * (0.5 + a * 0.5);
      ctx.fillRect(q.x - s / 2, q.y - s / 2, s, s);
    }
    ctx.globalAlpha = 1;
  }

  clear() {
    for (const q of this.p) q.on = false;
  }
}
