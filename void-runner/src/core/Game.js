import { World } from '../engine/World.js';
import { FIXED_DT, MAX_FRAME_DT, VIEW_H, VIEW_W_MIN, VIEW_W_MAX } from '../engine/constants.js';
import { getLevel, LEVEL_COUNT } from '../data/levels/index.js';
import { SKIN_BY_ID } from '../data/skins.js';
import { pickQuip, WIN_LINES } from '../data/quips.js';
import { Renderer } from '../render/Renderer.js';
import { Camera } from './Camera.js';
import { Particles } from './Particles.js';
import { makeRng } from '../engine/math.js';

/**
 * GameManager — la boucle et le cycle de vie d'une tentative.
 *
 * Le point le plus important de tout le projet tient en une ligne : entre la
 * mort et la reprise, il s'écoule 320 ms. Pas d'écran de game over, pas de
 * bouton « réessayer », pas de transition. C'est la seule chose qui rende un
 * jeu de mille morts supportable — et le joueur peut même écourter ce délai en
 * appuyant sur saut.
 *
 * La boucle est à pas fixe avec accumulateur : la simulation avance toujours
 * par pas de 1/60 s, quel que soit le taux de rafraîchissement de l'écran
 * (60, 90 ou 120 Hz). Sans cela, un même saut n'aurait pas la même hauteur sur
 * deux téléphones — faute rédhibitoire dans un jeu au pixel près.
 */
export class Game {
  constructor(canvas, deps) {
    this.canvas = canvas;
    this.save = deps.save;
    this.audio = deps.audio;
    this.haptics = deps.haptics;
    this.input = deps.input;
    this.achievements = deps.achievements;
    this.ui = null; // branché par UIManager

    this.renderer = new Renderer(canvas);
    this.camera = new Camera();
    this.particles = new Particles();
    this.view = { w: 640, h: VIEW_H, scale: 1 };

    this.mode = 'menu'; // menu | play | paused | win
    this.world = null;
    this.levelNum = 1;
    this.attempt = 1;
    this.deathsThisLevel = 0;
    this.deathHold = 0;
    this.levelMemory = new Map();
    this.checkpoint = null;
    this.prevTrace = null;
    this.hintShown = new Map();
    this.rng = makeRng(Date.now() & 0xffff);

    this.acc = 0;
    this.last = 0;
    this.raf = null;
    this.fpsSamples = [];

    addEventListener('resize', () => this.resize());
    addEventListener('orientationchange', () => setTimeout(() => this.resize(), 120));
    this.resize();
  }

  get skin() {
    return SKIN_BY_ID[this.save.data.skin] || SKIN_BY_ID.standard;
  }

  /**
   * Mise à l'échelle.
   *
   * La hauteur virtuelle est FIXE (360 px) et la largeur s'adapte au ratio :
   * un téléphone très allongé voit plus large, jamais plus haut. C'est la
   * seule façon d'avoir un level design identique sur tous les appareils —
   * personne ne doit pouvoir « voir le piège » parce qu'il a un meilleur
   * téléphone.
   */
  resize() {
    const dpr = Math.min(devicePixelRatio || 1, 2.5);
    const cw = this.canvas.clientWidth || innerWidth;
    const ch = this.canvas.clientHeight || innerHeight;
    const ratio = cw / ch;
    const vw = Math.round(Math.max(VIEW_W_MIN, Math.min(VIEW_W_MAX, VIEW_H * ratio)));
    this.view.w = vw;
    this.view.h = VIEW_H;
    this.canvas.width = Math.round(cw * dpr);
    this.canvas.height = Math.round(ch * dpr);
    const scale = Math.min(this.canvas.width / vw, this.canvas.height / VIEW_H);
    this.view.scale = scale;
    const offX = Math.round((this.canvas.width - vw * scale) / 2);
    const offY = Math.round((this.canvas.height - VIEW_H * scale) / 2);
    // Débord : la zone de canevas située HORS de la boîte de jeu, exprimée en
    // unités de jeu. Le fond y est peint aussi — sans quoi un écran au ratio
    // inhabituel (portrait, panneau étroit) encadre la salle de deux bandes
    // noires franches, ce qui a l'air d'un bug plutôt que d'un cadrage.
    this.view.padX = offX / scale;
    this.view.padY = offY / scale;
    this.view.fullW = this.canvas.width / scale;
    this.view.fullH = this.canvas.height / scale;
    const ctx = this.renderer.ctx;
    ctx.setTransform(scale, 0, 0, scale, offX, offY);
    ctx.imageSmoothingEnabled = false;
  }

  // --- Cycle de vie d'un niveau ---------------------------------------------

  startLevel(num) {
    this.levelNum = Math.max(1, Math.min(LEVEL_COUNT, num));
    this.attempt = 1;
    this.deathsThisLevel = 0;
    this.levelMemory = new Map();
    this.checkpoint = null;
    this.prevTrace = null;
    this.hintShown = new Map();
    this.mode = 'play';
    this.spawnWorld();
    this.camera.snap(this.world, this.view);
    this.audio.startMusic(this.world.level.chapterId);
    this.audio.duck(false);
    this.ui?.onLevelStart(this.world.level, this.attempt);
    this.start();
  }

  spawnWorld() {
    const level = getLevel(this.levelNum);
    this.world = new World(level, {
      dash: level.dash === true,
      memory: this.levelMemory,
      startAt: this.checkpoint,
      ghostTrace: this.prevTrace,
    });
    this.particles.clear();
    this.input.releaseAll();
    this.secretTaken = false;
  }

  restart() {
    this.attempt++;
    this.prevTrace = this.world?.trace || null;
    this.spawnWorld();
    this.camera.snap(this.world, this.view);
    this.ui?.onAttempt(this.attempt);
  }

  retry() {
    // Recommencer volontairement remet aussi le point de contrôle à zéro.
    this.checkpoint = null;
    this.restart();
  }

  nextLevel() {
    if (this.levelNum >= LEVEL_COUNT) { this.quitToMenu(); this.ui?.showScreen('menu'); return; }
    this.startLevel(this.levelNum + 1);
  }

  pause() {
    if (this.mode !== 'play') return;
    this.mode = 'paused';
    this.input.releaseAll();
    this.audio.duck(true);
    this.ui?.showScreen('pause');
    const rec = this.save.level(this.levelNum);
    this.ui?.setPauseInfo(
      `SALLE ${String(this.levelNum).padStart(2, '0')} — ${this.world.level.name.toUpperCase()}`
      + ` · ESSAI ${this.attempt} · ${rec.deaths} MORTS ICI`,
    );
  }

  resume() {
    if (this.mode !== 'paused') return;
    this.mode = 'play';
    this.audio.duck(false);
    this.last = performance.now();
    this.ui?.showScreen(null);
  }

  quitToMenu() {
    this.mode = 'menu';
    this.stop();
    this.audio.stopMusic();
    this.save.flush();
  }

  // --- Boucle ---------------------------------------------------------------

  start() {
    if (this.raf) return;
    this.last = performance.now();
    this.acc = 0;
    const tick = (ts) => {
      this.raf = requestAnimationFrame(tick);
      this.frame(ts);
    };
    this.raf = requestAnimationFrame(tick);
  }

  stop() {
    if (this.raf) cancelAnimationFrame(this.raf);
    this.raf = null;
  }

  frame(ts) {
    let dt = (ts - this.last) / 1000;
    this.last = ts;
    if (dt > MAX_FRAME_DT) dt = MAX_FRAME_DT; // retour d'arrière-plan : on n'accumule pas
    this.trackFps(dt);

    if (this.mode === 'play') {
      this.input.pollPad();
      this.acc += dt;
      let steps = 0;
      // Le premier pas de la frame reçoit les fronts montants ; les suivants
      // héritent d'un appui « maintenu ». Un tap de 15 ms est donc entendu
      // une fois, et une seule.
      while (this.acc >= FIXED_DT && steps < 6) {
        const inp = steps === 0
          ? this.input.consume()
          : { ...this.lastInput, jumpPressed: false, dashPressed: false };
        this.lastInput = inp;
        this.world.step(inp);
        this.acc -= FIXED_DT;
        steps++;
        this.drainWorld();
        if (this.world.state !== 'run') break;
      }
      if (steps >= 6) this.acc = 0;

      if (this.world.state === 'dead') this.tickDeath(dt);
      this.save.data.playTime += dt;
    }

    if (this.world) {
      this.camera.update(this.world, this.view, dt);
      this.particles.update(dt);
      this.renderer.drawWorld(this.world, this.camera, this.view, this.skin, dt, {
        particles: this.particles,
      });
      this.ui?.onFrame(this.world, this.mode);
    }
  }

  /** Qualité adaptative : si l'appareil peine, on allège le fond avant tout. */
  trackFps(dt) {
    this.fpsSamples.push(dt);
    if (this.fpsSamples.length < 90) return;
    const avg = this.fpsSamples.reduce((a, b) => a + b, 0) / this.fpsSamples.length;
    this.fpsSamples.length = 0;
    if (avg > 0.0235 && this.renderer.quality > 0) this.renderer.quality = 0;
    else if (avg < 0.0180 && this.renderer.quality < 1) this.renderer.quality = 1;
  }

  /** Traduit les évènements de simulation en son, vibration, particules. */
  drainWorld() {
    for (const ev of this.world.drainEvents()) {
      this.particles.handle(ev);
      switch (ev.type) {
        case 'death':
          this.audio.play('death');
          this.haptics.fire('death');
          this.camera.addShake(0.9);
          this.ui?.flash();
          this.onDeath(ev);
          break;
        case 'win':
          this.audio.play('win');
          this.haptics.fire('win');
          this.onWin();
          break;
        case 'jump': this.audio.play('jump'); break;
        case 'land':
          if (ev.power > 0.5) { this.audio.play('land'); this.camera.addShake(ev.power * 0.18); }
          break;
        case 'dash': this.audio.play('dash'); this.haptics.fire('impact'); break;
        case 'click': this.audio.play('click'); this.haptics.fire('mech'); break;
        case 'unclick': this.audio.play('unclick'); break;
        case 'mech': this.audio.play('mech'); break;
        case 'creak': this.audio.play('creak'); break;
        case 'crumble': this.audio.play('crumble'); this.camera.addShake(0.2); break;
        case 'slam': this.audio.play('slam'); this.haptics.fire('impact'); this.camera.addShake(0.5); break;
        case 'alarm': this.audio.play('alarm'); break;
        case 'laser': this.audio.play('laser'); break;
        case 'warp': case 'fakeExit': this.audio.play('warp'); break;
        case 'gravity': this.audio.play('gravity'); this.haptics.fire('mech'); this.camera.addShake(0.3); break;
        case 'checkpoint': this.audio.play('checkpoint'); this.haptics.fire('mech'); break;
        case 'checkpointFake': this.audio.play('checkpoint'); break;
        default: break;
      }
    }
    // Signal de salle secrète : émis par une zone, récompensé ici.
    if (this.world.isOn('secret') && !this.secretTaken) {
      this.secretTaken = true;
      if (this.save.markSecret(this.levelNum)) {
        this.audio.play('secret');
        this.ui?.toast('SALLE SECRÈTE', 'Quelqu\'un a caché ça ici exprès.');
        this.checkUnlocks();
      }
    }
    // Le point de contrôle survit à la mort : c'est tout son intérêt.
    if (this.world.checkpoint) this.checkpoint = this.world.checkpoint;
  }

  onDeath(ev) {
    this.deathsThisLevel++;
    this.save.addDeath(this.levelNum);
    this.deathHold = 0.32;

    const exit = this.world.findAll('exit').find((e) => !e.fake);
    const near = exit
      ? Math.hypot(ev.x - (exit.x + exit.w / 2), ev.y - (exit.y + exit.h / 2)) < 90
      : false;
    // Indice ciblé : certaines salles ont une leçon précise qu'une mort seule
    // ne transmet pas (« pourquoi ce laser s'est-il allumé ? »). On le montre
    // au plus deux fois — après, c'est le joueur qui n'écoute pas, et le lui
    // répéter serait condescendant.
    const hint = this.world.level.deathHints?.[ev.kind];
    const shown = this.hintShown.get(ev.kind) || 0;
    let quip;
    if (hint && shown < 2) {
      this.hintShown.set(ev.kind, shown + 1);
      quip = hint;
    } else {
      quip = pickQuip({
        attempt: this.attempt,
        totalDeaths: this.save.data.totalDeaths,
        nearExit: near,
        rng: this.rng,
      });
    }
    this.ui?.onDeath(this.attempt + 1, quip);
    this.checkUnlocks();
  }

  tickDeath(dt) {
    this.deathHold -= dt;
    // Appuyer sur saut pendant l'explosion relance immédiatement. Le joueur
    // impatient a toujours raison.
    if (this.deathHold <= 0 || this.input.jumpQueued) {
      this.input.jumpQueued = false;
      this.restart();
    }
  }

  onWin() {
    const time = this.world.time;
    const firstClear = this.save.complete(this.levelNum, time, this.deathsThisLevel);
    this.mode = 'win';
    this.audio.duck(true);
    const line = WIN_LINES[Math.floor(this.rng() * WIN_LINES.length)];
    const rec = this.save.level(this.levelNum);
    this.ui?.showWin({
      level: this.world.level,
      time,
      best: rec.best,
      deaths: this.deathsThisLevel,
      firstClear,
      line,
      isLast: this.levelNum >= LEVEL_COUNT,
    });
    this.checkUnlocks();
    this.save.flush();
  }

  checkUnlocks() {
    const news = this.achievements.check(LEVEL_COUNT);
    for (const n of news) {
      this.audio.play('unlock');
      this.ui?.toast(
        n.kind === 'skin' ? `ROBOT DÉBLOQUÉ — ${n.name}` : `ACCOMPLISSEMENT — ${n.name}`,
        n.desc,
      );
    }
  }
}
