import { Entity } from './base.js';
import { TILE, DEATH } from '../constants.js';
import { aabb } from '../math.js';

/**
 * Porte de sortie.
 *
 * Elle peut mentir, et c'est tout l'intérêt :
 *  - fake:'flee' — elle recule quand on approche (il existe un autre chemin) ;
 *  - fake:'warp' — elle téléporte le robot au départ (fausse sortie) ;
 *  - fake:'trap' — elle se referme en piège mortel ;
 *  - requires    — elle est verrouillée tant qu'un signal n'est pas actif.
 *
 * Une sortie menteuse doit TOUJOURS être identifiable après coup : une vraie
 * sortie a un halo cyan stable, une fausse a un halo qui vacille très
 * légèrement. Le joueur ne peut pas le voir la première fois, mais il le
 * revoit immédiatement à la deuxième tentative. C'est la limite entre une
 * bonne farce et une mort arbitraire.
 */
export class Exit extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.w = (spec.w ?? 1.4) * TILE;
    this.h = (spec.h ?? 2) * TILE;
    this.fake = spec.fake || false;
    this.to = spec.to || null;
    this.requires = spec.requires || null;
    this.range = (spec.range ?? 3) * TILE;
    this.fled = 0;
    this.style = 'exit';
    this.open = 1;
    this.used = false;
  }

  preStep(world, dt) {
    if (this.requires) {
      const ok = world.isOn(this.requires);
      this.open += ((ok ? 1 : 0) - this.open) * Math.min(1, dt * 6);
    }
    if (this.fake === 'flee' && this.to) {
      const p = world.player;
      const d = Math.hypot(p.cx - (this.x + this.w / 2), p.cy - (this.y + this.h / 2));
      // Une fois la fuite amorcée, elle va jusqu'au bout : sinon la sortie
      // s'arrêterait en plein vide dès que le joueur cesse de la suivre.
      if (d < this.range) this.fleeing = true;
      if (this.fleeing) this.fled = Math.min(1, this.fled + dt * 1.9);
      const tx = (this.spec.x * TILE) + (this.to[0] * TILE - this.spec.x * TILE) * this.fled;
      const ty = (this.spec.y * TILE) + (this.to[1] * TILE - this.spec.y * TILE) * this.fled;
      if (this.fled > 0 && (tx !== this.x || ty !== this.y)) {
        if (this.fled < 0.02) world.fx('mech', { x: this.x, y: this.y });
        this.x = tx;
        this.y = ty;
      }
    }
  }

  postStep(world) {
    if (this.used || !this.touchesPlayer(world, -2)) return;
    if (this.requires && !world.isOn(this.requires)) {
      this.flash = 1;
      return;
    }
    if (this.fake === 'warp') {
      const sp = this.to || [world.level.spawn.x, world.level.spawn.y];
      world.player.warp(sp[0] * TILE + world.player.w / 2, sp[1] * TILE + world.player.h / 2);
      world.setSignal(this.spec.emits || '__fakeExit', true);
      world.fx('fakeExit', { x: this.x + this.w / 2, y: this.y + this.h / 2 });
      return;
    }
    if (this.fake === 'trap') {
      world.kill(DEATH.BURN, this);
      return;
    }
    if (this.fake === 'flee' && this.fled < 0.98) return;
    world.win();
  }
}

/**
 * Bouton / plaque de pression.
 * mode : 'hold' (actif tant qu'on appuie) | 'toggle' | 'once'
 * Un bouton n'est jamais neutre dans ce jeu : il fait toujours QUELQUE CHOSE.
 */
export class Button extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.h = (spec.h ?? 0.4) * TILE;
    this.mode = spec.mode || 'once';
    this.out = spec.emits || spec.out || 'a';
    this.value = spec.value !== false;
    this.pressed = false;
    this.latched = false;
    this.style = 'button';
    this.press = 0;
    // Un bouton temporisé se réarme tout seul : c'est ce qui empêche un
    // niveau à séquence de devenir insoluble après une erreur d'ordre.
    this.autoOff = spec.autoOff ?? 0;
    this.offT = 0;
  }

  preStep(world, dt) {
    if (this.autoOff > 0 && this.offT > 0) {
      this.offT -= dt;
      if (this.offT <= 0) {
        world.setSignal(this.out, !this.value);
        this.latched = false;
        world.fx('unclick', { x: this.x + this.w / 2, y: this.y });
      }
    }
  }

  postStep(world) {
    const p = world.player;
    const touching = aabb(this.x, this.y - 3, this.w, this.h + 6, p.x, p.y, p.w, p.h);
    const rising = touching && !this.pressed;
    this.pressed = touching;
    this.press += ((touching ? 1 : 0) - this.press) * 0.35;

    if (this.mode === 'hold') {
      world.setSignal(this.out, touching ? this.value : !this.value);
    } else if (rising) {
      if (this.mode === 'toggle') world.setSignal(this.out, !world.isOn(this.out));
      else if (!this.latched) { world.setSignal(this.out, this.value); this.latched = true; }
      if (this.autoOff > 0) this.offT = this.autoOff;
      world.fx('click', { x: this.x + this.w / 2, y: this.y });
    }
  }
}

/**
 * Zone de déclenchement invisible.
 * mode : 'stay' | 'enter' | 'exit' | 'once'
 * `remember` inscrit une clé dans la mémoire persistante du niveau (voir
 * MemGate) : c'est la base des mécaniques qui évoluent entre deux tentatives.
 */
export class Zone extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.mode = spec.mode || 'once';
    this.out = spec.emits || spec.out || null;
    this.value = spec.value !== false;
    this.inside = false;
    this.latched = false;
    this.remember = spec.remember || null;
  }

  postStep(world) {
    const now = this.touchesPlayer(world);
    const enter = now && !this.inside;
    const exit = !now && this.inside;
    this.inside = now;

    if (enter && this.remember) {
      world.memory.set(this.remember, (world.memory.get(this.remember) || 0) + 1);
    }
    if (!this.out) return;
    if (this.mode === 'stay') world.setSignal(this.out, now ? this.value : !this.value);
    else if (this.mode === 'enter' && enter) world.setSignal(this.out, this.value);
    else if (this.mode === 'exit' && exit) world.setSignal(this.out, this.value);
    else if (this.mode === 'once' && enter && !this.latched) {
      this.latched = true;
      world.setSignal(this.out, this.value);
    }
  }
}

/**
 * Minuterie : rediffuse un signal après un délai.
 * Indispensable pour les portes temporisées et pour toutes les farces du type
 * « ça se referme trois secondes après ».
 */
export class Timer extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.from = spec.from || null; // null = démarre au début du niveau
    this.delay = spec.delay ?? 1;
    this.out = spec.emits || spec.out;
    this.value = spec.value !== false;
    this.repeat = spec.repeat === true;
    this.t = this.from ? -1 : this.delay;
    this.done = false;
    this.deadly = false;
  }

  preStep(world, dt) {
    if (this.done) return;
    if (this.t < 0) {
      if (this.from && world.isOn(this.from)) this.t = this.delay;
      return;
    }
    this.t -= dt;
    if (this.t <= 0) {
      world.setSignal(this.out, this.value);
      if (this.repeat) this.t = this.delay;
      else this.done = true;
    }
  }
}

/** Porte logique : permet de composer des mécanismes sans coder de niveau. */
export class Logic extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.op = spec.op || 'and';
    this.in = spec.in || [];
    this.out = spec.emits || spec.out;
  }

  preStep(world) {
    const vals = this.in.map((s) => world.isOn(s));
    let r;
    if (this.op === 'and') r = vals.every(Boolean);
    else if (this.op === 'or') r = vals.some(Boolean);
    else if (this.op === 'not') r = !vals[0];
    else r = vals.filter(Boolean).length === 1; // xor
    world.setSignal(this.out, r);
  }
}

/**
 * Point de contrôle.
 * `fake:true` : il a exactement l'allure d'un checkpoint, il fait le bon son,
 * il ne sauvegarde rien. Ce mensonge n'est utilisé qu'une seule fois dans tout
 * le jeu (niveau 18) — une farce répétée n'est plus une farce, c'est une taxe.
 */
export class Checkpoint extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.w = (spec.w ?? 1) * TILE;
    this.h = (spec.h ?? 1.5) * TILE;
    this.fake = spec.fake === true;
    this.taken = false;
    this.style = 'check';
  }

  postStep(world) {
    if (this.taken || !this.touchesPlayer(world)) return;
    this.taken = true;
    if (!this.fake) {
      world.checkpoint = { x: this.x + this.w / 2, y: this.y + this.h - 2 };
      world.fx('checkpoint', { x: this.x, y: this.y });
    } else {
      world.fx('checkpointFake', { x: this.x, y: this.y });
      if (this.spec.emits) world.setSignal(this.spec.emits, true);
    }
  }
}

/** Téléporteur. Envoie le robot sur `to` (en tuiles), une fois ou en boucle. */
export class Teleport extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.to = spec.to || [1, 1];
    this.cool = 0;
    this.keepVel = spec.keepVel === true;
    this.style = 'tele';
  }

  preStep(world, dt) {
    if (this.cool > 0) this.cool -= dt;
  }

  postStep(world) {
    if (this.cool > 0 || !this.touchesPlayer(world, -2)) return;
    this.cool = 0.45;
    world.player.warp(this.to[0] * TILE, this.to[1] * TILE, this.keepVel);
    if (this.spec.emits) world.setSignal(this.spec.emits, true);
  }
}

/**
 * Champ de gravité artificielle.
 * mode 'zone' : la gravité vaut `dir` tant qu'on est dedans.
 * mode 'flip' : elle s'inverse à l'entrée, et ça reste.
 */
export class Gravity extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.dir = spec.dir ?? -1;
    this.mode = spec.mode || 'flip';
    this.inside = false;
    this.style = 'grav';
  }

  postStep(world) {
    const now = this.touchesPlayer(world);
    const p = world.player;
    if (this.mode === 'zone') {
      if (now) p.gravDir = this.dir;
      else if (this.inside) p.gravDir = 1;
    } else if (now && !this.inside) {
      p.gravDir = this.dir === 'flip' ? -p.gravDir : this.dir;
      p.vy *= 0.2;
      world.fx('gravity', { x: p.cx, y: p.cy, dir: p.gravDir });
    }
    this.inside = now;
  }
}

/**
 * Mémoire inter-tentatives.
 * Émet un signal selon ce que le joueur a fait lors de ses essais précédents
 * (nombre de morts, zones visitées…). C'est ce qui permet à un niveau de
 * retourner la mémoire du joueur contre lui sans jamais tricher : la règle est
 * constante, c'est le joueur qui a changé.
 */
export class MemGate extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.key = spec.key;
    this.gte = spec.gte ?? null;
    this.lt = spec.lt ?? null;
    this.out = spec.emits || spec.out;
    this.value = spec.value !== false;
  }

  preStep(world) {
    const v = world.memory.get(this.key) || 0;
    let ok = true;
    if (this.gte != null) ok = ok && v >= this.gte;
    if (this.lt != null) ok = ok && v < this.lt;
    world.setSignal(this.out, ok ? this.value : !this.value);
  }
}

/**
 * Clone : rejoue le trajet enregistré lors de la tentative précédente.
 * S'il est mortel, le joueur affronte littéralement son propre passé.
 */
export class Clone extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.w = 20;
    this.h = 22;
    this.deadly = spec.deadly !== false;
    this.deathKind = DEATH.ZAP;
    this.offset = spec.delay ?? 0;
    // 'live'  : le robot poursuit son propre écho, avec `delay` de retard ;
    // 'prev'  : il affronte l'enregistrement de sa tentative précédente.
    this.source = spec.source || 'live';
    this.style = 'clone';
    this.facing = 1;
  }

  init(world) {
    this.trace = this.source === 'prev' ? (world.opts?.ghostTrace || null) : world.trace;
    this.active = !!this.trace;
  }

  preStep(world) {
    if (!this.trace) { this.deadly = false; return; }
    const i = Math.max(0, Math.floor((world.time - this.offset) * 60)) * 3;
    if (i + 2 >= this.trace.length) {
      this.deadly = false;
      this.alpha = 0;
      return;
    }
    if (world.time < this.offset) { this.deadly = false; this.alpha = 0; return; }
    this.alpha = 0.75;
    this.deadly = this.spec.deadly !== false;
    const nx = this.trace[i];
    this.facing = this.trace[i + 2] ? -1 : 1;
    this.x = nx;
    this.y = this.trace[i + 1];
  }
}

/**
 * Panneau d'information.
 * Peut dire la vérité. Peut mentir. Ne le fait jamais deux fois de la même
 * façon dans le même chapitre.
 */
export class Sign extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.text = spec.text || '';
    this.arrow = spec.arrow || null;
    this.always = spec.always === true;
    this.style = 'sign';
    this.vis = 0;
  }

  preStep(world, dt) {
    const target = this.always || this.touchesPlayer(world, TILE * 2) ? 1 : 0;
    this.vis += (target - this.vis) * Math.min(1, dt * 8);
  }
}

/** Décor. Certains décors sont vraiment du décor. Certains non. */
export class Deco extends Entity {
  constructor(spec, world) {
    super(spec, world);
    this.kind = spec.kind || 'panel';
    this.solid = spec.solid === true;
    this.style = 'deco';
  }
}
