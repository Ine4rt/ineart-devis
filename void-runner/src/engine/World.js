import { FIXED_DT, TILE, DEATH } from './constants.js';
import { aabb, makeRng } from './math.js';
import { Player } from './Player.js';
import { createEntity } from './entities/index.js';

/**
 * World — l'instance d'exécution d'un niveau.
 *
 * C'est le seul objet dont la physique dépend. Il ne connaît NI le DOM, NI le
 * canvas, NI l'audio : uniquement des données et du temps. Cette contrainte est
 * volontaire — elle permet de faire tourner un niveau en Node (voir tests/) et
 * donc de prouver automatiquement qu'il est franchissable.
 *
 * Les effets « sensibles » (son, vibration, particules) ne sont pas joués ici :
 * le monde les empile dans `events`, et la couche présentation les consomme.
 */
export class World {
  /**
   * @param {object} level      définition déclarative (voir src/data/levels/)
   * @param {object} [opts]     { dash: bool } capacités débloquées
   */
  constructor(level, opts = {}) {
    this.level = level;
    this.w = (level.w ?? 26) * TILE;
    this.h = (level.h ?? 15) * TILE;
    this.opts = opts;
    this.rng = makeRng(level.seed ?? 1234);

    this.time = 0; // temps écoulé dans la tentative
    this.frame = 0;
    this.state = 'run'; // 'run' | 'dead' | 'win'
    this.deathKind = null;
    this.deathAt = null; // {x,y} pour l'explosion
    this.events = []; // file d'effets pour la présentation

    this.signals = new Map(); // nom -> booléen (état courant)
    this.pulses = new Set(); // signaux déclenchés sur CETTE frame
    // Mémoire persistante entre deux tentatives d'un même niveau. Elle est
    // fournie de l'extérieur : c'est ce qui permet à un mécanisme de se
    // souvenir de ce que le joueur a fait AVANT de mourir (niveau 26).
    this.memory = opts.memory instanceof Map ? opts.memory : new Map();

    this.entities = [];
    for (const spec of level.entities ?? []) {
      const e = createEntity(spec, this);
      if (e) {
        e.world = this;
        this.entities.push(e);
      }
    }

    // Index de travail (recalculés à chaque frame, sans allocation)
    this.solids = [];
    this.hazards = [];

    // `startAt` : réapparition sur un point de contrôle atteint lors d'une
    // tentative précédente (en pixels, pas en tuiles).
    const start = opts.startAt;
    this.player = start
      ? new Player(start.x, start.y, this)
      : new Player((level.spawn?.x ?? 2) * TILE, (level.spawn?.y ?? 10) * TILE, this);
    this.checkpoint = start || null;

    // Trace du joueur, utilisée par les clones (mécanique chapitre 5)
    this.trace = [];

    for (const e of this.entities) e.init?.(this);
    this.reindex();
  }

  // --- Signaux ---------------------------------------------------------------

  /** État courant d'un signal (faux par défaut). */
  isOn(name) {
    return name ? this.signals.get(name) === true : false;
  }

  /** Pose un signal. Émet une impulsion sur le front montant. */
  setSignal(name, value) {
    if (!name) return;
    const prev = this.signals.get(name) === true;
    this.signals.set(name, value === true);
    if (value === true && !prev) this.pulses.add(name);
  }

  /** Le signal a-t-il été déclenché sur cette frame ? */
  pulsed(name) {
    return this.pulses.has(name);
  }

  // --- Effets (consommés par la couche présentation) -------------------------

  fx(type, data) {
    this.events.push({ type, ...data });
  }

  drainEvents() {
    const out = this.events;
    this.events = [];
    return out;
  }

  // --- Cycle de vie ----------------------------------------------------------

  reindex() {
    this.solids.length = 0;
    this.hazards.length = 0;
    for (const e of this.entities) {
      if (e.active === false) continue;
      if (e.solid) this.solids.push(e);
      if (e.deadly) this.hazards.push(e);
    }
  }

  kill(kind = DEATH.BOOM, src = null) {
    if (this.state !== 'run') return;
    this.state = 'dead';
    this.deathKind = kind;
    this.deathAt = { x: this.player.cx, y: this.player.cy };
    this.player.dead = true;
    this.fx('death', { kind, x: this.player.cx, y: this.player.cy, src });
  }

  win() {
    if (this.state !== 'run') return;
    this.state = 'win';
    this.fx('win', { x: this.player.cx, y: this.player.cy });
  }

  /**
   * Un pas de simulation. `input` : { left, right, jump, jumpPressed, dash }
   * Toujours appelé avec FIXED_DT — jamais avec un delta variable.
   */
  step(input) {
    if (this.state !== 'run') return;
    this.pulses.clear();
    this.time += FIXED_DT;
    this.frame++;

    // 1) Les solides bougent en premier : ils portent ou poussent le joueur.
    // On met à jour TOUTES les entités, y compris désactivées — une dalle
    // effondrée doit continuer à compter le temps qui la fera réapparaître.
    // (Filtrer sur `active` ici gelait définitivement toute dalle disparue :
    // bug trouvé par le solveur, pas à l'œil.)
    for (const e of this.entities) e.preStep?.(this, FIXED_DT);
    this.reindex();

    // 2) Le joueur agit — sauf s'il vient d'être écrasé pendant l'étape 1.
    // Faire avancer un robot déjà mort le projette hors de la salle et fausse
    // la position de l'explosion.
    if (this.state === 'run') this.player.step(input, FIXED_DT);

    // 3) Interactions dépendant de la position finale du joueur.
    for (const e of this.entities) e.postStep?.(this, FIXED_DT);

    // 4) Dangers. Vérifiés en dernier pour que la hitbox soit exacte.
    if (this.state === 'run') this.checkHazards();

    // 5) Sortie de salle. Vers le bas c'est une chute ; vers le haut aussi,
    // parce qu'en gravité inversée le vide est en haut.
    if (this.state === 'run' && (this.player.y > this.h + 80 || this.player.y < -240)) {
      this.kill(DEATH.FALL);
    }

    // Trace pour les clones (1 point / frame, borné).
    if (this.trace.length < 60 * 90) {
      this.trace.push(this.player.x, this.player.y, input.left ? 1 : 0);
    }
  }

  checkHazards() {
    const p = this.player;
    // Hitbox de mort légèrement plus petite que la hitbox physique : le joueur
    // ne doit jamais mourir "à un pixel près" sur un frôlement.
    const hx = p.x + 3, hy = p.y + 3, hw = p.w - 6, hh = p.h - 6;
    for (const e of this.hazards) {
      const b = e.hazardBox ? e.hazardBox() : e;
      if (!b) continue;
      if (aabb(hx, hy, hw, hh, b.x, b.y, b.w, b.h)) {
        this.kill(e.deathKind || DEATH.ZAP, e);
        return;
      }
    }
  }

  /** Le rectangle donné touche-t-il un solide (hors `ignore`) ? */
  overlapsSolid(x, y, w, h, ignore = null) {
    for (const s of this.solids) {
      if (s === ignore) continue;
      if (aabb(x, y, w, h, s.x, s.y, s.w, s.h)) return s;
    }
    return null;
  }

  /**
   * Copie complète et indépendante de l'état de simulation.
   *
   * Sert exclusivement au solveur de validation (tests/) : explorer un arbre
   * d'actions impose de repartir d'un état donné sans tout resimuler. Les
   * données immuables (spec, trajectoires) sont partagées volontairement.
   */
  clone() {
    const nw = Object.create(World.prototype);
    Object.assign(nw, this);
    nw.signals = new Map(this.signals);
    nw.pulses = new Set(this.pulses);
    nw.memory = new Map(this.memory);
    nw.events = [];
    nw.trace = this.trace.slice();
    nw.entities = this.entities.map((e) => {
      const ne = Object.create(Object.getPrototypeOf(e));
      Object.assign(ne, e);
      ne.world = nw;
      if (e.beam) ne.beam = { ...e.beam };
      if (e.trace === this.trace) ne.trace = nw.trace;
      return ne;
    });
    nw.player = Object.create(Object.getPrototypeOf(this.player));
    Object.assign(nw.player, this.player);
    nw.player.world = nw;
    nw.solids = [];
    nw.hazards = [];
    nw.reindex();
    return nw;
  }

  /** Première entité d'un type donné (utilitaire de level design). */
  find(type) {
    return this.entities.find((e) => e.type === type) || null;
  }

  findAll(type) {
    return this.entities.filter((e) => e.type === type);
  }
}
