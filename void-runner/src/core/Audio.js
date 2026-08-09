/**
 * Identité sonore — 100 % synthétisée à l'exécution (Web Audio).
 *
 * Choix technique : aucun fichier audio. Trois raisons, toutes décisives sur
 * mobile — poids nul (pas de 3 Mo de MP3 à télécharger), latence nulle (aucun
 * décodage, aucun buffer à charger avant le premier essai), et zéro risque de
 * réutiliser par accident un son existant.
 *
 * Le langage sonore est court et constant :
 *   sinus doux  → le robot (saut, atterrissage) ;
 *   carré filtré→ la machine (boutons, portes, mécanismes) ;
 *   bruit blanc → le danger (laser, effondrement, explosion).
 */
export class AudioManager {
  constructor() {
    this.ctx = null;
    this.master = null;
    this.musicGain = null;
    this.sfxGain = null;
    this.enabled = true;
    this.musicOn = true;
    this.started = false;
    this.step = 0;
    this.timer = null;
    this.scale = [0, 3, 5, 7, 10];
    this.root = 55; // La grave
  }

  /** Doit être appelé depuis un geste utilisateur (règle des navigateurs). */
  unlock() {
    if (this.ctx) { if (this.ctx.state === 'suspended') this.ctx.resume(); return; }
    const AC = window.AudioContext || window.webkitAudioContext;
    if (!AC) { this.enabled = false; return; }
    this.ctx = new AC();
    this.master = this.ctx.createGain();
    this.master.gain.value = 0.9;
    this.master.connect(this.ctx.destination);
    this.musicGain = this.ctx.createGain();
    this.musicGain.gain.value = 0.32;
    this.musicGain.connect(this.master);
    this.sfxGain = this.ctx.createGain();
    this.sfxGain.gain.value = 0.7;
    this.sfxGain.connect(this.master);
  }

  setVolumes({ music, sfx }) {
    this.musicOn = music > 0;
    if (this.musicGain) this.musicGain.gain.value = 0.32 * music;
    if (this.sfxGain) this.sfxGain.gain.value = 0.7 * sfx;
  }

  // --- Briques de synthèse ---------------------------------------------------

  tone(freq, dur, type = 'sine', vol = 0.3, dest = null, slideTo = null) {
    if (!this.ctx || !this.enabled) return;
    const t = this.ctx.currentTime;
    const o = this.ctx.createOscillator();
    const g = this.ctx.createGain();
    o.type = type;
    o.frequency.setValueAtTime(freq, t);
    if (slideTo) o.frequency.exponentialRampToValueAtTime(Math.max(20, slideTo), t + dur);
    g.gain.setValueAtTime(0, t);
    g.gain.linearRampToValueAtTime(vol, t + 0.008);
    g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
    o.connect(g).connect(dest || this.sfxGain);
    o.start(t);
    o.stop(t + dur + 0.02);
  }

  noise(dur, vol = 0.3, freq = 1200, q = 1, type = 'lowpass') {
    if (!this.ctx || !this.enabled) return;
    const t = this.ctx.currentTime;
    const n = Math.floor(this.ctx.sampleRate * dur);
    const buf = this.ctx.createBuffer(1, n, this.ctx.sampleRate);
    const d = buf.getChannelData(0);
    for (let i = 0; i < n; i++) d[i] = (Math.random() * 2 - 1) * (1 - i / n);
    const src = this.ctx.createBufferSource();
    src.buffer = buf;
    const f = this.ctx.createBiquadFilter();
    f.type = type;
    f.frequency.value = freq;
    f.Q.value = q;
    const g = this.ctx.createGain();
    g.gain.value = vol;
    src.connect(f).connect(g).connect(this.sfxGain);
    src.start(t);
  }

  // --- Vocabulaire du jeu ----------------------------------------------------

  play(name) {
    if (!this.ctx || !this.enabled) return;
    switch (name) {
      case 'jump': this.tone(430, 0.12, 'sine', 0.22, null, 660); break;
      case 'land': this.tone(150, 0.09, 'sine', 0.18, null, 90); break;
      case 'dash': this.tone(880, 0.14, 'sawtooth', 0.14, null, 300); this.noise(0.12, 0.1, 2400); break;
      case 'click': this.tone(720, 0.05, 'square', 0.14); this.tone(1080, 0.06, 'square', 0.07); break;
      case 'unclick': this.tone(380, 0.06, 'square', 0.1); break;
      case 'mech': this.tone(120, 0.22, 'square', 0.1, null, 180); break;
      case 'door': this.tone(90, 0.3, 'square', 0.12, null, 150); this.noise(0.25, 0.06, 700); break;
      case 'laser': this.tone(1400, 0.16, 'sawtooth', 0.1, null, 900); this.noise(0.14, 0.08, 3000, 3, 'bandpass'); break;
      case 'creak': this.noise(0.3, 0.12, 500, 2, 'bandpass'); break;
      case 'crumble': this.noise(0.45, 0.22, 380); this.tone(70, 0.3, 'square', 0.1, null, 40); break;
      case 'slam': this.noise(0.2, 0.3, 220); this.tone(60, 0.25, 'square', 0.22, null, 35); break;
      // Jaillissement d'une grille : claquement métallique bref et sec.
      case 'snap': this.noise(0.09, 0.3, 2600, 4, 'bandpass'); this.tone(140, 0.12, 'square', 0.2, null, 420); break;
      // Mirage traversé : le sol « décroche », son de circuit qui lâche.
      case 'mirage': this.tone(320, 0.2, 'triangle', 0.16, null, 70); this.noise(0.16, 0.1, 1400); break;
      case 'alarm': this.tone(880, 0.08, 'square', 0.09); setTimeout(() => this.tone(660, 0.08, 'square', 0.09), 90); break;
      case 'warp': this.tone(200, 0.28, 'sine', 0.16, null, 1600); break;
      case 'gravity': this.tone(300, 0.35, 'triangle', 0.18, null, 120); break;
      case 'checkpoint': this.tone(660, 0.1, 'sine', 0.18); setTimeout(() => this.tone(990, 0.16, 'sine', 0.16), 90); break;
      case 'secret':
        [660, 880, 1320].forEach((f, i) => setTimeout(() => this.tone(f, 0.18, 'triangle', 0.18), i * 90));
        break;
      case 'death':
        // Le fameux « effet cartoon court » : une descente rapide + un souffle.
        this.tone(500, 0.22, 'square', 0.2, null, 60);
        this.noise(0.28, 0.28, 900);
        break;
      case 'win':
        [523, 659, 784, 1046].forEach((f, i) => setTimeout(() => this.tone(f, 0.22, 'triangle', 0.2), i * 85));
        break;
      case 'ui': this.tone(560, 0.05, 'sine', 0.12); break;
      case 'unlock':
        [392, 523, 659, 880].forEach((f, i) => setTimeout(() => this.tone(f, 0.3, 'sine', 0.16), i * 110));
        break;
      default: break;
    }
  }

  // --- Musique ---------------------------------------------------------------

  /**
   * Séquenceur minimal : une nappe tenue + un arpège sur pentatonique mineure.
   * Calme, un peu inquiétant, jamais mélodique au point de fatiguer après la
   * quarantième mort dans la même salle.
   */
  startMusic(chapterId = 1) {
    if (!this.ctx || !this.musicOn || this.started) return;
    this.started = true;
    this.root = [55, 58.3, 49, 65.4, 43.7][(chapterId - 1) % 5];
    const pad = this.ctx.createOscillator();
    const padG = this.ctx.createGain();
    const filt = this.ctx.createBiquadFilter();
    filt.type = 'lowpass';
    filt.frequency.value = 420;
    pad.type = 'sawtooth';
    pad.frequency.value = this.root;
    padG.gain.value = 0.08;
    pad.connect(filt).connect(padG).connect(this.musicGain);
    pad.start();
    this.pad = pad;

    const beat = 0.32;
    this.timer = setInterval(() => {
      if (!this.musicOn) return;
      const s = this.step++;
      if (s % 2 === 0) {
        const deg = this.scale[(Math.floor(s / 2) * 3) % this.scale.length];
        const oct = 4 + ((Math.floor(s / 8)) % 2);
        const f = this.root * Math.pow(2, oct + deg / 12);
        this.tone(f, 0.5, 'triangle', 0.05, this.musicGain);
      }
      if (s % 8 === 0) this.tone(this.root * 2, 0.6, 'sine', 0.06, this.musicGain);
    }, beat * 1000);
  }

  stopMusic() {
    if (this.timer) clearInterval(this.timer);
    this.timer = null;
    if (this.pad) { try { this.pad.stop(); } catch { /* déjà arrêté */ } this.pad = null; }
    this.started = false;
  }

  /** Assourdit la musique pendant les menus / écrans de victoire. */
  duck(on) {
    if (!this.musicGain || !this.ctx) return;
    const t = this.ctx.currentTime;
    this.musicGain.gain.cancelScheduledValues(t);
    this.musicGain.gain.linearRampToValueAtTime(on ? 0.12 : 0.32, t + 0.25);
  }
}
