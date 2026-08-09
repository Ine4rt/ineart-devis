import { CHAPTERS, chapterRanges, LEVEL_COUNT } from '../data/levels/index.js';
import { SKINS, SKIN_BY_ID, isSkinUnlocked } from '../data/skins.js';
import { drawRobotPortrait } from '../render/Robot.js';
import { formatTime, formatInt } from '../engine/math.js';

/**
 * UIManager — tous les écrans hors jeu, en DOM.
 *
 * Pourquoi le DOM plutôt que du dessin sur le canvas : accessibilité (lecteurs
 * d'écran, tailles de police système), défilement natif fluide au doigt, zones
 * de frappe correctes sans les recoder, et surtout zéro coût pendant le jeu —
 * un écran masqué ne consomme rien, alors qu'une UI dessinée coûte à chaque
 * frame, y compris quand elle est invisible.
 */
export class UIManager {
  constructor(game) {
    this.game = game;
    game.ui = this;
    this.el = {};
    const id = (s) => document.getElementById(s);
    this.el.hud = id('hud');
    this.el.touch = id('touch');
    this.el.num = id('hud-num');
    this.el.name = id('hud-name');
    this.el.attempt = id('hud-attempt');
    this.el.time = id('hud-time');
    this.el.band = id('deathband');
    this.el.bandText = id('deathband-text');
    this.el.flash = id('flash');
    this.el.toasts = id('toast-wrap');
    this.el.dash = document.querySelector('.pad-dash');
    this.screens = {};
    for (const s of document.querySelectorAll('.screen')) {
      this.screens[s.dataset.screen] = s;
    }
    this.current = 'menu';
    this.bandTimer = null;
    this.menuT = 0;

    this.bindNav();
    this.applySettings();
    this.showScreen('menu');
    this.animateMenuRobot();
  }

  // --- Navigation ------------------------------------------------------------

  bindNav() {
    document.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-go]');
      if (!btn) return;
      this.game.audio.unlock();
      this.game.audio.play('ui');
      this.go(btn.dataset.go);
    });
    document.getElementById('btn-pause').addEventListener('click', () => this.game.pause());
    addEventListener('keydown', (e) => {
      if (e.code === 'Escape') {
        if (this.game.mode === 'play') this.game.pause();
        else if (this.game.mode === 'paused') this.game.resume();
      }
    });
    // Perdre le focus (appel, notification) met en pause : on ne meurt pas
    // parce que le téléphone a sonné.
    document.addEventListener('visibilitychange', () => {
      if (document.hidden && this.game.mode === 'play') this.game.pause();
    });
  }

  go(where) {
    const g = this.game;
    switch (where) {
      case 'play': g.startLevel(g.save.maxUnlockedLevel(LEVEL_COUNT)); this.showScreen(null); break;
      // L'écran est affiché AVANT d'être rempli : les aperçus animés se
      // coupent d'eux-mêmes tant que leur écran n'est pas l'écran courant, et
      // resteraient donc vides si on inversait ces deux lignes.
      case 'levels': this.showScreen('levels'); this.buildLevels(); break;
      case 'robot': this.showScreen('robot'); this.buildSkins(); break;
      case 'stats': this.showScreen('stats'); this.buildStats(); break;
      case 'settings': this.showScreen('settings'); this.buildSettings(); break;
      case 'menu': g.quitToMenu(); this.refreshMenu(); this.showScreen('menu'); break;
      case 'resume': g.resume(); break;
      case 'retry': g.mode = 'play'; g.audio.duck(false); g.retry(); this.showScreen(null); g.start(); break;
      case 'next': g.nextLevel(); this.showScreen(null); break;
      default: break;
    }
  }

  showScreen(name) {
    for (const [k, el] of Object.entries(this.screens)) el.hidden = k !== name;
    this.current = name;
    const playing = name === null;
    this.el.hud.hidden = !playing;
    this.el.touch.hidden = !playing;
    if (name === 'menu') this.refreshMenu();
  }

  // --- Pendant le jeu --------------------------------------------------------

  onLevelStart(level, attempt) {
    this.el.num.textContent = String(level.num).padStart(2, '0');
    this.el.name.textContent = level.name.toUpperCase();
    this.el.attempt.textContent = attempt;
    this.el.dash.hidden = !level.dash;
    if (level.hint) this.band(level.hint, 1800);
  }

  onAttempt(n) {
    this.el.attempt.textContent = n;
  }

  setPauseInfo(text) {
    document.getElementById('pause-sub').textContent = text;
  }

  onFrame(world, mode) {
    if (mode === 'play' || mode === 'paused') {
      this.el.time.textContent = world.time.toFixed(2);
    }
  }

  onDeath(nextAttempt, quip) {
    // Le compteur d'essai est TOUJOURS affiché, la vanne seulement parfois.
    this.band(quip ? `${quip} · ESSAI ${nextAttempt}` : `ESSAI ${nextAttempt}`, 900);
  }

  band(text, ms) {
    this.el.bandText.textContent = text;
    this.el.band.hidden = false;
    this.el.band.style.animation = 'none';
    void this.el.band.offsetWidth;
    this.el.band.style.animation = '';
    clearTimeout(this.bandTimer);
    this.bandTimer = setTimeout(() => { this.el.band.hidden = true; }, ms);
  }

  flash() {
    this.el.flash.classList.remove('on');
    void this.el.flash.offsetWidth;
    this.el.flash.classList.add('on');
  }

  toast(title, desc) {
    const d = document.createElement('div');
    d.className = 'toast';
    d.innerHTML = `<b></b><small></small>`;
    d.querySelector('b').textContent = title;
    d.querySelector('small').textContent = desc || '';
    this.el.toasts.appendChild(d);
    setTimeout(() => {
      d.style.transition = 'opacity .35s, transform .35s';
      d.style.opacity = '0';
      d.style.transform = 'translateX(12px)';
      setTimeout(() => d.remove(), 400);
    }, 3200);
  }

  showWin(info) {
    document.getElementById('win-title').textContent =
      `SALLE ${String(info.level.num).padStart(2, '0')} — ${info.level.name.toUpperCase()}`;
    document.getElementById('win-line').textContent =
      info.isLast ? 'Tu es sorti. La station, elle, est toujours là.' : info.line;
    document.getElementById('win-time').textContent = formatTime(info.time);
    document.getElementById('win-deaths').textContent = info.deaths;
    document.getElementById('win-best').textContent = formatTime(info.best);
    const next = document.querySelector('[data-go="next"]');
    next.textContent = info.isLast ? 'FIN' : 'SALLE SUIVANTE';
    this.showScreen('win');
  }

  // --- Menu ------------------------------------------------------------------

  refreshMenu() {
    const n = this.game.save.maxUnlockedLevel(LEVEL_COUNT);
    const done = this.game.save.stats(LEVEL_COUNT).completedCount;
    document.getElementById('menu-play-label').textContent = done ? 'CONTINUER' : 'JOUER';
    document.getElementById('menu-play-sub').textContent =
      done >= LEVEL_COUNT ? 'Toutes les salles franchies' : `Salle ${String(n).padStart(2, '0')}`;
  }

  animateMenuRobot() {
    const cv = document.getElementById('menu-robot');
    const ctx = cv.getContext('2d');
    const loop = () => {
      requestAnimationFrame(loop);
      if (this.current !== 'menu') return;
      this.menuT += 1 / 60;
      ctx.clearRect(0, 0, cv.width, cv.height);
      ctx.save();
      ctx.translate(cv.width / 2, cv.height / 2 + Math.sin(this.menuT * 1.6) * 6);
      drawRobotPortrait(ctx, 0, 0, this.game.skin, 3.4, this.menuT, 'idle');
      ctx.restore();
    };
    loop();
  }

  // --- Sélection de niveaux --------------------------------------------------

  buildLevels() {
    const save = this.game.save;
    const maxUnlocked = save.maxUnlockedLevel(LEVEL_COUNT);
    const host = document.getElementById('levels-list');
    host.innerHTML = '';
    const st = save.stats(LEVEL_COUNT);
    document.getElementById('levels-meta').textContent =
      `${st.completedCount}/${LEVEL_COUNT} · ${formatInt(st.totalDeaths)} morts`;

    for (const { chapter, from, to } of chapterRanges()) {
      const head = document.createElement('div');
      head.className = 'chapter-title';
      head.innerHTML = '<h3></h3><span></span>';
      head.querySelector('h3').textContent = `CHAPITRE ${chapter.id} — ${chapter.name.toUpperCase()}`;
      head.querySelector('span').textContent = chapter.subtitle;
      host.appendChild(head);

      const grid = document.createElement('div');
      grid.className = 'lv-grid';
      for (let n = from; n <= to; n++) {
        const rec = save.data.levels[String(n)] || {};
        const locked = n > maxUnlocked;
        const b = document.createElement('button');
        b.className = `lv${rec.done ? ' done' : ''}${locked ? ' locked' : ''}`;
        b.disabled = locked;
        const lvl = CHAPTERS[chapter.id - 1].levels[n - from];
        b.innerHTML = `
          <div class="lv-num">${String(n).padStart(2, '0')}<i>${rec.done ? '✓' : ''}</i></div>
          <div class="lv-name"></div>
          <div class="lv-stat"><span>☠ <b>${rec.deaths || 0}</b></span><span>${rec.best ? formatTime(rec.best) : '—'}</span>${rec.secret ? '<span class="lv-secret">★</span>' : ''}</div>`;
        b.querySelector('.lv-name').textContent = locked ? '— verrouillé —' : lvl.name;
        b.addEventListener('click', () => {
          this.game.audio.unlock();
          this.game.startLevel(n);
          this.showScreen(null);
        });
        grid.appendChild(b);
      }
      host.appendChild(grid);
    }
  }

  // --- Robots ----------------------------------------------------------------

  buildSkins() {
    const save = this.game.save;
    const stats = save.stats(LEVEL_COUNT);
    const host = document.getElementById('skins-list');
    host.innerHTML = '';
    const unlockedCount = SKINS.filter((s) => isSkinUnlocked(s, stats)).length;
    document.getElementById('robot-meta').textContent = `${unlockedCount}/${SKINS.length} débloqués`;

    const preview = document.getElementById('robot-preview');
    const pctx = preview.getContext('2d');
    const showInfo = (skin, unlocked) => {
      document.getElementById('robot-name').textContent = skin.name;
      document.getElementById('robot-desc').textContent = skin.desc;
      document.getElementById('robot-lock').textContent =
        unlocked ? '' : `Verrouillé — ${skin.unlock.label}`;
      let t = 0;
      const draw = () => {
        if (this.current !== 'robot' || this.previewId !== skin.id) return;
        requestAnimationFrame(draw);
        t += 1 / 60;
        pctx.clearRect(0, 0, preview.width, preview.height);
        pctx.save();
        pctx.translate(preview.width / 2, preview.height / 2 + Math.sin(t * 1.8) * 5);
        drawRobotPortrait(pctx, 0, 0, skin, 3.6, t, 'idle');
        pctx.restore();
      };
      this.previewId = skin.id;
      draw();
    };

    for (const skin of SKINS) {
      const unlocked = save.data.unlocked.includes(skin.id) || isSkinUnlocked(skin, stats);
      const cell = document.createElement('button');
      cell.className = `skin${save.data.skin === skin.id ? ' sel' : ''}${unlocked ? '' : ' locked'}`;
      cell.innerHTML = '<canvas width="120" height="80"></canvas><span></span>';
      cell.querySelector('span').textContent = unlocked ? skin.name : '???';
      const c = cell.querySelector('canvas');
      const cx = c.getContext('2d');
      cx.save();
      cx.translate(60, 42);
      // t = 1,2 s et non 0 : à t = 0 le robot est en plein clignement et
      // toutes les vignettes auraient les yeux fermés.
      drawRobotPortrait(cx, 0, 0, skin, 2.4, 1.2, unlocked ? 'idle' : 'dead');
      cx.restore();
      cell.addEventListener('click', () => {
        showInfo(skin, unlocked);
        if (!unlocked) { this.game.audio.play('unclick'); return; }
        save.data.skin = skin.id;
        save.save();
        this.game.audio.play('click');
        for (const el of host.querySelectorAll('.skin')) el.classList.remove('sel');
        cell.classList.add('sel');
      });
      host.appendChild(cell);
    }
    const cur = SKIN_BY_ID[save.data.skin] || SKINS[0];
    showInfo(cur, true);
  }

  // --- Statistiques ----------------------------------------------------------

  buildStats() {
    const save = this.game.save;
    const s = save.stats(LEVEL_COUNT);
    const host = document.getElementById('stats-body');
    const worst = s.worstLevel && s.worstLevel[1].deaths > 0
      ? `Salle ${String(s.worstLevel[0]).padStart(2, '0')} — ${s.worstLevel[1].deaths} morts`
      : 'Aucune salle ne t\'a encore vaincu';
    const hours = Math.floor(s.playTime / 3600);
    const mins = Math.floor((s.playTime % 3600) / 60);

    // Le compteur de morts est mis en avant, pas caché : dans ce jeu, mourir
    // est la mécanique principale. Autant en faire un trophée.
    host.innerHTML = `
      <div class="stat-grid">
        <div class="stat"><b>${formatInt(s.totalDeaths)}</b><small>MORTS TOTALES</small></div>
        <div class="stat"><b>${formatInt(s.deathsToday)}</b><small>MORTS AUJOURD'HUI</small></div>
        <div class="stat"><b>${s.completedCount}/${LEVEL_COUNT}</b><small>SALLES FRANCHIES</small></div>
        <div class="stat"><b>${s.secrets}/3</b><small>SALLES SECRÈTES</small></div>
        <div class="stat"><b>${s.bestTimeAny ? formatTime(s.bestTimeAny) : '—'}</b><small>MEILLEUR TEMPS</small></div>
        <div class="stat"><b>${s.bestCleanStreak}</b><small>SÉRIE SANS MOURIR</small></div>
        <div class="stat wide"><b>${worst}</b><small>TA NÉMÉSIS</small></div>
        <div class="stat wide"><b>${hours} h ${String(mins).padStart(2, '0')}</b><small>TEMPS DE JEU</small></div>
      </div>
      <div class="chapter-title"><h3>ACCOMPLISSEMENTS</h3></div>
      <div id="ach-list"></div>`;

    const list = host.querySelector('#ach-list');
    for (const a of this.game.achievements.earned()) {
      const row = document.createElement('div');
      row.className = `ach ${a.done ? 'done' : 'todo'}`;
      row.innerHTML = `<i>${a.done ? '★' : '☆'}</i><div><b></b><small></small></div>`;
      row.querySelector('b').textContent = a.name;
      row.querySelector('small').textContent = a.desc;
      list.appendChild(row);
    }
  }

  // --- Paramètres ------------------------------------------------------------

  buildSettings() {
    const st = this.game.save.data.settings;
    const host = document.getElementById('settings-body');
    host.innerHTML = `
      <div class="row"><label>MUSIQUE</label><input type="range" id="set-music" min="0" max="1" step="0.05"></div>
      <div class="row"><label>EFFETS SONORES</label><input type="range" id="set-sfx" min="0" max="1" step="0.05"></div>
      <div class="row"><label>VIBRATIONS<small>Mort, impact, réussite, mécanisme</small></label>
        <div class="seg" id="set-hap"><button data-v="1">OUI</button><button data-v="0">NON</button></div></div>
      <div class="row"><label>MAIN DIRECTRICE<small>Inverse la position des commandes</small></label>
        <div class="seg" id="set-hand"><button data-v="right">DROITE</button><button data-v="left">GAUCHE</button></div></div>
      <div class="row"><label>EFFACER LA PROGRESSION<small>Irréversible : niveaux, morts, robots, records</small></label>
        <button class="btn danger-btn" id="set-reset">EFFACER</button></div>
      <div class="row"><label style="opacity:.5">VOID RUNNER — prototype jouable<small>30 salles · moteur déterministe · aucune publicité intrusive</small></label></div>`;

    const music = host.querySelector('#set-music');
    const sfx = host.querySelector('#set-sfx');
    music.value = st.music;
    sfx.value = st.sfx;
    const apply = () => {
      st.music = Number(music.value);
      st.sfx = Number(sfx.value);
      this.game.save.save();
      this.applySettings();
    };
    music.addEventListener('input', apply);
    sfx.addEventListener('input', () => { apply(); this.game.audio.play('ui'); });

    const seg = (host2, value, onPick) => {
      for (const b of host2.querySelectorAll('button')) {
        b.classList.toggle('on', String(value) === b.dataset.v);
        b.addEventListener('click', () => {
          onPick(b.dataset.v);
          for (const o of host2.querySelectorAll('button')) o.classList.toggle('on', o === b);
        });
      }
    };
    seg(host.querySelector('#set-hap'), st.haptics ? 1 : 0, (v) => {
      st.haptics = v === '1';
      this.game.save.save();
      this.applySettings();
      if (st.haptics) this.game.haptics.fire('impact');
    });
    seg(host.querySelector('#set-hand'), st.hand, (v) => {
      st.hand = v;
      this.game.save.save();
      this.applySettings();
    });

    let armed = false;
    const reset = host.querySelector('#set-reset');
    reset.addEventListener('click', () => {
      // Double confirmation sur place : pas de boîte de dialogue système, mais
      // pas d'effacement accidentel non plus.
      if (!armed) {
        armed = true;
        reset.textContent = 'CONFIRMER ?';
        setTimeout(() => { armed = false; reset.textContent = 'EFFACER'; }, 3500);
        return;
      }
      this.game.save.reset();
      this.applySettings();
      this.buildSettings();
      this.toast('PROGRESSION EFFACÉE', 'On repart de la salle 01.');
    });
  }

  applySettings() {
    const st = this.game.save.data.settings;
    this.game.audio.setVolumes({ music: st.music, sfx: st.sfx });
    this.game.haptics.enabled = st.haptics !== false;
    document.body.classList.toggle('hand-left', st.hand === 'left');
  }
}
