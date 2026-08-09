/**
 * Entrées : tactile, clavier, manette.
 *
 * Deux principes qui décident de tout le reste :
 *
 * 1. Les boutons tactiles sont des éléments DOM, pas des zones dessinées sur
 *    le canvas. Le navigateur gère alors nativement le multi-touch, les zones
 *    de frappe et l'accessibilité — et surtout on peut agrandir la surface
 *    tactile bien au-delà du visuel, ce qui est LA clé du confort au pouce.
 *
 * 2. L'appui de saut est mis en file d'attente, jamais lu « à l'instant T ».
 *    Un appui de 30 ms entre deux pas de simulation doit être entendu : sinon
 *    le joueur jure qu'il a appuyé, et il a raison.
 */
export class Input {
  constructor(root) {
    this.state = { left: false, right: false, jump: false, dash: false };
    this.jumpQueued = false;
    this.dashQueued = false;
    this.pointers = new Map();
    this.padIndex = null;
    this.padPrev = { jump: false, dash: false };
    this.onAnyInput = null;
    this.bindKeyboard();
    if (root) this.bindTouch(root);
    this.bindPad();
  }

  press(key, down) {
    if (down && this.onAnyInput) this.onAnyInput();
    if (key === 'jump') {
      if (down && !this.state.jump) this.jumpQueued = true;
      this.state.jump = down;
    } else if (key === 'dash') {
      if (down && !this.state.dash) this.dashQueued = true;
      this.state.dash = down;
    } else {
      this.state[key] = down;
    }
  }

  bindKeyboard() {
    const map = {
      ArrowLeft: 'left', KeyA: 'left', KeyQ: 'left',
      ArrowRight: 'right', KeyD: 'right',
      ArrowUp: 'jump', Space: 'jump', KeyW: 'jump', KeyZ: 'jump',
      ShiftLeft: 'dash', ShiftRight: 'dash', KeyX: 'dash',
    };
    addEventListener('keydown', (e) => {
      const k = map[e.code];
      if (!k) return;
      if (!e.repeat) this.press(k, true);
      e.preventDefault();
    });
    addEventListener('keyup', (e) => {
      const k = map[e.code];
      if (!k) return;
      this.press(k, false);
      e.preventDefault();
    });
    // Perdre le focus ne doit jamais laisser une touche « collée ».
    addEventListener('blur', () => this.releaseAll());
  }

  bindTouch(root) {
    const buttons = root.querySelectorAll('[data-key]');
    for (const btn of buttons) {
      const key = btn.dataset.key;
      const down = (e) => {
        e.preventDefault();
        this.pointers.set(e.pointerId, key);
        btn.classList.add('is-down');
        this.press(key, true);
        try { btn.setPointerCapture(e.pointerId); } catch { /* non capturable */ }
      };
      const up = (e) => {
        e.preventDefault();
        this.pointers.delete(e.pointerId);
        btn.classList.remove('is-down');
        this.press(key, false);
      };
      btn.addEventListener('pointerdown', down);
      btn.addEventListener('pointerup', up);
      btn.addEventListener('pointercancel', up);
      btn.addEventListener('pointerleave', (e) => {
        // Glisser hors du bouton relâche : évite les touches fantômes.
        if (this.pointers.has(e.pointerId)) up(e);
      });
      btn.addEventListener('contextmenu', (e) => e.preventDefault());
    }
  }

  bindPad() {
    addEventListener('gamepadconnected', (e) => { this.padIndex = e.gamepad.index; });
    addEventListener('gamepaddisconnected', () => { this.padIndex = null; });
  }

  /** Lecture manette, appelée une fois par frame. */
  pollPad() {
    if (this.padIndex == null || !navigator.getGamepads) return;
    const gp = navigator.getGamepads()[this.padIndex];
    if (!gp) return;
    const ax = gp.axes[0] ?? 0;
    const dpadL = gp.buttons[14]?.pressed;
    const dpadR = gp.buttons[15]?.pressed;
    this.state.left = dpadL || ax < -0.35;
    this.state.right = dpadR || ax > 0.35;
    const jump = gp.buttons[0]?.pressed || gp.buttons[1]?.pressed;
    const dash = gp.buttons[2]?.pressed || gp.buttons[7]?.pressed;
    if (jump && !this.padPrev.jump) this.jumpQueued = true;
    if (dash && !this.padPrev.dash) this.dashQueued = true;
    this.state.jump = jump || this.state.jump;
    this.state.dash = dash || this.state.dash;
    this.padPrev.jump = jump;
    this.padPrev.dash = dash;
  }

  /**
   * Snapshot pour UN pas de simulation. Les fronts montants ne sont servis
   * qu'une seule fois — c'est ce qui empêche un appui d'être compté deux fois
   * quand la frame contient plusieurs pas fixes.
   */
  consume() {
    const s = {
      left: this.state.left,
      right: this.state.right,
      jump: this.state.jump,
      jumpPressed: this.jumpQueued,
      dashPressed: this.dashQueued,
    };
    this.jumpQueued = false;
    this.dashQueued = false;
    return s;
  }

  releaseAll() {
    this.state.left = this.state.right = this.state.jump = this.state.dash = false;
    this.jumpQueued = this.dashQueued = false;
    for (const el of document.querySelectorAll('[data-key].is-down')) el.classList.remove('is-down');
  }
}
