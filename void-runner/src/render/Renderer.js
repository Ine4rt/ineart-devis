import { C, palette, roundRect } from './theme.js';
import { drawRobot } from './Robot.js';
import { TILE } from '../engine/constants.js';

/**
 * Rendu de la salle.
 *
 * Contraintes mobiles qui dictent toute l'implémentation :
 *  - un seul canvas, un seul contexte, aucune texture ;
 *  - `shadowBlur` réservé aux petits éléments (sortie, faisceaux, robot) ;
 *  - le fond est redessiné à plat, sans blur plein écran ;
 *  - rien n'est dessiné hors du champ de la caméra.
 * Résultat : 60 FPS tenus sur un mobile d'entrée de gamme.
 */
export class Renderer {
  constructor(canvas) {
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d', { alpha: false, desynchronized: true });
    this.t = 0;
    this.quality = 1; // 1 = complet, 0 = économe (baissé automatiquement)
  }

  /** Le fond : dégradé, grille en parallaxe, brume. Zéro allocation par frame. */
  drawBackground(w, h, cam, pal) {
    const ctx = this.ctx;
    const g = ctx.createLinearGradient(0, 0, 0, h);
    g.addColorStop(0, pal.bg0);
    g.addColorStop(1, pal.bg1);
    ctx.fillStyle = g;
    ctx.fillRect(0, 0, w, h);

    // Grille lointaine — parallaxe 0,35 : elle donne la vitesse sans distraire.
    const step = 48;
    const ox = -((cam.x * 0.35) % step);
    const oy = -((cam.y * 0.35) % step);
    ctx.strokeStyle = pal.grid;
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let x = ox; x < w + step; x += step) {
      ctx.moveTo(Math.round(x) + 0.5, 0);
      ctx.lineTo(Math.round(x) + 0.5, h);
    }
    for (let y = oy; y < h + step; y += step) {
      ctx.moveTo(0, Math.round(y) + 0.5);
      ctx.lineTo(w, Math.round(y) + 0.5);
    }
    ctx.stroke();

    if (this.quality > 0) {
      const hz = ctx.createRadialGradient(w * 0.5, h * 0.45, 10, w * 0.5, h * 0.45, w * 0.7);
      hz.addColorStop(0, pal.haze);
      hz.addColorStop(1, 'rgba(0,0,0,0)');
      ctx.fillStyle = hz;
      ctx.fillRect(0, 0, w, h);
    }
  }

  // --- Pièces du décor -------------------------------------------------------

  drawSolid(e) {
    const ctx = this.ctx;
    if (e.style === 'bound') return; // les murs de bordure restent invisibles
    const g = ctx.createLinearGradient(0, e.y, 0, e.y + e.h);
    g.addColorStop(0, C.wallLit);
    g.addColorStop(0.18, C.wall);
    g.addColorStop(1, '#1a1f42');
    ctx.fillStyle = g;
    roundRect(ctx, e.x, e.y, e.w, e.h, 3);
    ctx.fill();
    // Arête supérieure : c'est elle qui rend le sol lisible d'un coup d'œil.
    ctx.fillStyle = C.wallEdge;
    ctx.fillRect(e.x + 1.5, e.y, e.w - 3, 2);
  }

  drawCrumble(e) {
    const ctx = this.ctx;
    const shake = e.shake || 0;
    const dx = shake ? (Math.random() - 0.5) * shake * 2.4 : 0;
    ctx.save();
    ctx.globalAlpha = e.active ? 1 : 0.14;
    ctx.translate(dx, 0);
    ctx.fillStyle = e.active ? '#2c2f5c' : '#191c38';
    roundRect(ctx, e.x, e.y, e.w, e.h, 2);
    ctx.fill();
    ctx.strokeStyle = shake > 0.3 ? C.danger : '#4a4f8c';
    ctx.lineWidth = 1;
    ctx.stroke();
    // Fissures : elles s'ouvrent avec le compte à rebours.
    if (e.active) {
      ctx.strokeStyle = `rgba(255,138,61,${0.15 + shake * 0.75})`;
      ctx.lineWidth = 0.9;
      ctx.beginPath();
      ctx.moveTo(e.x + e.w * 0.2, e.y);
      ctx.lineTo(e.x + e.w * 0.45, e.y + e.h * 0.6);
      ctx.lineTo(e.x + e.w * 0.3, e.y + e.h);
      ctx.moveTo(e.x + e.w * 0.7, e.y);
      ctx.lineTo(e.x + e.w * 0.6, e.y + e.h * 0.5);
      ctx.stroke();
    }
    ctx.restore();
  }

  drawVanish(e) {
    const ctx = this.ctx;
    ctx.save();
    ctx.globalAlpha = Math.max(0.12, e.alpha);
    ctx.fillStyle = 'rgba(63,232,255,0.20)';
    roundRect(ctx, e.x, e.y, e.w, e.h, 3);
    ctx.fill();
    ctx.strokeStyle = C.cyan;
    ctx.lineWidth = 1.2;
    ctx.stroke();
    ctx.restore();
  }

  drawGhost(e) {
    const ctx = this.ctx;
    const v = e.vis ?? 0;
    if (v < 0.02) return;
    ctx.save();
    ctx.globalAlpha = v * 0.85;
    ctx.fillStyle = 'rgba(160,190,255,0.18)';
    roundRect(ctx, e.x, e.y, e.w, e.h, 3);
    ctx.fill();
    ctx.setLineDash([4, 4]);
    ctx.strokeStyle = 'rgba(200,220,255,0.9)';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.restore();
  }

  drawMover(e) {
    const ctx = this.ctx;
    if (e.style === 'saw') {
      const cx = e.x + e.w / 2;
      const cy = e.y + e.h / 2;
      const r = Math.max(e.w, e.h) * 0.62;
      ctx.save();
      ctx.translate(cx, cy);
      ctx.rotate(this.t * 9);
      ctx.fillStyle = C.danger;
      ctx.beginPath();
      for (let i = 0; i < 8; i++) {
        const a = (i / 8) * Math.PI * 2;
        ctx.lineTo(Math.cos(a) * r, Math.sin(a) * r);
        const b = a + Math.PI / 8;
        ctx.lineTo(Math.cos(b) * r * 0.62, Math.sin(b) * r * 0.62);
      }
      ctx.closePath();
      ctx.fill();
      ctx.fillStyle = C.dangerDeep;
      ctx.beginPath();
      ctx.arc(0, 0, r * 0.3, 0, Math.PI * 2);
      ctx.fill();
      ctx.restore();
      return;
    }
    const g = ctx.createLinearGradient(0, e.y, 0, e.y + e.h);
    g.addColorStop(0, e.style === 'loose' ? '#5a4a2e' : '#2f3a72');
    g.addColorStop(1, e.style === 'loose' ? '#3a2f1d' : '#1d2450');
    ctx.fillStyle = g;
    roundRect(ctx, e.x, e.y, e.w, e.h, 3);
    ctx.fill();
    ctx.fillStyle = e.style === 'loose' ? C.gold : C.cyan;
    ctx.fillRect(e.x + 2, e.y, e.w - 4, 1.8);
  }

  drawDoor(e) {
    const ctx = this.ctx;
    ctx.fillStyle = '#39407a';
    roundRect(ctx, e.x, e.y, e.w, e.h, 2);
    ctx.fill();
    ctx.strokeStyle = C.cyanDim;
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = 'rgba(63,232,255,0.35)';
    for (let i = 1; i < 4; i++) {
      ctx.fillRect(e.x + 2, e.y + (e.h * i) / 4, e.w - 4, 1);
    }
  }

  drawCrusher(e) {
    const ctx = this.ctx;
    const g = ctx.createLinearGradient(0, e.y, 0, e.y + e.h);
    g.addColorStop(0, '#4a3050');
    g.addColorStop(1, '#2a1b30');
    ctx.fillStyle = g;
    roundRect(ctx, e.x, e.y, e.w, e.h, 3);
    ctx.fill();
    // Bandes d'avertissement : elles virent au rouge pendant la frappe.
    const warn = e.phase === 'strike' || e.phase === 'hold';
    ctx.fillStyle = warn ? C.alert : C.danger;
    ctx.save();
    ctx.beginPath();
    ctx.rect(e.x, e.y + e.h - 5, e.w, 5);
    ctx.clip();
    for (let i = -1; i < e.w / 8 + 1; i++) {
      ctx.beginPath();
      ctx.moveTo(e.x + i * 8, e.y + e.h);
      ctx.lineTo(e.x + i * 8 + 4, e.y + e.h - 5);
      ctx.lineTo(e.x + i * 8 + 8, e.y + e.h - 5);
      ctx.lineTo(e.x + i * 8 + 4, e.y + e.h);
      ctx.fill();
    }
    ctx.restore();
  }

  drawZap(e) {
    const ctx = this.ctx;
    if (!e.deadly) {
      ctx.globalAlpha = 0.28;
    }
    const up = e.dir !== 'down';
    ctx.fillStyle = C.dangerDeep;
    roundRect(ctx, e.x, e.y, e.w, e.h, 1.5);
    ctx.fill();
    // Arcs électriques : déterministes en position, animés en intensité.
    ctx.strokeStyle = C.danger;
    ctx.lineWidth = 1.4;
    ctx.beginPath();
    const n = Math.max(2, Math.floor(e.w / 9));
    for (let i = 0; i <= n; i++) {
      const x = e.x + (e.w * i) / n;
      const wob = Math.sin(this.t * 14 + i * 1.7) * 2.4;
      const y = up ? e.y - 1 : e.y + e.h + 1;
      i === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y + (up ? -Math.abs(wob) : Math.abs(wob)));
    }
    ctx.stroke();
    ctx.globalAlpha = 1;
  }

  drawLaser(e) {
    const ctx = this.ctx;
    // Émetteur
    ctx.fillStyle = '#3a4380';
    roundRect(ctx, e.x, e.y, e.w, e.h, 2);
    ctx.fill();
    const lens = e.on ? C.danger : (e.charging ? C.alert : '#5a63a8');
    ctx.fillStyle = lens;
    ctx.beginPath();
    ctx.arc(e.x + e.w / 2, e.y + e.h / 2, 2.4, 0, Math.PI * 2);
    ctx.fill();

    // Pré-signal : le point rouge annonce le tir. C'est ce qui fait la
    // différence entre un piège et une injustice.
    if (e.charging && !e.on) {
      ctx.save();
      ctx.globalAlpha = 0.5 + Math.sin(this.t * 30) * 0.3;
      ctx.strokeStyle = C.alert;
      ctx.setLineDash([3, 5]);
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(e.beam.x, e.beam.y + e.beam.h / 2);
      ctx.lineTo(e.beam.x + e.beam.w, e.beam.y + e.beam.h / 2);
      ctx.stroke();
      ctx.restore();
    }

    if (!e.on) return;
    const b = e.beam;
    ctx.save();
    ctx.shadowColor = C.danger;
    ctx.shadowBlur = 12;
    ctx.fillStyle = 'rgba(255,138,61,0.35)';
    ctx.fillRect(b.x, b.y, b.w, b.h);
    ctx.fillStyle = '#fff3e6';
    const core = 1.6;
    if (b.w > b.h) ctx.fillRect(b.x, b.y + b.h / 2 - core / 2, b.w, core);
    else ctx.fillRect(b.x + b.w / 2 - core / 2, b.y, core, b.h);
    ctx.restore();
  }

  drawExit(e) {
    const ctx = this.ctx;
    const cx = e.x + e.w / 2;
    // Une fausse sortie a un halo qui vacille très légèrement. Invisible la
    // première fois, évident la deuxième : c'est exactement le contrat.
    const unstable = e.fake ? 0.16 * Math.sin(this.t * 17.3) + 0.08 * Math.sin(this.t * 6.1) : 0;
    const open = e.requires ? e.open : 1;
    ctx.save();
    ctx.shadowColor = C.cyan;
    ctx.shadowBlur = 16 + Math.sin(this.t * 2.4) * 4;
    ctx.fillStyle = `rgba(63,232,255,${(0.10 + open * 0.16 + unstable).toFixed(3)})`;
    roundRect(ctx, e.x - 3, e.y - 3, e.w + 6, e.h + 6, 6);
    ctx.fill();
    ctx.restore();

    ctx.fillStyle = '#0d1638';
    roundRect(ctx, e.x, e.y, e.w, e.h, 4);
    ctx.fill();
    ctx.strokeStyle = open > 0.5 ? C.cyan : C.alert;
    ctx.lineWidth = 1.6;
    ctx.stroke();

    // Trait lumineux qui monte dans l'encadrement
    const p = (this.t * 0.45 + (e.x % 7) / 7) % 1;
    ctx.save();
    ctx.globalAlpha = 0.75 * open;
    ctx.fillStyle = C.cyan;
    ctx.fillRect(e.x + 3, e.y + e.h - p * e.h, e.w - 6, 2);
    ctx.restore();

    ctx.fillStyle = `rgba(234,246,255,${0.5 + open * 0.4})`;
    ctx.font = '600 7px system-ui, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('SORTIE', cx, e.y - 6);
  }

  drawButton(e) {
    const ctx = this.ctx;
    const dz = e.press * 2.2;
    ctx.fillStyle = '#2a2f5e';
    roundRect(ctx, e.x - 2, e.y + e.h - 1, e.w + 4, 4, 1.5);
    ctx.fill();
    ctx.fillStyle = e.latched || e.pressed ? C.cyan : '#8f9ad6';
    roundRect(ctx, e.x, e.y + dz, e.w, e.h, 2);
    ctx.fill();
  }

  drawCheck(e) {
    const ctx = this.ctx;
    ctx.fillStyle = '#2a2f5e';
    ctx.fillRect(e.x + e.w / 2 - 1.2, e.y, 2.4, e.h);
    const on = e.taken;
    ctx.save();
    if (on) { ctx.shadowColor = C.cyan; ctx.shadowBlur = 10; }
    ctx.fillStyle = on ? C.cyan : '#6f79b8';
    ctx.beginPath();
    ctx.moveTo(e.x + e.w / 2 + 1, e.y + 1);
    ctx.lineTo(e.x + e.w / 2 + 10, e.y + 4.5);
    ctx.lineTo(e.x + e.w / 2 + 1, e.y + 8);
    ctx.closePath();
    ctx.fill();
    ctx.restore();
  }

  drawTele(e) {
    const ctx = this.ctx;
    const broken = e.spec.broken === true;
    const col = broken ? '#8a5cff' : C.cyan;
    ctx.save();
    ctx.shadowColor = col;
    ctx.shadowBlur = broken ? 5 : 12;
    ctx.fillStyle = broken ? 'rgba(138,92,255,0.25)' : 'rgba(63,232,255,0.3)';
    roundRect(ctx, e.x, e.y, e.w, e.h, 3);
    ctx.fill();
    ctx.restore();
    ctx.strokeStyle = col;
    ctx.lineWidth = 1.2;
    ctx.setLineDash(broken ? [3, 3] : []);
    roundRect(ctx, e.x, e.y, e.w, e.h, 3);
    ctx.stroke();
    ctx.setLineDash([]);
  }

  drawGrav(e) {
    const ctx = this.ctx;
    ctx.save();
    ctx.globalAlpha = 0.22 + Math.sin(this.t * 3) * 0.05;
    const g = ctx.createLinearGradient(e.x, 0, e.x + e.w, 0);
    g.addColorStop(0, 'rgba(123,92,255,0)');
    g.addColorStop(0.5, 'rgba(163,120,255,0.85)');
    g.addColorStop(1, 'rgba(123,92,255,0)');
    ctx.fillStyle = g;
    ctx.fillRect(e.x, e.y, e.w, e.h);
    ctx.restore();
    // Chevrons qui défilent : ils indiquent le sens de renversement.
    ctx.strokeStyle = 'rgba(200,180,255,0.8)';
    ctx.lineWidth = 1.2;
    const cx = e.x + e.w / 2;
    for (let i = 0; i < 4; i++) {
      const p = ((this.t * 0.6 + i / 4) % 1);
      const y = e.y + e.h - p * e.h;
      ctx.globalAlpha = Math.sin(p * Math.PI);
      ctx.beginPath();
      ctx.moveTo(cx - 5, y + 4);
      ctx.lineTo(cx, y);
      ctx.lineTo(cx + 5, y + 4);
      ctx.stroke();
    }
    ctx.globalAlpha = 1;
  }

  drawFan(e) {
    const ctx = this.ctx;
    ctx.save();
    ctx.globalAlpha = e.on ? 0.20 : 0.05;
    const vertical = e.dir === 'up' || e.dir === 'down';
    const g = vertical
      ? ctx.createLinearGradient(0, e.y, 0, e.y + e.h)
      : ctx.createLinearGradient(e.x, 0, e.x + e.w, 0);
    const a = 'rgba(63,232,255,0.55)';
    const b = 'rgba(63,232,255,0)';
    g.addColorStop(0, e.dir === 'up' || e.dir === 'left' ? b : a);
    g.addColorStop(1, e.dir === 'up' || e.dir === 'left' ? a : b);
    ctx.fillStyle = g;
    ctx.fillRect(e.x, e.y, e.w, e.h);
    ctx.restore();
    if (!e.on) return;
    ctx.strokeStyle = 'rgba(180,240,255,0.55)';
    ctx.lineWidth = 1;
    for (let i = 0; i < 5; i++) {
      const p = ((this.t * 1.3 + i / 5) % 1);
      ctx.globalAlpha = Math.sin(p * Math.PI) * 0.8;
      ctx.beginPath();
      if (vertical) {
        const y = e.dir === 'up' ? e.y + e.h - p * e.h : e.y + p * e.h;
        const x = e.x + ((i * 37) % e.w);
        ctx.moveTo(x, y);
        ctx.lineTo(x, y + (e.dir === 'up' ? 8 : -8));
      } else {
        const x = e.dir === 'left' ? e.x + e.w - p * e.w : e.x + p * e.w;
        const y = e.y + ((i * 29) % e.h);
        ctx.moveTo(x, y);
        ctx.lineTo(x + (e.dir === 'left' ? 8 : -8), y);
      }
      ctx.stroke();
    }
    ctx.globalAlpha = 1;
  }

  drawField(e) {
    const ctx = this.ctx;
    // La couleur dit la vérité : un champ qui attire vers un danger est
    // orange, un champ inoffensif est cyan. Le joueur peut donc décider AVANT
    // d'y entrer — la surprise vient d'ailleurs.
    const col = e.kind === 'magnet'
      ? (e.spec.hostile ? 'rgba(255,138,61,' : 'rgba(63,232,255,')
      : e.kind === 'fast' ? 'rgba(255,209,102,' : e.kind === 'ice' ? 'rgba(150,220,255,' : 'rgba(123,92,255,';
    ctx.save();
    ctx.fillStyle = col + '0.10)';
    ctx.fillRect(e.x, e.y, e.w, e.h);
    ctx.strokeStyle = col + '0.45)';
    ctx.setLineDash([5, 5]);
    ctx.lineWidth = 1;
    ctx.strokeRect(e.x + 0.5, e.y + 0.5, e.w - 1, e.h - 1);
    ctx.setLineDash([]);
    if (e.kind === 'magnet') {
      const ax = e.anchor ? e.spec.anchor[0] * TILE : e.x + e.w / 2;
      const ay = e.anchor ? e.spec.anchor[1] * TILE : e.y + e.h / 2;
      ctx.strokeStyle = col + '0.5)';
      for (let i = 0; i < 3; i++) {
        const r = 8 + ((this.t * 22 + i * 16) % 46);
        ctx.globalAlpha = 1 - r / 46;
        ctx.beginPath();
        ctx.arc(ax, ay, r, 0, Math.PI * 2);
        ctx.stroke();
      }
    }
    ctx.restore();
  }

  drawSign(e) {
    const ctx = this.ctx;
    const v = e.vis;
    if (v < 0.03) return;
    ctx.save();
    ctx.globalAlpha = v;
    ctx.fillStyle = 'rgba(8,12,32,0.82)';
    roundRect(ctx, e.x, e.y, e.w, e.h, 3);
    ctx.fill();
    ctx.strokeStyle = 'rgba(63,232,255,0.5)';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = C.cyan;
    ctx.font = '600 6.5px system-ui, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const lines = e.text.split('\n');
    lines.forEach((ln, i) => {
      ctx.fillText(ln, e.x + e.w / 2, e.y + e.h / 2 + (i - (lines.length - 1) / 2) * 8);
    });
    ctx.restore();
  }

  drawDeco(e) {
    const ctx = this.ctx;
    switch (e.kind) {
      case 'lamp': {
        const on = e.spec.signal ? e.world.isOn(e.spec.signal) : false;
        ctx.save();
        if (on) { ctx.shadowColor = '#5dff9b'; ctx.shadowBlur = 12; }
        ctx.fillStyle = on ? '#5dff9b' : '#4a2030';
        ctx.beginPath();
        ctx.arc(e.x + e.w / 2, e.y + e.h / 2, e.w / 2, 0, Math.PI * 2);
        ctx.fill();
        ctx.restore();
        break;
      }
      case 'crate':
        ctx.fillStyle = '#33397a';
        roundRect(ctx, e.x, e.y, e.w, e.h, 2);
        ctx.fill();
        ctx.strokeStyle = '#5a63a8';
        ctx.lineWidth = 1;
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(e.x, e.y);
        ctx.lineTo(e.x + e.w, e.y + e.h);
        ctx.moveTo(e.x + e.w, e.y);
        ctx.lineTo(e.x, e.y + e.h);
        ctx.stroke();
        break;
      case 'bolt':
        ctx.fillStyle = C.gold;
        ctx.beginPath();
        ctx.arc(e.x + e.w / 2, e.y + e.h / 2, e.w / 2, 0, Math.PI * 2);
        ctx.fill();
        break;
      case 'stripe': {
        ctx.save();
        ctx.beginPath();
        ctx.rect(e.x, e.y, e.w, e.h);
        ctx.clip();
        ctx.fillStyle = C.danger;
        for (let i = -1; i < e.w / 8 + 1; i++) {
          ctx.beginPath();
          ctx.moveTo(e.x + i * 8, e.y + e.h);
          ctx.lineTo(e.x + i * 8 + 4, e.y);
          ctx.lineTo(e.x + i * 8 + 8, e.y);
          ctx.lineTo(e.x + i * 8 + 4, e.y + e.h);
          ctx.fill();
        }
        ctx.restore();
        break;
      }
      case 'grid':
        ctx.strokeStyle = '#5a63a8';
        ctx.lineWidth = 1;
        for (let x = e.x + 2; x < e.x + e.w; x += 4) {
          ctx.beginPath();
          ctx.moveTo(x, e.y);
          ctx.lineTo(x, e.y + e.h);
          ctx.stroke();
        }
        break;
      case 'pipe':
      case 'rail':
        ctx.fillStyle = '#2b3160';
        roundRect(ctx, e.x, e.y, e.w, e.h, e.h / 2);
        ctx.fill();
        break;
      case 'vent':
        ctx.fillStyle = '#1c2148';
        roundRect(ctx, e.x, e.y, e.w, e.h, 2);
        ctx.fill();
        ctx.strokeStyle = '#4a5290';
        for (let i = 1; i < 4; i++) {
          ctx.beginPath();
          ctx.moveTo(e.x + 2, e.y + (e.h * i) / 4);
          ctx.lineTo(e.x + e.w - 2, e.y + (e.h * i) / 4);
          ctx.stroke();
        }
        break;
      case 'gift':
      case 'core': {
        ctx.save();
        ctx.shadowColor = C.gold;
        ctx.shadowBlur = 10 + Math.sin(this.t * 3) * 4;
        ctx.fillStyle = C.gold;
        const cx = e.x + e.w / 2;
        const cy = e.y + e.h / 2 + Math.sin(this.t * 2) * 1.5;
        const r = Math.min(e.w, e.h) * 0.36;
        ctx.beginPath();
        for (let i = 0; i < 6; i++) {
          const a = (i / 6) * Math.PI * 2 + this.t * 0.6;
          ctx.lineTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r);
        }
        ctx.closePath();
        ctx.fill();
        ctx.restore();
        break;
      }
      case 'dust': {
        ctx.save();
        ctx.fillStyle = 'rgba(200,225,255,0.5)';
        for (let i = 0; i < 5; i++) {
          const p = ((this.t * 0.25 + i / 5) % 1);
          const x = e.x + ((i * 53) % e.w);
          const y = e.y + e.h - p * e.h;
          ctx.globalAlpha = Math.sin(p * Math.PI) * 0.55;
          ctx.fillRect(x, y, 1.3, 1.3);
        }
        ctx.restore();
        break;
      }
      default:
        ctx.fillStyle = '#252b58';
        roundRect(ctx, e.x, e.y, e.w, e.h, 2);
        ctx.fill();
        ctx.strokeStyle = '#454d90';
        ctx.lineWidth = 1;
        ctx.stroke();
    }
  }

  // --- Boucle principale -----------------------------------------------------

  drawWorld(world, cam, view, skin, dt, opts = {}) {
    const ctx = this.ctx;
    this.t += dt;
    const pal = palette(world.level.palette);
    this.drawBackground(view.w, view.h, cam, pal);

    ctx.save();
    ctx.translate(Math.round(-cam.x + cam.shakeX), Math.round(-cam.y + cam.shakeY));

    const L = cam.x - 40, R = cam.x + view.w + 40;
    const T = cam.y - 40, B = cam.y + view.h + 40;
    const visible = (e) => e.x + e.w > L && e.x < R && e.y + e.h > T && e.y < B;

    // Deux passes : décor et champs derrière, structures et dangers devant.
    for (const e of world.entities) {
      if (!visible(e)) continue;
      switch (e.type) {
        case 'deco': this.drawDeco(e); break;
        case 'field': this.drawField(e); break;
        case 'fan': this.drawFan(e); break;
        case 'grav': this.drawGrav(e); break;
        case 'sign': this.drawSign(e); break;
        default: break;
      }
    }
    for (const e of world.entities) {
      if (!visible(e)) continue;
      switch (e.type) {
        case 'solid': this.drawSolid(e); break;
        case 'crumble': this.drawCrumble(e); break;
        case 'vanish': this.drawVanish(e); break;
        case 'ghost': this.drawGhost(e); break;
        case 'mover': this.drawMover(e); break;
        case 'door': this.drawDoor(e); break;
        case 'crusher': this.drawCrusher(e); break;
        case 'zap': this.drawZap(e); break;
        case 'laser': this.drawLaser(e); break;
        case 'exit': this.drawExit(e); break;
        case 'button': this.drawButton(e); break;
        case 'check': this.drawCheck(e); break;
        case 'tele': this.drawTele(e); break;
        case 'clone':
          if (e.alpha > 0.02) {
            drawRobot(ctx, {
              x: e.x + e.w / 2, y: e.y + e.h / 2, skin: { ...skin, shell: 'rgba(255,120,150,0.5)', shell2: 'rgba(180,60,90,0.45)', eye: C.alert, trim: C.alert, screen: 'rgba(20,4,10,0.6)' },
              state: 'run', facing: e.facing, squash: 0, t: this.t, alpha: e.alpha,
            });
          }
          break;
        default: break;
      }
    }

    opts.particles?.draw(ctx);

    // Hors-salle : quand la pièce est plus étroite que l'écran, on masque
    // franchement l'extérieur. Sans ça, le sol « s'arrête dans le vide » et la
    // salle a l'air inachevée.
    ctx.fillStyle = 'rgba(3,4,12,0.86)';
    if (cam.x < 0) ctx.fillRect(cam.x - 200, cam.y - 200, 200 - cam.x, view.h + 400);
    if (cam.x + view.w > world.w) {
      ctx.fillRect(world.w, cam.y - 200, cam.x + view.w - world.w + 200, view.h + 400);
    }
    if (cam.y + view.h > world.h) {
      ctx.fillRect(cam.x - 200, world.h, view.w + 400, cam.y + view.h - world.h + 200);
    }

    // Le robot, toujours au premier plan.
    const p = world.player;
    if (world.state !== 'dead') {
      drawRobot(ctx, {
        x: p.cx, y: p.cy, skin,
        state: world.state === 'win' ? 'win' : p.anim,
        facing: p.facing,
        squash: p.squash * p.gravDir,
        t: this.t,
        roll: p.gravDir < 0 ? Math.PI : 0,
        groundY: p.grounded ? (p.gravDir > 0 ? p.y + p.h + 1 : null) : null,
      });
    }

    ctx.restore();
    this.drawVignette(view);
  }

  drawVignette(view) {
    const ctx = this.ctx;
    const g = ctx.createRadialGradient(
      view.w / 2, view.h / 2, view.h * 0.34,
      view.w / 2, view.h / 2, view.h * 0.92,
    );
    g.addColorStop(0, 'rgba(0,0,0,0)');
    g.addColorStop(1, 'rgba(0,0,0,0.42)');
    ctx.fillStyle = g;
    ctx.fillRect(0, 0, view.w, view.h);
  }
}
