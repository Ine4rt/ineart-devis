import { C, roundRect } from './theme.js';

/**
 * Le robot.
 *
 * Entièrement vectoriel, dessiné image par image : aucun sprite, aucune
 * feuille d'animation. C'est ce qui permet huit skins pour zéro octet d'assets
 * et une netteté parfaite de 480p à 1440p.
 *
 * Toute l'expressivité passe par UN écran facial. Le corps ne fait que se
 * déformer (squash & stretch) ; c'est le visage qui raconte. Un robot qui a
 * l'air surpris une demi-seconde avant de mourir transforme une punition en
 * gag — et un gag, on veut le revoir.
 */

const EYES = {
  //         forme      espacement  hauteur
  normal: { shape: 'dot', gap: 5.2, y: 0, size: 2.3 },
  happy: { shape: 'arc', gap: 5.4, y: 0.5, size: 2.6 },
  surprise: { shape: 'ring', gap: 5.6, y: -0.2, size: 3.0 },
  focus: { shape: 'bar', gap: 5.2, y: 0.2, size: 2.2 },
  worry: { shape: 'small', gap: 5.0, y: 0.6, size: 1.7 },
  dead: { shape: 'cross', gap: 5.4, y: 0, size: 2.8 },
  zap: { shape: 'spiral', gap: 5.4, y: 0, size: 2.7 },
};

/** Quel visage pour quel état ? C'est toute la mise en scène du personnage. */
export function faceFor(state) {
  switch (state) {
    case 'run': return 'focus';
    case 'jump': return 'surprise';
    case 'fall': return 'worry';
    case 'dash': return 'focus';
    case 'win': return 'happy';
    case 'dead': return 'dead';
    case 'zap': return 'zap';
    default: return 'normal';
  }
}

function drawEye(ctx, x, y, e, color, blink) {
  ctx.fillStyle = color;
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.4;
  ctx.lineCap = 'round';
  if (blink) {
    ctx.beginPath();
    ctx.moveTo(x - e.size, y);
    ctx.lineTo(x + e.size, y);
    ctx.stroke();
    return;
  }
  switch (e.shape) {
    case 'arc':
      ctx.beginPath();
      ctx.arc(x, y + e.size * 0.5, e.size, Math.PI * 1.15, Math.PI * 1.85);
      ctx.stroke();
      break;
    case 'ring':
      ctx.beginPath();
      ctx.arc(x, y, e.size, 0, Math.PI * 2);
      ctx.stroke();
      break;
    case 'bar':
      roundRect(ctx, x - e.size * 0.75, y - e.size * 0.5, e.size * 1.5, e.size, 1);
      ctx.fill();
      break;
    case 'small':
      ctx.beginPath();
      ctx.arc(x, y, e.size * 0.8, 0, Math.PI * 2);
      ctx.fill();
      break;
    case 'cross':
      ctx.beginPath();
      ctx.moveTo(x - e.size, y - e.size);
      ctx.lineTo(x + e.size, y + e.size);
      ctx.moveTo(x + e.size, y - e.size);
      ctx.lineTo(x - e.size, y + e.size);
      ctx.stroke();
      break;
    case 'spiral':
      ctx.beginPath();
      for (let a = 0; a < Math.PI * 3.2; a += 0.28) {
        const r = (a / (Math.PI * 3.2)) * e.size;
        const px = x + Math.cos(a) * r;
        const py = y + Math.sin(a) * r;
        a === 0 ? ctx.moveTo(px, py) : ctx.lineTo(px, py);
      }
      ctx.stroke();
      break;
    default:
      ctx.beginPath();
      ctx.arc(x, y, e.size, 0, Math.PI * 2);
      ctx.fill();
  }
}

/**
 * @param {CanvasRenderingContext2D} ctx
 * @param {object} o { x, y (centre), skin, state, facing, squash, t, alpha }
 */
export function drawRobot(ctx, o) {
  const skin = o.skin;
  const scale = (skin.scale ?? 1) * (o.scale ?? 1);
  const R = 11 * scale;
  const sq = o.squash || 0;
  // Squash & stretch : volume constant, sinon le robot « respire » faux.
  const sx = 1 + sq * 0.35;
  const sy = 1 - sq * 0.35;
  const face = EYES[faceFor(o.state)] || EYES.normal;
  const t = o.t || 0;
  // Clignement : toutes les ~3,4 s, très court. Détail invisible mais c'est ce
  // qui fait qu'on trouve le robot « vivant ».
  const blink = o.state !== 'dead' && (t % 3.4) < 0.11;

  ctx.save();
  ctx.globalAlpha = o.alpha ?? 1;
  ctx.translate(o.x, o.y);
  if (o.roll) ctx.rotate(o.roll);
  ctx.scale(sx * (o.facing < 0 ? -1 : 1), sy);

  // Ombre portée douce au sol
  if (o.groundY != null) {
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.globalAlpha = 0.25 * (o.alpha ?? 1);
    ctx.fillStyle = '#000';
    ctx.beginPath();
    ctx.ellipse(o.x, o.groundY, R * 0.9, R * 0.28, 0, 0, Math.PI * 2);
    ctx.fill();
    ctx.restore();
  }

  // Corps
  const g = ctx.createLinearGradient(0, -R, 0, R);
  g.addColorStop(0, skin.shell);
  g.addColorStop(1, skin.shell2);
  ctx.fillStyle = g;
  ctx.beginPath();
  ctx.arc(0, 0, R, 0, Math.PI * 2);
  ctx.fill();

  if (skin.stripes) {
    ctx.save();
    ctx.clip();
    ctx.fillStyle = 'rgba(30,30,30,0.55)';
    for (let i = -2; i <= 2; i++) {
      ctx.save();
      ctx.rotate(-0.6);
      ctx.fillRect(i * 5 * scale - 1.2, -R * 1.4, 2.4, R * 2.8);
      ctx.restore();
    }
    ctx.restore();
  }

  // Liseré lumineux
  ctx.strokeStyle = skin.trim;
  ctx.globalAlpha = (o.alpha ?? 1) * 0.85;
  ctx.lineWidth = 1.3;
  ctx.beginPath();
  ctx.arc(0, 0, R - 0.7, 0, Math.PI * 2);
  ctx.stroke();
  ctx.globalAlpha = o.alpha ?? 1;

  // Antenne
  if (skin.antenna && skin.antenna !== 'none') {
    ctx.strokeStyle = skin.shell2;
    ctx.lineWidth = 1.6;
    const bob = Math.sin(t * 6) * 0.8 * (o.state === 'run' ? 1 : 0.25);
    if (skin.antenna === 'twin') {
      for (const s of [-1, 1]) {
        ctx.beginPath();
        ctx.moveTo(s * 4 * scale, -R + 1);
        ctx.lineTo(s * 6 * scale, -R - 4 * scale + bob);
        ctx.stroke();
        ctx.fillStyle = skin.trim;
        ctx.beginPath();
        ctx.arc(s * 6 * scale, -R - 4 * scale + bob, 1.6 * scale, 0, Math.PI * 2);
        ctx.fill();
      }
    } else {
      const len = skin.antenna === 'rod' ? 6 : 3;
      ctx.beginPath();
      ctx.moveTo(0, -R + 1);
      ctx.lineTo(bob, -R - len * scale);
      ctx.stroke();
      ctx.fillStyle = skin.trim;
      ctx.beginPath();
      ctx.arc(bob, -R - len * scale, 1.8 * scale, 0, Math.PI * 2);
      ctx.fill();
    }
  }

  // Écran facial
  const sw = R * 1.35;
  const sh = R * 0.95;
  ctx.fillStyle = skin.screen;
  roundRect(ctx, -sw / 2, -sh / 2 - R * 0.06, sw, sh, 3.4 * scale);
  ctx.fill();

  // Yeux
  ctx.save();
  ctx.shadowColor = skin.eye;
  ctx.shadowBlur = 6;
  const ey = face.y * scale - R * 0.06;
  drawEye(ctx, -face.gap * 0.5 * scale, ey, { ...face, size: face.size * scale }, skin.eye, blink);
  drawEye(ctx, face.gap * 0.5 * scale, ey, { ...face, size: face.size * scale }, skin.eye, blink);
  ctx.restore();

  // Bouche : une seule courbe, seulement quand elle apporte quelque chose.
  if (o.state === 'win' || o.state === 'fall') {
    ctx.strokeStyle = skin.eye;
    ctx.lineWidth = 1.1;
    ctx.beginPath();
    const my = ey + 4.2 * scale;
    if (o.state === 'win') ctx.arc(0, my - 1.5 * scale, 2.4 * scale, 0.15 * Math.PI, 0.85 * Math.PI);
    else ctx.arc(0, my + 1.6 * scale, 2.2 * scale, 1.15 * Math.PI, 1.85 * Math.PI);
    ctx.stroke();
  }

  if (skin.cracked) {
    ctx.strokeStyle = 'rgba(255,255,255,0.55)';
    ctx.lineWidth = 0.8;
    ctx.beginPath();
    ctx.moveTo(-sw * 0.4, -sh * 0.4);
    ctx.lineTo(-sw * 0.1, 0);
    ctx.lineTo(-sw * 0.25, sh * 0.35);
    ctx.moveTo(-sw * 0.1, 0);
    ctx.lineTo(sw * 0.35, -sh * 0.15);
    ctx.stroke();
  }

  if (skin.shine) {
    ctx.globalAlpha = (o.alpha ?? 1) * 0.5;
    ctx.fillStyle = '#fff';
    ctx.beginPath();
    ctx.ellipse(-R * 0.42, -R * 0.5, R * 0.22, R * 0.12, -0.7, 0, Math.PI * 2);
    ctx.fill();
  }

  ctx.restore();
}

/** Portrait statique pour les menus (sélection de robot). */
export function drawRobotPortrait(ctx, x, y, skin, scale, t, state = 'idle') {
  drawRobot(ctx, { x, y, skin, state, facing: 1, squash: Math.sin(t * 2) * 0.05, t, scale });
}

export const ROBOT_COLORS = C;
