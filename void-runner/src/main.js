import { Game } from './core/Game.js';
import { UIManager } from './ui/UIManager.js';
import { SaveManager } from './core/Save.js';
import { AudioManager } from './core/Audio.js';
import { Haptics } from './core/Haptics.js';
import { Input } from './core/Input.js';
import { AchievementManager } from './core/Achievements.js';

/**
 * Point d'entrée.
 *
 * Ordre volontaire : sauvegarde → services → jeu → interface. L'audio n'est
 * jamais initialisé ici — les navigateurs mobiles exigent un geste utilisateur,
 * et forcer un contexte audio au chargement coûte du temps de démarrage pour
 * rien. Il se déverrouille au premier appui, où qu'il soit.
 */

const canvas = document.getElementById('game');
const save = new SaveManager();
const audio = new AudioManager();
const haptics = new Haptics();
const input = new Input(document.getElementById('touch'));
const achievements = new AchievementManager(save);

const game = new Game(canvas, { save, audio, haptics, input, achievements });
const ui = new UIManager(game);

input.onAnyInput = () => audio.unlock();
addEventListener('pointerdown', () => audio.unlock(), { once: true });

// Sauvegarde garantie avant fermeture ou passage en arrière-plan.
addEventListener('pagehide', () => save.flush());
document.addEventListener('visibilitychange', () => { if (document.hidden) save.flush(); });

// Le double-tap zoom d'iOS ruine un jeu d'action : on le neutralise.
document.addEventListener('dblclick', (e) => e.preventDefault(), { passive: false });
document.addEventListener('gesturestart', (e) => e.preventDefault());

// Utile pour tester une salle précise : index.html#12
const deep = Number(location.hash.replace('#', ''));
if (deep >= 1 && deep <= 30) {
  audio.unlock();
  game.startLevel(deep);
  ui.showScreen(null);
}

if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  navigator.serviceWorker.register('sw.js').catch(() => { /* hors ligne non critique */ });
}

// Exposé pour le débogage manuel depuis la console (jamais utilisé par le jeu).
window.VOID = { game, ui, save };
