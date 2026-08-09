import { CHAPTER_1 } from './chapter1.js';
import { CHAPTER_2 } from './chapter2.js';
import { CHAPTER_3 } from './chapter3.js';
import { CHAPTER_4 } from './chapter4.js';
import { CHAPTER_5 } from './chapter5.js';

export const CHAPTERS = [CHAPTER_1, CHAPTER_2, CHAPTER_3, CHAPTER_4, CHAPTER_5];

/**
 * Normalisation : c'est ici, et nulle part ailleurs, qu'un niveau écrit à la
 * main devient un niveau exécutable.
 *
 * Ajouter les murs de bordure automatiquement évite de répéter trois entités
 * dans chacun des trente niveaux — et surtout évite l'oubli, qui produirait un
 * bug de sortie de salle très difficile à diagnostiquer.
 */
function normalize(level, chapter, indexInChapter, globalIndex) {
  const w = level.w ?? 26;
  const h = level.h ?? 15;
  return {
    ...level,
    w, h,
    id: `c${chapter.id}-${String(indexInChapter + 1).padStart(2, '0')}`,
    num: globalIndex + 1,
    chapterId: chapter.id,
    chapterName: chapter.name,
    palette: chapter.palette,
    dash: chapter.dash === true,
    entities: [
      // Murs latéraux : on ne sort jamais d'une salle par le côté.
      { t: 'solid', x: -2, y: -30, w: 2, h: h + 60, style: 'bound' },
      { t: 'solid', x: w, y: -30, w: 2, h: h + 60, style: 'bound' },
      ...level.entities,
    ],
  };
}

/** Liste à plat des 30 niveaux, dans l'ordre de progression. */
export const LEVELS = (() => {
  const out = [];
  for (const ch of CHAPTERS) {
    ch.levels.forEach((lv, i) => out.push(normalize(lv, ch, i, out.length)));
  }
  return out;
})();

export const LEVEL_COUNT = LEVELS.length;

export function getLevel(num) {
  return LEVELS[num - 1] || null;
}

export function chapterOf(num) {
  const lv = getLevel(num);
  return lv ? CHAPTERS.find((c) => c.id === lv.chapterId) : null;
}

/** Index (base 1) du premier niveau de chaque chapitre. */
export function chapterRanges() {
  const ranges = [];
  let start = 1;
  for (const ch of CHAPTERS) {
    ranges.push({ chapter: ch, from: start, to: start + ch.levels.length - 1 });
    start += ch.levels.length;
  }
  return ranges;
}
