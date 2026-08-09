/**
 * Répliques d'après-mort.
 *
 * Dosage : jamais à chaque mort. Une vanne systématique devient un temps de
 * chargement. On n'affiche un message que dans trois cas — une mort qui fait
 * franchir un palier rond, une mort « ironique » (très près du but), ou un
 * tirage rare. Le reste du temps, on redémarre en silence, tout de suite.
 */

const GENERIC = [
  'Presque.',
  'Tu y étais.',
  'Vraiment ?',
  'Encore une fois.',
  'Le bouton était suspect.',
  'Tu l\'avais vu, pourtant.',
  'Cette fois, ça devrait marcher.',
  'Statistiquement, ça devient inquiétant.',
  'La station note tout, tu sais.',
  'C\'était courageux.',
  'On appelle ça une hypothèse.',
  'Le sol n\'y est pour rien.',
];

const CLOSE = [
  'À deux pas.',
  'La porte t\'a vu arriver.',
  'Encore trente centimètres.',
  'Elle était juste là.',
];

const MILESTONE = {
  5: 'Cinq essais. On échauffe.',
  10: 'Dixième tentative. Le niveau apprend aussi.',
  20: 'Vingt. Tu as ma sympathie.',
  30: 'Trente. Tu as ma sympathie ET mon inquiétude.',
  50: 'Cinquante essais sur cette salle. On peut en parler.',
  100: 'Cent. Sincèrement, bravo pour la ténacité.',
};

const TOTAL_MILESTONE = {
  100: '100 morts au compteur. Un début.',
  250: '250 morts. Le rapport de maintenance est long.',
  500: '500 morts. On a arrêté d\'imprimer.',
  1000: '1 000 morts. Tu es une légende locale.',
};

/**
 * @param {object} ctx { attempt, totalDeaths, nearExit, rng }
 * @returns {string|null} message, ou null pour redémarrer en silence
 */
export function pickQuip(ctx) {
  if (TOTAL_MILESTONE[ctx.totalDeaths]) return TOTAL_MILESTONE[ctx.totalDeaths];
  if (MILESTONE[ctx.attempt]) return MILESTONE[ctx.attempt];
  const r = ctx.rng();
  if (ctx.nearExit && r < 0.55) return CLOSE[Math.floor(r * 1000) % CLOSE.length];
  // ~18 % des morts « ordinaires » sont commentées. Au-delà, ça lasse.
  if (r < 0.18) return GENERIC[Math.floor(r * 10000) % GENERIC.length];
  return null;
}

export const WIN_LINES = [
  'Propre.',
  'Voilà.',
  'C\'était pas si compliqué.',
  'La station est déçue.',
  'Salle suivante.',
  'Tu commences à comprendre.',
];
