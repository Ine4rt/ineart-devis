/**
 * Robots déblocables — purement cosmétiques.
 *
 * Règle non négociable : aucun skin ne modifie une statistique, une hitbox ou
 * une capacité. Un jeu de précision qui vend de l'avantage cosmétique perd le
 * droit de dire à son joueur « c'est ta faute ».
 *
 * Chaque skin est décrit par des couleurs et deux ou trois détails de forme :
 * tout est dessiné en vectoriel à l'exécution, il n'y a aucun asset image dans
 * le projet (chargement instantané, poids nul, netteté à toute résolution).
 */

export const SKINS = [
  {
    id: 'standard',
    name: 'UNITÉ-01',
    desc: 'Sortie d\'usine. Fiable. Optimiste.',
    shell: '#eaf6ff', shell2: '#b9d4e8', screen: '#0d1230', eye: '#3fe8ff',
    trim: '#3fe8ff', antenna: 'nub',
    unlock: null,
  },
  {
    id: 'retro',
    name: 'VX-70',
    desc: 'Rétro-futuriste. Écran cathodique, fierté intacte.',
    shell: '#f2e4c8', shell2: '#cdb894', screen: '#161005', eye: '#ffb547',
    trim: '#ffb547', antenna: 'rod',
    unlock: { type: 'deaths', value: 100, label: '100 morts' },
  },
  {
    id: 'mini',
    name: 'PUCE',
    desc: 'Plus petit à l\'écran. Exactement la même hitbox.',
    shell: '#dff5ff', shell2: '#9fd0e5', screen: '#08122a', eye: '#7bffde',
    trim: '#7bffde', antenna: 'none', scale: 0.82,
    unlock: { type: 'fastLevel', value: 10, label: 'un niveau en moins de 10 s' },
  },
  {
    id: 'indus',
    name: 'MANUTENTION',
    desc: 'Bâti lourd. Peinture écaillée. Ne pose pas de questions.',
    shell: '#ffb43d', shell2: '#c07a17', screen: '#1a1004', eye: '#ffe9b0',
    trim: '#2a2a2a', antenna: 'rod', stripes: true,
    unlock: { type: 'levels', value: 12, label: 'terminer 12 niveaux' },
  },
  {
    id: 'gold',
    name: 'PROTOTYPE OR',
    desc: 'Coûteux. Inutile. Magnifique.',
    shell: '#ffd166', shell2: '#c79a2b', screen: '#241704', eye: '#fff3c4',
    trim: '#fff3c4', antenna: 'nub', shine: true,
    unlock: { type: 'noDeathStreak', value: 5, label: '5 niveaux d\'affilée sans mourir' },
  },
  {
    id: 'ghost',
    name: 'ÉCHO',
    desc: 'Techniquement, cette unité a été perdue au secteur C.',
    shell: 'rgba(190,230,255,0.45)', shell2: 'rgba(120,180,220,0.35)',
    screen: 'rgba(10,16,40,0.6)', eye: '#bff3ff', trim: '#bff3ff',
    antenna: 'none', translucent: true,
    unlock: { type: 'chapter', value: 5, label: 'atteindre le secteur interdit' },
  },
  {
    id: 'alien',
    name: 'VISITEUR',
    desc: 'Origine inconnue. Ne figure sur aucun manifeste.',
    shell: '#9dffb0', shell2: '#4fb56a', screen: '#04180c', eye: '#eaffef',
    trim: '#eaffef', antenna: 'twin',
    unlock: { type: 'secrets', value: 3, label: 'trouver 3 salles secrètes' },
  },
  {
    id: 'cracked',
    name: 'ÉCRAN FÊLÉ',
    desc: 'Il voit encore. Enfin, en gros.',
    shell: '#dfe6f5', shell2: '#a8b2c8', screen: '#0a0d1c', eye: '#ff8a3d',
    trim: '#ff8a3d', antenna: 'nub', cracked: true,
    unlock: { type: 'deaths', value: 500, label: '500 morts' },
  },
];

export const SKIN_BY_ID = Object.fromEntries(SKINS.map((s) => [s.id, s]));

/** Un skin est-il débloqué au vu des statistiques du joueur ? */
export function isSkinUnlocked(skin, stats) {
  const u = skin.unlock;
  if (!u) return true;
  switch (u.type) {
    case 'deaths': return stats.totalDeaths >= u.value;
    case 'levels': return stats.completedCount >= u.value;
    case 'fastLevel': return stats.bestTimeAny > 0 && stats.bestTimeAny < u.value;
    case 'noDeathStreak': return stats.bestCleanStreak >= u.value;
    case 'secrets': return stats.secrets >= u.value;
    case 'chapter': return stats.maxChapter >= u.value;
    default: return false;
  }
}
