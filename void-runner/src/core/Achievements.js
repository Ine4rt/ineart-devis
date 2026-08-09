import { SKINS, isSkinUnlocked } from '../data/skins.js';

/**
 * Accomplissements.
 *
 * Ils ne servent qu'à débloquer des robots, et ils sont conçus pour récompenser
 * DEUX profils opposés : celui qui persévère (100 morts, 500 morts) et celui
 * qui maîtrise (un niveau sous 10 s, cinq sans mourir). Un joueur qui échoue
 * beaucoup doit débloquer des choses lui aussi — sinon le jeu punit deux fois.
 */
export const ACHIEVEMENTS = [
  { id: 'first', name: 'Premier pas', desc: 'Terminer un niveau.', test: (s) => s.completedCount >= 1 },
  { id: 'd100', name: 'Persévérant', desc: '100 morts au total.', test: (s) => s.totalDeaths >= 100 },
  { id: 'd500', name: 'Increvable', desc: '500 morts au total.', test: (s) => s.totalDeaths >= 500 },
  { id: 'clean5', name: 'Sans une éraflure', desc: '5 niveaux d\'affilée sans mourir.', test: (s) => s.bestCleanStreak >= 5 },
  { id: 'fast10', name: 'Expéditif', desc: 'Un niveau en moins de 10 secondes.', test: (s) => s.bestTimeAny > 0 && s.bestTimeAny < 10 },
  { id: 'secret1', name: 'Curieux', desc: 'Trouver une salle secrète.', test: (s) => s.secrets >= 1 },
  { id: 'secret3', name: 'Fouineur', desc: 'Trouver les 3 salles secrètes.', test: (s) => s.secrets >= 3 },
  { id: 'ch2', name: 'Sortie de secteur', desc: 'Atteindre le laboratoire.', test: (s) => s.maxChapter >= 2 },
  { id: 'ch5', name: 'Zone interdite', desc: 'Atteindre le secteur interdit.', test: (s) => s.maxChapter >= 5 },
  { id: 'all', name: 'Extraction complète', desc: 'Terminer les 30 salles.', test: (s) => s.completedCount >= s.levelCount },
];

export class AchievementManager {
  constructor(save) {
    this.save = save;
  }

  /**
   * Réévalue tout à chaque appel (c'est trivialement rapide) et renvoie la
   * liste des nouveautés — accomplissements ET robots — pour que l'UI les
   * annonce sans que le reste du code ait à s'en occuper.
   */
  check(levelCount) {
    const stats = this.save.stats(levelCount);
    const news = [];

    for (const a of ACHIEVEMENTS) {
      if (this.save.data.achievements.includes(a.id)) continue;
      if (a.test(stats)) {
        this.save.data.achievements.push(a.id);
        news.push({ kind: 'achievement', id: a.id, name: a.name, desc: a.desc });
      }
    }
    for (const skin of SKINS) {
      if (this.save.data.unlocked.includes(skin.id)) continue;
      if (isSkinUnlocked(skin, stats)) {
        this.save.data.unlocked.push(skin.id);
        news.push({ kind: 'skin', id: skin.id, name: skin.name, desc: skin.unlock?.label || '' });
      }
    }
    if (news.length) this.save.save();
    return news;
  }

  earned() {
    return ACHIEVEMENTS.map((a) => ({
      ...a, done: this.save.data.achievements.includes(a.id),
    }));
  }
}
