/**
 * Sauvegarde locale.
 *
 * Format versionné et migrable dès le départ : ajouter un chapitre ou une
 * statistique ne doit jamais effacer la progression d'un joueur. Toute la
 * lecture passe par des valeurs par défaut, ce qui rend une sauvegarde
 * partiellement corrompue inoffensive.
 *
 * L'API est volontairement synchrone et plate — `load()` / `save()` — pour
 * qu'un backend distant (iCloud, Google Play Jeux) puisse être branché plus
 * tard en remplaçant uniquement `read()` et `write()`.
 */

const KEY = 'voidrunner.save.v1';

const DEFAULTS = {
  version: 1,
  levels: {}, // num -> { done, deaths, best, secret }
  totalDeaths: 0,
  deathsByDay: {}, // 'AAAA-MM-JJ' -> nombre
  skin: 'standard',
  unlocked: ['standard'],
  achievements: [],
  cleanStreak: 0,
  bestCleanStreak: 0,
  settings: { music: 1, sfx: 1, haptics: true, hand: 'right', quality: 1 },
  seenIntro: false,
  playTime: 0,
};

function today() {
  return new Date().toISOString().slice(0, 10);
}

export class SaveManager {
  constructor() {
    this.data = this.load();
    this.dirty = false;
    this.flushTimer = null;
  }

  read() {
    try { return localStorage.getItem(KEY); } catch { return null; }
  }

  write(str) {
    try { localStorage.setItem(KEY, str); } catch { /* mode privé : on continue sans sauvegarde */ }
  }

  load() {
    const raw = this.read();
    if (!raw) return structuredClone(DEFAULTS);
    try {
      const parsed = JSON.parse(raw);
      // Fusion défensive : un champ manquant reprend sa valeur par défaut.
      return {
        ...structuredClone(DEFAULTS),
        ...parsed,
        settings: { ...DEFAULTS.settings, ...(parsed.settings || {}) },
        levels: parsed.levels || {},
      };
    } catch {
      return structuredClone(DEFAULTS);
    }
  }

  /** Écriture différée : on n'écrit jamais localStorage en plein niveau. */
  save() {
    this.dirty = true;
    if (this.flushTimer) return;
    this.flushTimer = setTimeout(() => {
      this.flushTimer = null;
      if (!this.dirty) return;
      this.dirty = false;
      this.write(JSON.stringify(this.data));
    }, 400);
  }

  flush() {
    if (this.flushTimer) { clearTimeout(this.flushTimer); this.flushTimer = null; }
    this.write(JSON.stringify(this.data));
    this.dirty = false;
  }

  // --- Accès niveau ----------------------------------------------------------

  level(num) {
    const k = String(num);
    if (!this.data.levels[k]) this.data.levels[k] = { done: false, deaths: 0, best: null, secret: false };
    return this.data.levels[k];
  }

  addDeath(num) {
    const l = this.level(num);
    l.deaths++;
    this.data.totalDeaths++;
    const d = today();
    this.data.deathsByDay[d] = (this.data.deathsByDay[d] || 0) + 1;
    this.data.cleanStreak = 0;
    this.save();
  }

  complete(num, time, deathsThisRun) {
    const l = this.level(num);
    const first = !l.done;
    l.done = true;
    if (l.best == null || time < l.best) l.best = time;
    if (deathsThisRun === 0) {
      this.data.cleanStreak++;
      this.data.bestCleanStreak = Math.max(this.data.bestCleanStreak, this.data.cleanStreak);
    } else {
      this.data.cleanStreak = 0;
    }
    this.save();
    return first;
  }

  markSecret(num) {
    const l = this.level(num);
    if (l.secret) return false;
    l.secret = true;
    this.save();
    return true;
  }

  unlock(skinId) {
    if (this.data.unlocked.includes(skinId)) return false;
    this.data.unlocked.push(skinId);
    this.save();
    return true;
  }

  // --- Statistiques agrégées ---------------------------------------------------

  stats(levelCount) {
    const lv = Object.entries(this.data.levels);
    const completed = lv.filter(([, v]) => v.done);
    const times = completed.map(([, v]) => v.best).filter((t) => t != null);
    const chapterOfNum = (n) => Math.ceil(n / 6);
    return {
      totalDeaths: this.data.totalDeaths,
      deathsToday: this.data.deathsByDay[today()] || 0,
      completedCount: completed.length,
      levelCount,
      secrets: lv.filter(([, v]) => v.secret).length,
      bestTimeAny: times.length ? Math.min(...times) : 0,
      totalTime: times.reduce((a, b) => a + b, 0),
      bestCleanStreak: this.data.bestCleanStreak,
      maxChapter: completed.length
        ? Math.max(...completed.map(([k]) => chapterOfNum(Number(k))))
        : 1,
      worstLevel: lv.slice().sort((a, b) => b[1].deaths - a[1].deaths)[0] || null,
      playTime: this.data.playTime,
    };
  }

  /** Progression : le plus haut niveau jouable (les suivants sont verrouillés). */
  maxUnlockedLevel(levelCount) {
    let n = 1;
    while (n < levelCount && this.data.levels[String(n)]?.done) n++;
    return n;
  }

  reset() {
    this.data = structuredClone(DEFAULTS);
    this.flush();
  }
}
