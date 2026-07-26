// Le moteur d'histoires hors ligne : longueur pilotée par le curseur, chapitre
// différent chaque soir, feuilleton qui s'enchaîne, bruitages cohérents,
// journée de l'enfant tissée, réécoute reproductible.
import { chromium } from 'playwright';

const FILE = new URL('../plume-test.html', import.meta.url).href;
const results = [];
const check = (name, ok, detail = '') => results.push({ name, ok: !!ok, detail: String(detail) });

const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium',
  args: ['--autoplay-policy=no-user-gesture-required', '--mute-audio'],
});
const page = await browser.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
await page.route('**text.pollinations.ai/**', (r) => r.abort()); // IA injoignable

await page.goto(FILE);
await page.waitForTimeout(500);
await page.click('text=Commencer');
await page.fill('#in-name', 'Léo');
await page.click('#s-name >> text=Continuer');
await page.click('#s-age >> text=Continuer');
await page.click('#comp-cards .card >> nth=2');
await page.click('text=C\'est lui');
await page.waitForTimeout(200);

const compose = (n, minutes, checkin = { victory: '', fear: '' }) => page.evaluate(
  ({ n, minutes, checkin }) => {
    S.minutes = minutes; S.checkin = checkin;
    const ch = composeChapter(n);
    return {
      title: ch.title,
      scenes: ch.scenes.length,
      words: ch.scenes.reduce((s, x) => s + x.text.split(/\s+/).length, 0),
      sfx: ch.scenes.map((s) => s.sfxKey),
      texts: ch.scenes.map((s) => s.text),
      summary: ch.summary,
      weaves: ch.scenes.filter((s) => s.weave).length,
      recap: ch.scenes[0].recap === true,
    };
  }, { n, minutes, checkin },
);

// ─── Longueur pilotée par le curseur ───────────────────────────────────────
const short = await compose(1, 2);
const long = await compose(1, 9);
check('2 min : chapitre court', short.words >= 120 && short.words <= 330, short.words + ' mots');
check('9 min : chapitre nettement plus long', long.words > short.words * 2,
  `${short.words} → ${long.words} mots`);
check('9 min : beaucoup de scènes', long.scenes >= 8, long.scenes + ' scènes');
check('9 min demandées : le moteur hors ligne compose au moins 5 min',
  long.words / 115 >= 5, (long.words / 115).toFixed(1) + ' min composées');

// ─── Un chapitre différent chaque soir ─────────────────────────────────────
const titles = new Set();
for (let n = 1; n <= 8; n++) titles.add((await compose(n, 5)).title);
check('chapitres variés d\'un soir à l\'autre', titles.size >= 5,
  titles.size + ' histoires différentes sur 8 soirs : ' + [...titles].join(', '));

const again = await compose(3, 5);
const same = await compose(3, 5);
check('réécoute reproductible (même chapitre = même texte)',
  again.title === same.title && again.texts[0] === same.texts[0]);

// ─── Bruitages cohérents ───────────────────────────────────────────────────
const known = await page.evaluate(() => Object.keys(SFX));
check('tous les bruitages existent réellement',
  long.sfx.every((k) => k === null || known.includes(k)), JSON.stringify(long.sfx));
check('la dernière scène porte le son de la nuit',
  long.sfx[long.sfx.length - 1] === 'close3', long.sfx[long.sfx.length - 1]);
check('la fin conduit au sommeil',
  /Bonne nuit/.test(long.texts[long.texts.length - 1]));

// ─── Feuilleton et journée de l'enfant ─────────────────────────────────────
await page.evaluate(() => { S.lastSummary = 'une nuit de lucioles a illuminé le jardin'; });
const withRecap = await compose(4, 5);
check('rappel du chapitre précédent en ouverture',
  withRecap.recap && /La dernière fois/.test(withRecap.texts[0]), withRecap.texts[0].slice(0, 50));
check('résumé fourni pour le rappel de demain', !!withRecap.summary, withRecap.summary);

const woven = await compose(2, 5, { victory: 'du vélo sans les roulettes', fear: 'le dentiste' });
check('réussite et inquiétude du jour tissées', woven.weaves === 2, woven.weaves + ' scènes tissées');
check('la réussite est dite sans nommer la situation réelle',
  woven.texts.some((t) => /Bravo/.test(t)) && !woven.texts.some((t) => /vélo|roulettes/.test(t)));
check('l\'inquiétude est portée par un animal',
  woven.texts.some((t) => /hérisson/.test(t) && /peur/.test(t))
  && !woven.texts.some((t) => /dentiste/.test(t)));

// ─── Lecture réelle de bout en bout ────────────────────────────────────────
await page.evaluate(() => { S.minutes = 3; S.checkin = { victory: '', fear: '' }; persist(); });
await page.click('text=C\'est l\'heure de l\'histoire');
await page.click('text=Passer');
await page.waitForTimeout(2000);
const playing = await page.evaluate(() => ({
  composed: !!(EP && EP.composed),
  source: document.getElementById('story-source').textContent,
  text: document.getElementById('story-text').textContent.slice(0, 40),
}));
check('chapitre composé et joué', playing.composed, JSON.stringify(playing));
check('provenance affichée à l\'écran', playing.source.includes('hors ligne'), playing.source);
// Honnêteté : la durée annoncée est celle réellement composée, pas la demandée.
const honest = await page.evaluate(() => {
  const words = EP.scenes.reduce((n, s) => n + s.text.split(/\s+/).length, 0);
  const shown = +document.getElementById('story-source').textContent.match(/~(\d+) min/)[1];
  return { shown, real: Math.max(1, Math.round(words / 115)) };
});
check('la durée affichée est la durée réelle', honest.shown === honest.real,
  JSON.stringify(honest));
check('aucune erreur JS', errors.length === 0, errors.join(' | ').slice(0, 200));

await browser.close();
let failed = 0;
for (const r of results) {
  if (!r.ok) failed++;
  console.log(`${r.ok ? 'PASS' : 'FAIL'}  ${r.name}${r.detail ? '  [' + r.detail + ']' : ''}`);
}
console.log(`\n${results.length - failed}/${results.length} vérifications OK`);
process.exit(failed ? 1 : 0);
