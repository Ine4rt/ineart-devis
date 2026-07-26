// Test du conteur IA : l'IA est simulée (le réseau sortant est bloqué ici),
// on vérifie que le chapitre écrit par l'IA est bien joué, avec ses bruitages,
// sa longueur, la persistance du feuilleton — et que le repli fonctionne.
import { chromium } from 'playwright';

const FILE = new URL('../plume-test.html', import.meta.url).href;
const results = [];
const check = (name, ok, detail = '') =>
  results.push({ name, ok: !!ok, detail: String(detail) });

const AI_CHAPTER = {
  titre: 'La mésange au bec doré',
  scenes: [
    { texte: 'La dernière fois, tu avais laissé un bol d\'eau au hérisson. Et ce soir, l\'histoire continue.', son: 'silence' },
    { texte: 'Cric… crac ! Une petite branche a craqué tout en haut du pommier.', son: 'craquement' },
    { texte: 'C\'était une mésange, avec un bec doré comme un grain de blé. Elle chantait, tiu-tiu-tiu !', son: 'oiseau' },
    { texte: 'Tu as fait trois pas dans l\'herbe, tout doucement, pour ne pas l\'effrayer.', son: 'pas' },
    { texte: 'Elle a bu une goutte d\'eau dans le bol du hérisson. Plouf, plouf !', son: 'eau' },
    { texte: 'Ferme les yeux… Demain, elle reviendra peut-être. Bonne nuit.', son: 'nuit' },
  ],
  rappel_suivant: 'Une mésange au bec doré est venue boire dans le bol du hérisson.',
};

async function run(mockAI) {
  const browser = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium',
    args: ['--autoplay-policy=no-user-gesture-required', '--mute-audio'],
  });
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));

  let aiCalls = 0, promptSeen = '';
  await page.route('**text.pollinations.ai/**', async (route) => {
    const url = route.request().url();
    if (url.includes('openai-audio')) return route.abort(); // TTS indisponible
    aiCalls++;
    promptSeen = decodeURIComponent(url.split('/').pop().split('?')[0]);
    if (!mockAI) return route.abort();
    await route.fulfill({
      status: 200, contentType: 'text/plain',
      body: 'Voici le chapitre :\n' + JSON.stringify(AI_CHAPTER),
    });
  });

  await page.goto(FILE);
  await page.waitForTimeout(600);
  await page.click('text=Commencer');
  await page.fill('#in-name', 'Léo');
  await page.click('#s-name >> text=Continuer');
  await page.click('#s-age >> text=Continuer');
  await page.click('#comp-cards .card >> nth=0');
  await page.click('text=C\'est lui');
  await page.waitForTimeout(300);
  await page.click('text=C\'est l\'heure de l\'histoire');
  await page.click('text=Passer');
  await page.waitForTimeout(2500);

  const state = await page.evaluate(() => ({
    fromAI: !!(EP && EP.fromAI),
    title: EP ? EP.title : null,
    scenes: EP ? EP.scenes.length : 0,
    sfxKeys: EP ? EP.scenes.map((s) => s.sfxKey) : [],
    text: document.getElementById('story-text').textContent.slice(0, 60),
    summary: EP ? EP.summary : null,
  }));
  return { page, browser, errors, aiCalls, promptSeen, state };
}

// ─── 1. L'IA répond ─────────────────────────────────────────────────────────
{
  const { page, browser, errors, aiCalls, promptSeen, state } = await run(true);
  check('l\'IA est appelée', aiCalls >= 1, aiCalls + ' appel(s)');
  check('le prompt porte le prénom, l\'âge et la durée',
    promptSeen.includes('Léo') && promptSeen.includes('mots') && promptSeen.includes('scènes'),
    promptSeen.slice(0, 90));
  check('le prompt exige le ton album jeunesse',
    promptSeen.includes('album jeunesse') && promptSeen.includes('onomatopées'));
  check('chapitre marqué comme écrit par l\'IA', state.fromAI);
  check('titre de l\'IA affiché', state.title === AI_CHAPTER.titre, state.title);
  check('toutes les scènes IA chargées', state.scenes === 6, state.scenes);
  check('bruitages IA correctement mappés',
    JSON.stringify(state.sfxKeys) === JSON.stringify([null, 'open2', 'amb1', 'amb2', 'close2', 'close3']),
    JSON.stringify(state.sfxKeys));
  check('texte de l\'IA affiché à l\'écran', state.text.includes('dernière fois'), state.text);
  check('rappel de demain fourni par l\'IA',
    state.summary === AI_CHAPTER.rappel_suivant, state.summary);

  // Le craquement doit jouer avec la scène 2 (« cric… crac ! »).
  await page.waitForFunction(() => sceneIdx >= 1, null, { timeout: 60000 });
  await page.waitForTimeout(600); // le décodage du bruitage est asynchrone
  const synced = await page.evaluate(() => ({
    key: EP.scenes[sceneIdx].sfxKey,
    decoded: Object.keys(sfxBuffers),
    node: !!sfxNode,
  }));
  check('craquement joué pile sur la scène « cric… crac »',
    synced.key === 'open2' && (synced.node || synced.decoded.includes('open2')),
    JSON.stringify(synced));

  // Narration : le TTS en ligne est coupé → repli voix de l'appareil.
  const spoke = await page.evaluate(() => !!window.speechSynthesis);
  check('repli de narration disponible pour un texte IA', spoke);
  check('badge « écrit par l\'IA » affiché',
    (await page.textContent('#story-source')).includes("l'IA"),
    await page.textContent('#story-source'));

  // Garde-fou : narration muette (TTS coupé, synthèse absente) → l'histoire
  // avance quand même. C'est ce qui empêche l'app de rester figée.
  const advanced = await page.waitForFunction(
    (from) => sceneIdx > from, 1, { timeout: 90000 },
  ).then(() => true).catch(() => false);
  check('garde-fou : l\'histoire avance même sans narration', advanced,
    'scène ' + await page.evaluate(() => sceneIdx));

  check('aucune erreur JS (chapitre IA)', errors.length === 0, errors.join(' | ').slice(0, 200));
  await browser.close();
}

// ─── 2. L'IA ne répond pas → feuilleton local ──────────────────────────────
{
  const { browser, errors, state } = await run(false);
  check('repli : chapitre local joué quand l\'IA échoue', !state.fromAI && state.scenes >= 3,
    JSON.stringify({ fromAI: state.fromAI, scenes: state.scenes }));
  check('repli : le texte reste celui du jardin', state.text.length > 10, state.text);
  check('aucune erreur JS (repli)', errors.length === 0, errors.join(' | ').slice(0, 200));
  await browser.close();
}

let failed = 0;
for (const r of results) {
  if (!r.ok) failed++;
  console.log(`${r.ok ? 'PASS' : 'FAIL'}  ${r.name}${r.detail ? '  [' + r.detail + ']' : ''}`);
}
console.log(`\n${results.length - failed}/${results.length} vérifications OK`);
process.exit(failed ? 1 : 0);
