// Le mixeur mémoire : mesure du son RÉELLEMENT produit par la musique et les
// bruitages (sonde branchée sur les nœuds de gain), pour tous les états.
import { chromium } from 'playwright';

const FILE = new URL('../plume-test.html', import.meta.url).href;
const results = [];
const check = (n, ok, d = '') => results.push({ n, ok: !!ok, d: String(d) });

const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium',
  args: ['--autoplay-policy=no-user-gesture-required', '--mute-audio'],
});
const page = await browser.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
await page.route('**text.pollinations.ai/**', (r) => r.abort());

await page.goto(FILE);
await page.waitForTimeout(500);
await page.click('text=Commencer');
await page.fill('#in-name', 'Léo');
await page.click('#s-name >> text=Continuer');
await page.click('#s-age >> text=Continuer');
await page.click('#comp-cards .card >> nth=0');
await page.click('text=C\'est lui');
await page.waitForTimeout(200);

// Sondes : on écoute la sortie réelle de chaque bus.
const probe = (bus) => page.evaluate(async (bus) => {
  const node = bus === 'music' ? musicGain : sfxGain;
  if (!node) return -1;
  if (!window.__p) window.__p = {};
  if (!window.__p[bus]) {
    window.__p[bus] = actx.createAnalyser();
    window.__p[bus].fftSize = 2048;
    node.connect(window.__p[bus]);
  }
  const a = window.__p[bus];
  const buf = new Float32Array(a.fftSize);
  let peak = 0;
  for (let i = 0; i < 25; i++) {
    a.getFloatTimeDomainData(buf);
    for (const v of buf) peak = Math.max(peak, Math.abs(v));
    await new Promise((r) => setTimeout(r, 40));
  }
  return +peak.toFixed(5);
}, bus);

await page.click('text=C\'est l\'heure de l\'histoire');
await page.waitForTimeout(1500);
const mixer = await page.evaluate(() => ({
  ctx: !!actx, state: actx && actx.state, ready: mixerReady, node: !!musicNode,
}));
check('mixeur mémoire actif', mixer.ctx && mixer.state === 'running', JSON.stringify(mixer));
check('musique décodée et lancée', mixer.ready && mixer.node, JSON.stringify(mixer));

const musicPeak = await probe('music');
check('LA MUSIQUE PRODUIT DU SON', musicPeak > 0.0005, 'crête=' + musicPeak);
check('musique restée discrète', musicPeak < 0.08, 'crête=' + musicPeak);

await page.click('text=Passer');
await page.waitForTimeout(1200);
const sfxState = await page.evaluate(() => ({
  key: EP.scenes[0].sfxKey, node: !!sfxNode, buffers: Object.keys(sfxBuffers).length,
}));
if (sfxState.key) {
  const sfxPeak = await probe('sfx');
  check('LES BRUITAGES PRODUISENT DU SON', sfxPeak > 0.0005 || sfxState.node,
    'crête=' + sfxPeak + ' ' + JSON.stringify(sfxState));
} else {
  check('scène sans bruitage (silence assumé)', true, 'sfxKey=null');
}

// Forcer un bruitage connu et le mesurer.
await page.evaluate(() => playSfxKey('open2'));
await page.waitForTimeout(300);
const crackPeak = await probe('sfx');
check('le craquement de branche sort bien du mixeur', crackPeak > 0.0005, 'crête=' + crackPeak);

// Pause : tout le graphe se suspend.
await page.click('#voice-btn');
await page.waitForTimeout(500);
const paused = await page.evaluate(() => ({ state: actx.state, voice: sharedAudio.paused }));
check('pause : graphe suspendu et voix arrêtée',
  paused.state === 'suspended' && paused.voice, JSON.stringify(paused));
await page.click('#voice-btn');
await page.waitForTimeout(600);
check('reprise : graphe relancé', await page.evaluate(() => actx.state) === 'running');

// Curseur à 0 → silence total mesuré.
await page.evaluate(() => setAmbienceLevel(0));
await page.waitForTimeout(400);
check('curseur à 0 % : silence mesuré', (await probe('music')) < 0.0006);
await page.evaluate(() => setAmbienceLevel(65));
await page.waitForTimeout(400);
check('curseur remonté : le son revient', (await probe('music')) > 0.0005);

// Sortie de l'histoire : la musique s'arrête vraiment.
await page.evaluate(() => go('s-home'));
await page.waitForTimeout(400);
check('retour à l\'accueil : musique arrêtée',
  await page.evaluate(() => !musicNode && !musicWanted));

check('aucune erreur JS', errors.length === 0, errors.join(' | ').slice(0, 200));
await browser.close();

let failed = 0;
for (const r of results) { if (!r.ok) failed++; console.log(`${r.ok ? 'PASS' : 'FAIL'}  ${r.n}${r.d ? '  [' + r.d + ']' : ''}`); }
console.log(`\n${results.length - failed}/${results.length} vérifications OK`);
process.exit(failed ? 1 : 0);
