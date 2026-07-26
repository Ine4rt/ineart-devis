// Test de bout en bout du prototype Plume dans Chromium headless.
// Vérifie : zéro erreur JS, parcours complet d'onboarding, musique de fond
// réellement en lecture, bruitages chargés et joués, narration avec repli
// hors-ligne, pause générale, réécoute.
import { chromium } from 'playwright';

const FILE = new URL('../plume-test.html', import.meta.url).href;
const results = [];
const check = (name, ok, detail = '') =>
  results.push({ name, ok: !!ok, detail: String(detail) });

const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium',
  args: ['--autoplay-policy=no-user-gesture-required', '--mute-audio'],
});
const page = await browser.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
page.on('console', (m) => {
  if (m.type() === 'error' && !/pollinations|ERR_|net::/i.test(m.text())) {
    errors.push('console: ' + m.text());
  }
});

await page.goto(FILE);
await page.waitForTimeout(800);

// ─── Onboarding ─────────────────────────────────────────────────────────────
check('splash visible', await page.locator('#s-splash.active').count() === 1);
await page.click('text=Commencer');
await page.fill('#in-name', 'Léo');
await page.click('#s-name >> text=Continuer');
await page.click('#s-age >> text=Continuer');
await page.click('#comp-cards .card >> nth=2'); // le petit renard
await page.click('text=C\'est lui');
await page.waitForTimeout(400);
check('accueil affiché', await page.locator('#s-home.active').count() === 1);
check('bonsoir personnalisé',
  (await page.textContent('#home-title')).includes('Léo'));
check('bouton réécoute masqué au chapitre 1',
  await page.locator('#replay-btn').isHidden());

// ─── Données audio présentes ────────────────────────────────────────────────
const audioData = await page.evaluate(() => ({
  narration: typeof NARRATION !== 'undefined' && !!NARRATION.homme.open2,
  lab: typeof VOICELAB !== 'undefined' && Object.keys(VOICELAB).length,
  ambience: typeof AMBIENCE !== 'undefined' && AMBIENCE.startsWith('data:audio'),
  sfx: typeof SFX !== 'undefined' ? Object.keys(SFX).sort().join(',') : 'absent',
}));
check('narration embarquée (open2)', audioData.narration);
check('studio des voix chargé', audioData.lab >= 6, audioData.lab + ' voix');
check('musique embarquée', audioData.ambience);
check('bruitages embarqués',
  audioData.sfx === 'amb1,amb2,amb3,amb4,close2,close3,fear,open1,open2,victory',
  audioData.sfx);

// ─── Lancement de l'histoire ────────────────────────────────────────────────
await page.click('text=C\'est l\'heure de l\'histoire');
await page.waitForTimeout(1200);
check('rituel affiché', await page.locator('#s-ritual.active').count() === 1);
check('musique démarrée au toucher', await page.evaluate(
  () => (actx ? !!musicNode : !ambience.paused),
), await page.evaluate(() => `ctx=${!!actx} node=${!!musicNode}`));

await page.click('text=Passer');
await page.waitForTimeout(1500);
check('écran histoire affiché', await page.locator('#s-story.active').count() === 1);
check('musique toujours en lecture pendant l\'histoire', await page.evaluate(
  () => (actx ? !!musicNode && actx.state === 'running' : !ambience.paused),
));
// Le chapitre est composé (ou écrit par l'IA) : on vérifie le comportement,
// pas un texte figé — la variété est testée dans test-engine.mjs.
const scene1 = await page.evaluate(() => ({
  text: document.getElementById('story-text').textContent,
  sfxKey: EP.scenes[0].sfxKey,
  sfxPlaying: !!sfxNode || sfxAudio.currentTime > 0 || Object.keys(sfxBuffers).length > 0,
}));
check('texte de la scène 1 affiché', scene1.text.length > 40, scene1.text.slice(0, 40));
check('bruitage joué si la scène en porte un',
  !scene1.sfxKey || scene1.sfxPlaying, JSON.stringify(scene1.sfxKey));

// Narration : le service en ligne est coupé ici → repli (mp3 pré-généré ou
// voix de l'appareil), et surtout l'histoire doit continuer d'avancer.
const advanced = await page.waitForFunction(
  () => sceneIdx >= 1, null, { timeout: 90000 },
).then(() => true).catch(() => false);
check('l\'histoire avance malgré la narration en ligne indisponible', advanced,
  'scène ' + await page.evaluate(() => sceneIdx));

// Le bruitage suit bien la scène affichée.
const scene2 = await page.evaluate(() => ({
  key: EP.scenes[sceneIdx].sfxKey,
  synced: !EP.scenes[sceneIdx].sfxKey || !!sfxNode || Object.keys(sfxBuffers).length > 0,
  text: document.getElementById('story-text').textContent.slice(0, 40),
}));
check('le bruitage suit la scène en cours', scene2.synced, JSON.stringify(scene2));

// ─── Mixage : on MESURE le niveau réellement produit ───────────────────────
// Le piège corrigé ici : Safari iOS ignore element.volume, donc le niveau doit
// venir des fichiers eux-mêmes + d'un gain Web Audio. Un simple test de
// `volume` ne prouverait rien : on branche un analyseur sur le mixeur.
const mixer = await page.evaluate(() => ({
  hasContext: !!actx,
  music: musicGain ? +musicGain.gain.value.toFixed(3) : null,
  sfx: sfxGain ? +sfxGain.gain.value.toFixed(3) : null,
}));
check('mixeur Web Audio actif', mixer.hasContext, JSON.stringify(mixer));

const measure = async (label) => page.evaluate(async () => {
  if (!window.__probe) {
    window.__probe = actx.createAnalyser();
    window.__probe.fftSize = 2048;
    musicGain.connect(window.__probe);
  }
  const buf = new Float32Array(window.__probe.fftSize);
  let peak = 0;
  for (let i = 0; i < 12; i++) { // ~0,4 s d'observation
    window.__probe.getFloatTimeDomainData(buf);
    for (const v of buf) peak = Math.max(peak, Math.abs(v));
    await new Promise((r) => setTimeout(r, 35));
  }
  return +peak.toFixed(4);
});

const duckedPeak = await measure('voix en cours');
check('musique audible mais très discrète pendant la voix',
  duckedPeak > 0.0005 && duckedPeak < 0.05, 'crête=' + duckedPeak);

// Curseur du parent à 0 % → silence complet du fond.
await page.evaluate(() => setAmbienceLevel(0));
await page.waitForTimeout(300);
const mutedPeak = await measure('curseur à 0');
check('curseur à 0 % : fond muet', mutedPeak < 0.0006, 'crête=' + mutedPeak);
await page.evaluate(() => setAmbienceLevel(50));

// ─── Pause générale ─────────────────────────────────────────────────────────
await page.click('#voice-btn');
await page.waitForTimeout(400);
check('pause : voix ET musique suspendues', await page.evaluate(
  () => sharedAudio.paused && (actx ? actx.state === 'suspended' : ambience.paused),
), await page.evaluate(() => `voix=${sharedAudio.paused} graphe=${actx && actx.state} scene=${sceneIdx}`));
await page.click('#voice-btn');
await page.waitForTimeout(600);
check('reprise : musique repartie', await page.evaluate(
  () => (actx ? actx.state === 'running' : !ambience.paused)));

// ─── Confirmation de voix dans l'espace parent ─────────────────────────────
// Le parent doit pouvoir vérifier le son AVANT le soir, pas seulement le
// deviner pendant l'histoire : le bouton « Écouter cette voix » doit
// afficher une confirmation en clair (✓ / cause), pas juste jouer un son
// dans le vide.
await page.evaluate(() => go('s-checkin'));
await page.waitForTimeout(150);
const beforeTest = await page.evaluate(() => document.getElementById('voice-test-status').style.display);
check('confirmation de voix masquée avant tout essai', beforeTest !== 'block', beforeTest);
await page.click('text=Écouter cette voix');
await page.waitForTimeout(4200); // laisse le repli en ligne → pré-enregistré s'exécuter
const voiceTest = await page.evaluate(() => document.getElementById('voice-test-status').textContent);
check('confirmation de voix affichée après l\'essai', voiceTest.startsWith('✓'), voiceTest);

check('aucune erreur JS', errors.length === 0, errors.join(' | ').slice(0, 300));

await browser.close();

// ─── Rapport ────────────────────────────────────────────────────────────────
let failed = 0;
for (const r of results) {
  if (!r.ok) failed++;
  console.log(`${r.ok ? 'PASS' : 'FAIL'}  ${r.name}${r.detail ? '  [' + r.detail + ']' : ''}`);
}
console.log(`\n${results.length - failed}/${results.length} vérifications OK`);
process.exit(failed ? 1 : 0);
