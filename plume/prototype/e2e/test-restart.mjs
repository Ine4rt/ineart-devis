// Test de la remise à zéro : reprendre au chapitre 1 en gardant le profil,
// et tout effacer. Vérifie aussi la confirmation en deux touchers.
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
await page.route('**text.pollinations.ai/**', (r) => r.abort()); // hors ligne

await page.goto(FILE);
await page.waitForTimeout(500);
await page.click('text=Commencer');
await page.fill('#in-name', 'Léo');
await page.click('#s-name >> text=Continuer');
await page.click('#s-age >> text=Continuer');
await page.click('#comp-cards .card >> nth=1'); // le chien
await page.click('text=C\'est lui');

// On simule une aventure déjà avancée (comme l'état réel du client).
await page.evaluate(() => {
  S.episode = 4;
  S.lastSummary = 'La cabane de Grattouille est terminée.';
  S.told = [{ n: 1, title: 'Le jardin du soir', date: '1 juillet' },
            { n: 2, title: 'Le retour du hérisson', date: '2 juillet' },
            { n: 3, title: 'Le passage secret', date: '3 juillet' }];
  persist(); renderHome();
});
await page.waitForTimeout(200);
check('aventure avancée au chapitre 4',
  (await page.textContent('#ep-label')).includes('Chapitre 4'));
check('rappel du chapitre précédent affiché', await page.locator('#seednote').isVisible());

// ─── Espace parent → reprendre au chapitre 1 ───────────────────────────────
await page.evaluate(() => openParent());
await page.waitForTimeout(300);
await page.click('#btn-restart');
const armed = await page.textContent('#btn-restart');
check('premier appui : demande confirmation', armed.includes('Appuyez à nouveau'), armed);
check('rien effacé avant confirmation', await page.evaluate(() => S.episode) === 4);

await page.click('#btn-restart');
await page.waitForTimeout(400);
const after = await page.evaluate(() => ({
  episode: S.episode, told: S.told.length, summary: S.lastSummary,
  name: S.name, comp: S.comp && S.comp.id, saved: JSON.parse(localStorage.getItem('plume-state')).episode,
}));
check('retour au chapitre 1', after.episode === 1, JSON.stringify(after));
check('bibliothèque vidée', after.told === 0);
check('plus aucun rappel du passé', after.summary === '');
check('profil conservé (prénom, compagnon)', after.name === 'Léo' && after.comp === 'chien');
check('remise à zéro enregistrée', after.saved === 1);
check('accueil affiché au chapitre 1',
  (await page.textContent('#ep-label')).includes('Chapitre 1'));
check('bandeau de rappel masqué', await page.locator('#seednote').isHidden());
check('bouton de réécoute masqué', await page.locator('#replay-btn').isHidden());

// ─── Tout effacer ──────────────────────────────────────────────────────────
await page.evaluate(() => openParent());
await page.waitForTimeout(200);
await page.click('#btn-wipe');
await page.click('#btn-wipe');
await page.waitForTimeout(900);
check('tout effacer : retour au premier soir',
  await page.locator('#s-splash.active').count() === 1);
check('stockage vidé', await page.evaluate(() => !localStorage.getItem('plume-state')));
check('aucune erreur JS', errors.length === 0, errors.join(' | ').slice(0, 200));

await browser.close();
let failed = 0;
for (const r of results) {
  if (!r.ok) failed++;
  console.log(`${r.ok ? 'PASS' : 'FAIL'}  ${r.name}${r.detail ? '  [' + r.detail + ']' : ''}`);
}
console.log(`\n${results.length - failed}/${results.length} vérifications OK`);
process.exit(failed ? 1 : 0);
