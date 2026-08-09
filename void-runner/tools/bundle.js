import { readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve, relative } from 'node:path';

/**
 * Assemble le jeu en UN seul fichier HTML autonome.
 *
 * Pourquoi un assembleur maison plutôt que Rollup ou esbuild : le projet n'a
 * aucune dépendance, aucun import npm, aucun export par défaut et aucun cycle.
 * Résoudre ces imports-là tient en quarante lignes — y ajouter un outil de
 * construction coûterait plus cher que le problème qu'il résout, et ferait
 * perdre au projet sa propriété la plus utile : il s'ouvre et se lit tel quel.
 *
 * Le résultat sert à partager une version jouable d'un clic (page unique, zéro
 * requête réseau). La version de développement, elle, reste en modules ES
 * séparés — c'est celle qu'on modifie.
 *
 *   node tools/bundle.js  →  dist/void-runner.html
 */

const ROOT = resolve(dirname(new URL(import.meta.url).pathname), '..');
const ENTRY = resolve(ROOT, 'src/main.js');

const IMPORT_RE = /^import\s+[\s\S]*?\bfrom\s+'([^']+)';?[ \t]*$/gm;

const seen = new Map();
const order = [];

/** Parcours en profondeur : chaque module est émis après ses dépendances. */
async function walk(file) {
  if (seen.has(file)) return;
  seen.set(file, true);
  const src = await readFile(file, 'utf8');
  const deps = [...src.matchAll(IMPORT_RE)].map((m) => m[1]);
  for (const d of deps) {
    if (!d.startsWith('.')) throw new Error(`Import non relatif dans ${file} : ${d}`);
    await walk(resolve(dirname(file), d));
  }
  order.push({ file, src });
}

await walk(ENTRY);

// Les modules deviennent une seule portée : deux symboles de même nom se
// masqueraient silencieusement. On préfère échouer bruyamment.
const DECL = /^(?:export\s+)?(?:const|let|var|class|function)\s+([A-Za-z_$][\w$]*)/gm;
const owners = new Map();
const chunks = [];

for (const { file, src } of order) {
  const rel = relative(ROOT, file);
  for (const m of src.matchAll(DECL)) {
    const name = m[1];
    if (owners.has(name)) {
      throw new Error(`Symbole « ${name} » déclaré dans ${owners.get(name)} et ${rel}`);
    }
    owners.set(name, rel);
  }
  const body = src
    .replace(IMPORT_RE, '')
    .replace(/^export\s+(const|let|var|class|function|async)\b/gm, '$1');
  if (/^\s*export\b/m.test(body)) throw new Error(`Export non géré dans ${rel}`);
  chunks.push(`/* ── ${rel} ${'─'.repeat(Math.max(0, 62 - rel.length))} */\n${body.trim()}`);
}

let js = chunks.join('\n\n');
// Pas de service worker dans une page unique : il n'y a rien à mettre en cache.
js = js.replace(/if \('serviceWorker' in navigator[\s\S]*?\n}\n/, '');

const css = await readFile(resolve(ROOT, 'styles.css'), 'utf8');
const html = await readFile(resolve(ROOT, 'index.html'), 'utf8');
const body = html
  .slice(html.indexOf('<body>') + 6, html.indexOf('</body>'))
  .replace(/<script[\s\S]*?<\/script>/g, '')
  .trim();

const out = `<title>VOID RUNNER</title>
<style>
${css.trim()}
/* La page hôte peut être n'importe quoi : on impose notre fond et notre cadre. */
html, body { background: #05060f; margin: 0; }
</style>
${body}
<script type="module">
${js}
</script>
`;

await writeFile(resolve(ROOT, 'dist/void-runner.html'), out);
console.log(
  `dist/void-runner.html — ${order.length} modules, ${(out.length / 1024).toFixed(0)} Ko`,
);
