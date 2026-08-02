/**
 * Construit une version « un seul fichier » du site.
 *
 * Le site normal est un dossier : index.html plus des feuilles, des polices et
 * des photos. Certaines destinations n'acceptent qu'un fichier unique et
 * interdisent toute requête vers l'extérieur — un aperçu partageable, une
 * pièce jointe, une clé USB. Ce script y répond en repliant tout dans le HTML :
 * la CSS et le script en ligne, les polices et les images en data: URI.
 *
 * Les photos sont recompressées au passage : en base64 une image pèse un tiers
 * de plus qu'en binaire, et un fichier unique n'a pas de chargement différé —
 * tout arrive d'un bloc. Sans cette étape, l'aperçu dépasserait deux mégaoctets.
 *
 *   node build-embed.mjs            → chauffage-demarche.embed.html
 *
 * Le dossier reste la version de référence ; ce fichier n'en est qu'un tirage.
 */

import { readFileSync, writeFileSync, statSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const p = (...s) => join(here, ...s);

const dataUri = (file, mime) =>
  `data:${mime};base64,${readFileSync(file).toString("base64")}`;

/* Recompression via Pillow : plus petit que l'original de dossier, parce que
   l'aperçu n'a pas besoin d'une photo imprimable. */
function photo(name, largeur, qualite) {
  const src = p("assets/img", name);
  const out = `/tmp/embed-${name}`;
  execFileSync("python3", ["-c", `
from PIL import Image
im = Image.open("${src}")
if im.width > ${largeur}:
    im = im.resize((${largeur}, round(im.height * ${largeur} / im.width)), Image.LANCZOS)
im.convert("RGB").save("${out}", "JPEG", quality=${qualite}, optimize=True, progressive=True)
`]);
  return dataUri(out, "image/jpeg");
}

let html = readFileSync(p("index.html"), "utf8");

/* --- Polices et styles --------------------------------------------------- */
let css = readFileSync(p("assets/css/style.css"), "utf8")
  .replace('url("../fonts/demarche-display.woff2")', `url("${dataUri(p("assets/fonts/demarche-display.woff2"), "font/woff2")}")`)
  .replace('url("../fonts/demarche-text.woff2")', `url("${dataUri(p("assets/fonts/demarche-text.woff2"), "font/woff2")}")`);

html = html.replace(
  '<link rel="stylesheet" href="assets/css/style.css" />',
  `<style>\n${css}\n</style>`
);

/* --- Script -------------------------------------------------------------- */
html = html.replace(
  '<script src="assets/js/main.js" defer></script>',
  `<script>\n${readFileSync(p("assets/js/main.js"), "utf8")}\n</script>`
);

/* --- Images -------------------------------------------------------------- */
const images = {
  "assets/img/sol-1.jpg": photo("sol-1.jpg", 1100, 68),
  "assets/img/collecteur.jpg": photo("collecteur.jpg", 700, 70),
  "assets/img/tuyaux.jpg": photo("tuyaux.jpg", 620, 72),
  "assets/img/sol-2.jpg": photo("sol-2.jpg", 620, 72),
  "assets/img/sol-3.jpg": photo("sol-3.jpg", 620, 72),
};

for (const [chemin, uri] of Object.entries(images)) {
  html = html.replaceAll(chemin, uri);
}

/* Le chargement différé n'a plus de sens : les images sont déjà dans le
   document, l'attribut ne ferait que retarder leur affichage. */
html = html.replaceAll(' loading="lazy"', "");

const sortie = p("chauffage-demarche.embed.html");
writeFileSync(sortie, html);

/* --- Sortie 2 : fragment ---------------------------------------------------
   Certains hôtes d'aperçu fournissent eux-mêmes <html>, <head> et <body> et
   n'acceptent que le contenu. On leur donne le titre, les données de
   référencement, les styles et le corps — sans l'enveloppe.
   -------------------------------------------------------------------------- */
const entre = (balise, source) =>
  source.match(new RegExp(`<${balise}[^>]*>([\\s\\S]*)</${balise}>`))?.[1] ?? "";

const fragment = [
  html.match(/<title>[\s\S]*?<\/title>/)?.[0] ?? "",
  html.match(/<script type="application\/ld\+json">[\s\S]*?<\/script>/)?.[0] ?? "",
  html.match(/<style>[\s\S]*?<\/style>/)?.[0] ?? "",
  entre("body", html).trim(),
].join("\n");

writeFileSync(p("chauffage-demarche.artifact.html"), fragment);

/* Garde-fou : une seule ressource oubliée et la page se dégrade en silence
   là où elle est ouverte. On ne regarde que les attributs — les commentaires
   du fichier citent les chemins d'origine, c'est voulu. */
const oublis = [...html.matchAll(/(?:src|href)="(assets\/[^"]+)"/g)].map((m) => m[1]);
if (oublis.length) throw new Error(`ressource(s) non incorporée(s) : ${oublis.join(", ")}`);

console.log(`chauffage-demarche.embed.html    — ${Math.round(statSync(sortie).size / 1024)} Ko`);
console.log(`chauffage-demarche.artifact.html — ${Math.round(statSync(p("chauffage-demarche.artifact.html")).size / 1024)} Ko`);
