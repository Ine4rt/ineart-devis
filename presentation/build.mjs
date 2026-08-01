import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Construit la page de présentation.
 *
 * Une seule source (`page.body.html`) produit deux sorties :
 *   · index.html          — document complet, déployable tel quel sur n'importe
 *                           quel hébergement statique (aucune dépendance).
 *   · page.embed.html     — le même contenu sans <html>/<head>/<body>, pour les
 *                           contextes qui fournissent déjà l'enveloppe.
 *
 * Les polices sont injectées en base64 à la construction : la page publiée ne
 * fait aucune requête réseau, donc aucun repli silencieux sur une police
 * système si un CDN est bloqué ou lent.
 */

const here = dirname(fileURLToPath(import.meta.url));
const fontsDir = process.env.FONTS_DIR ?? join(here, "fonts");

const source = readFileSync(join(here, "page.body.html"), "utf8");

const withFonts = source.replace(
  "__ARCHIVO__",
  readFileSync(join(fontsDir, "archivo.woff2.b64"), "utf8").trim(),
);

// --- Sortie 1 : fragment embarquable ---------------------------------------
writeFileSync(join(here, "page.embed.html"), withFonts);

// --- Sortie 2 : document autonome ------------------------------------------
const title = withFonts.match(/<title>([^<]*)<\/title>/)?.[1] ?? "Ine4rt";
const body = withFonts.replace(/<title>[^<]*<\/title>\s*/, "");

const standalone = `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<meta name="description" content="Sites web sur mesure pour les indépendants et les petites entreprises qui n\u0027en ont pas encore. Nom de domaine, hébergement et entretien pris en charge.">
<meta name="theme-color" content="#121316">
<meta property="og:type" content="website">
<meta property="og:title" content="${title}">
<meta property="og:description" content="IneWeb conçoit des sites web sur mesure et en assure l'hébergement et le suivi. Les entreprises que nous accompagnons n'ont rien à gérer.">
<meta property="og:locale" content="fr_BE">
<link rel="icon" href="data:image/svg+xml,${encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="6" fill="#e8641c"/><text x="16" y="22" font-family="Helvetica,Arial,sans-serif" font-size="14" font-weight="700" fill="#180d02" text-anchor="middle">iW</text></svg>',
)}">
<style>
  *, *::before, *::after { box-sizing: border-box; }
  html { -webkit-text-size-adjust: 100%; scroll-behavior: smooth; }
  @media (prefers-reduced-motion: reduce) { html { scroll-behavior: auto; } }
</style>
</head>
<body>
${body}
</body>
</html>
`;

writeFileSync(join(here, "index.html"), standalone);

console.log(
  `✓ index.html (${Math.round(standalone.length / 1024)} Ko) et page.embed.html générés.`,
);
