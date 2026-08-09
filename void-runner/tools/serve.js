import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { extname, join, normalize } from 'node:path';

/**
 * Serveur de développement minimal.
 * Le projet n'a aucune étape de build : ouvrir une page suffit. Ce serveur
 * n'existe que parce que les modules ES ne se chargent pas en file://.
 */
const ROOT = new URL('..', import.meta.url).pathname;
const PORT = Number(process.env.PORT || 8080);
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.webmanifest': 'application/manifest+json', '.svg': 'image/svg+xml',
};

createServer(async (req, res) => {
  try {
    let p = decodeURIComponent(req.url.split('?')[0]);
    if (p === '/') p = '/index.html';
    const file = join(ROOT, normalize(p).replace(/^(\.\.[/\\])+/, ''));
    const s = await stat(file);
    if (!s.isFile()) throw new Error('not a file');
    res.writeHead(200, {
      'content-type': MIME[extname(file)] || 'application/octet-stream',
      'cache-control': 'no-cache',
    });
    res.end(await readFile(file));
  } catch {
    res.writeHead(404, { 'content-type': 'text/plain' });
    res.end('404');
  }
}).listen(PORT, () => console.log(`VOID RUNNER → http://localhost:${PORT}`));
