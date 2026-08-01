/*
 * Junta tudo num único ficheiro. Lê src/pagina.html e substitui cada
 * referência local (<link href> e <script src>) pelo respectivo conteúdo,
 * produzindo o index.html que se abre com dois cliques ou se publica.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const raiz = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const molde = fs.readFileSync(path.join(raiz, 'src', 'pagina.html'), 'utf8');

function ler(rel) {
  return fs.readFileSync(path.join(raiz, rel), 'utf8');
}

let saida = molde;

// <link rel="stylesheet" href="assets/style.css"> -> <style>...</style>
saida = saida.replace(/<link[^>]*href="([^"]+)"[^>]*>/g, (m, href) => {
  if (/^https?:/.test(href)) return m;
  return `<style>\n${ler(href)}\n</style>`;
});

// <script src="vendor/jszip.min.js"></script> -> <script>...</script>
saida = saida.replace(/<script[^>]*src="([^"]+)"[^>]*><\/script>/g, (m, src) => {
  if (/^https?:/.test(src)) return m;
  return `<script>\n${ler(src)}\n</script>`;
});

const destino = path.join(raiz, 'index.html');
fs.writeFileSync(destino, saida);
const kb = (Buffer.byteLength(saida) / 1024).toFixed(0);
console.log(`index.html escrito (${kb} KB).`);
