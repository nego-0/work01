#!/usr/bin/env node
/**
 * construir.mjs — junta tudo num único ficheiro HTML.
 *
 *   node tools/construir.mjs [destino]      (por omissão: index.html)
 *
 * O resultado não depende de mais nenhum ficheiro: estilos, bibliotecas e
 * código vão embutidos. Funciona servido por HTTP e também aberto
 * directamente do disco.
 *
 * Duas transformações merecem explicação:
 *
 *  - o pdf.js é um módulo ES; o `export{a as b,...}` final é convertido num
 *    objecto (`const __pdfjs = {b: a, ...}`) para o resto do código lhe aceder
 *    sem precisar de importar nada;
 *  - o worker do pdf.js é embrulhado numa função e registado em
 *    `globalThis.pdfjsWorker`, que é onde o pdf.js o procura antes de tentar
 *    buscá-lo à rede. Assim não há segundo ficheiro nem Web Worker — tudo
 *    corre no mesmo sítio, o que também evita os bloqueios de `file://`.
 */

import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const raiz = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const destino = path.resolve(raiz, process.argv[2] || 'index.html');

const ler = (p) => fs.readFile(path.join(raiz, p), 'utf8');

/** Converte o `export{...}` final de um módulo num objecto com esse nome. */
function exportacoesParaObjecto(codigo, nomeObjecto) {
  const m = /export\s*\{([^}]*)\}\s*;?\s*$/.exec(codigo.trimEnd());
  if (!m) throw new Error('não encontrei a lista de exportações');
  const pares = m[1].split(',').map((p) => p.trim()).filter(Boolean).map((p) => {
    const [local, exportado] = p.split(/\s+as\s+/).map((x) => x.trim());
    return `${JSON.stringify(exportado || local)}: ${local}`;
  });
  return `${codigo.slice(0, m.index)}\nconst ${nomeObjecto} = {${pares.join(',')}};\n`;
}

/** Tira os `import`/`export` de um módulo nosso, para poder ser concatenado. */
function semImportsNemExports(codigo) {
  return codigo
    .replace(/^\s*import[^;]*;\s*$/gm, '')
    .replace(/^export\s+(const|function|async function|class|let)\b/gm, '$1');
}

const [estilos, pagina, exceljs, jszip, pdf, pdfWorker, parser, report, leitor, app] =
  await Promise.all([
    ler('assets/style.css'),
    ler('src/pagina.html'),
    ler('vendor/exceljs.min.js'),
    ler('vendor/jszip.min.js'),
    ler('vendor/pdf.min.mjs'),
    ler('vendor/pdf.worker.min.mjs'),
    ler('assets/parser.js'),
    ler('assets/report.js'),
    ler('assets/leitor-relatorio.js'),
    ler('assets/app.js'),
  ]);

const bibliotecas = [
  '/* ExcelJS — https://github.com/exceljs/exceljs (MIT) */',
  exceljs,
  '\n/* JSZip — https://stuk.github.io/jszip/ (MIT) */',
  jszip,
].join('\n');

const worker = exportacoesParaObjecto(pdfWorker, '__pdfjsWorker')
  + '\nglobalThis.pdfjsWorker = __pdfjsWorker;\n';

const aplicacao = [
  '/* pdf.js — https://mozilla.github.io/pdf.js/ (Apache-2.0) */',
  exportacoesParaObjecto(pdf, '__pdfjs'),
  '(() => {',
  worker,
  '})();',
  '{',
  'const pdfjsLib = __pdfjs;',
  ...[parser, report, leitor, app].map(semImportsNemExports),
  '}',
].join('\n');

const html = pagina
  .replace('/*ESTILOS*/', () => estilos)
  .replace('/*BIBLIOTECAS*/', () => bibliotecas)
  .replace('/*APLICACAO*/', () => aplicacao);

await fs.writeFile(destino, html);
const kb = Math.round((await fs.stat(destino)).size / 1024);
console.log(`${path.relative(raiz, destino)} — ${kb.toLocaleString('pt-PT')} KB`);
