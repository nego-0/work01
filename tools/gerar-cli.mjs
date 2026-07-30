#!/usr/bin/env node
/**
 * gerar-cli.mjs — a mesma lógica da aplicação web, executada na linha de comandos.
 * Útil para testes e para processar pastas inteiras sem abrir o browser.
 *
 *   npm install
 *   node tools/gerar-cli.mjs <pasta-com-pdfs> [pasta-de-saida]
 */

import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';

import {
  parseFileName,
  extractRowsFromDocument,
  agruparPorDiasConsecutivos,
  nomeDoGrupo,
  rotuloDoGrupo,
  formatDatePt,
} from '../assets/parser.js';
import { criarWorkbook } from '../assets/report.js';

const entrada = process.argv[2];
const saida = process.argv[3] || 'saida';

if (!entrada) {
  console.error('Uso: node tools/gerar-cli.mjs <pasta-com-pdfs> [pasta-de-saida]');
  process.exit(1);
}

const nomes = (await fs.readdir(entrada)).filter((n) => /\.pdf$/i.test(n)).sort();
if (!nomes.length) {
  console.error(`Sem PDFs em ${entrada}`);
  process.exit(1);
}

const ficheiros = [];
for (const nome of nomes) {
  const dados = new Uint8Array(await fs.readFile(path.join(entrada, nome)));
  const tarefa = pdfjsLib.getDocument({ data: dados });
  const doc = await tarefa.promise;
  const registos = await extractRowsFromDocument(doc);
  await tarefa.destroy();

  const { periodo, dataIso } = parseFileName(nome);
  let data = dataIso;
  if (!data && registos.length) {
    const contagem = {};
    for (const r of registos) contagem[r.dataReg] = (contagem[r.dataReg] || 0) + 1;
    data = Object.entries(contagem).sort((a, b) => b[1] - a[1])[0][0];
  }

  const soma = registos.reduce((a, r) => a + (r.totalTaxas || 0), 0);
  console.log(
    `${nome}\n  período=${periodo || '?'} data=${data ? formatDatePt(data) : '?'} `
    + `registos=${registos.length} total=${soma.toLocaleString('pt-PT')}`
  );
  ficheiros.push({ nome, periodo, dataIso: data, registos });
}

// MODO=3 (por omissão) | MODO=1 (um ficheiro por dia) | MODO=consolidado
const modo = process.env.MODO || process.env.DIAS || '3';
const grupos = modo === 'consolidado'
  ? []
  : agruparPorDiasConsecutivos(ficheiros, Number(modo) || 3);
await fs.mkdir(saida, { recursive: true });

for (const g of grupos) {
  const grupo = { ...g, rotulo: rotuloDoGrupo(g) };
  const buffer = await criarWorkbook(ExcelJS, grupo, {});
  const destino = path.join(saida, `${nomeDoGrupo(g)}.xlsx`);
  await fs.writeFile(destino, Buffer.from(buffer));
  console.log(`→ ${destino} (${grupo.ficheiros.reduce((a, f) => a + f.registos.length, 0)} registos)`);
}

const consolidado = { rotulo: 'Consolidado', ficheiros };
const buffer = await criarWorkbook(ExcelJS, consolidado, { consolidado: true });
const destino = path.join(saida, 'Relatorio_Consolidado.xlsx');
await fs.writeFile(destino, Buffer.from(buffer));
console.log(`→ ${destino} (${ficheiros.reduce((a, f) => a + f.registos.length, 0)} registos)`);
