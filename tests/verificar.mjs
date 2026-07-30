#!/usr/bin/env node
/**
 * verificar.mjs — testes de precisão sobre PDFs reais.
 *
 *   npm install
 *   node tests/verificar.mjs <pasta-com-pdfs>
 *
 * Cada PDF é lido com o mesmo extractor usado na aplicação web e submetido a
 * verificações independentes do próprio extractor:
 *
 *   1. O nº de registos é igual ao nº de marcadores "IMV" do documento — cada
 *      processo tem exactamente um, o que dá uma contagem de referência obtida
 *      sem usar as colunas.
 *   2. Todos os registos têm um Total das Taxas numérico e positivo.
 *   3. O total lido aparece textualmente na linha original do PDF.
 *   4. O total lido não é o Nº do DU nem o número da liquidação (colunas
 *      vizinhas), confirmando que não houve troca de colunas.
 *   5. A Data de Reg. de cada registo coincide com a data do nome do ficheiro.
 *   6. Não há Nº do DU repetido dentro do mesmo ficheiro.
 *   7. O Excel gerado contém exactamente os mesmos registos, nas 7 colunas e
 *      na ordem pedidas, com Estado preenchido, Técnico vazio e Tipo em fórmula.
 *   8. Cada PDF contribui com todas as suas linhas: uma linha por processo, sem
 *      agregações nem repetições, e cada linha identifica o PDF de origem.
 *   9. O resumo tem, à direita, as estatísticas por Estado (Pago / Não Pago) e por
 *      Tipo [7]; a coluna Estado tem lista pendente com esses valores; a tabela
 *      de técnicos tem metade das linhas de cada lado; e tudo o que está acima
 *      dos dados, incluindo o cabeçalho das colunas, fica congelado.
 *  10. Os grupos de 3 dias cobrem cada ficheiro exactamente uma vez.
 */

import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';

import {
  parseFileName,
  extractRowsFromDocument,
  itemsFromTextContent,
  groupIntoLines,
  agruparPorDiasConsecutivos,
  rotuloDoGrupo,
  dayNumber,
} from '../assets/parser.js';
import { criarWorkbook, COLUNAS, COLUNA_ORIGEM } from '../assets/report.js';

const pasta = process.argv[2];
if (!pasta) {
  console.error('Uso: node tests/verificar.mjs <pasta-com-pdfs>');
  process.exit(1);
}

let falhas = 0;
let verificacoes = 0;

function verificar(condicao, descricao, detalhe = '') {
  verificacoes++;
  if (!condicao) {
    falhas++;
    console.log(`   ✗ ${descricao}${detalhe ? ` — ${detalhe}` : ''}`);
  }
  return condicao;
}

/** Texto completo de cada linha do PDF, para conferir os valores lidos. */
async function linhasDeTexto(doc) {
  const linhas = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const pagina = await doc.getPage(p);
    const items = itemsFromTextContent(await pagina.getTextContent());
    for (const l of groupIntoLines(items)) {
      linhas.push(l.items.map((i) => i.text).join(' '));
    }
  }
  return linhas;
}

const nomes = (await fs.readdir(pasta)).filter((n) => /\.pdf$/i.test(n)).sort();
if (!nomes.length) {
  console.error(`Sem PDFs em ${pasta}`);
  process.exit(1);
}

const ficheiros = [];

for (const nome of nomes) {
  console.log(`\n▸ ${nome}`);
  const dados = new Uint8Array(await fs.readFile(path.join(pasta, nome)));
  const tarefa = pdfjsLib.getDocument({ data: dados });
  const doc = await tarefa.promise;
  const registos = await extractRowsFromDocument(doc);
  const linhas = await linhasDeTexto(doc);
  const textoTodo = linhas.join('\n');
  await tarefa.destroy();

  const { periodo, dataIso } = parseFileName(nome);
  ficheiros.push({ nome, periodo, dataIso, registos });

  // 1 — contagem de referência pelos marcadores IMV
  const marcadores = (textoTodo.match(/(^|\s)IMV(\s|$)/g) || []).length
    - (textoTodo.match(/igual a: IMV/g) || []).length;
  verificar(
    registos.length === marcadores,
    'nº de registos igual ao nº de processos do documento',
    `extraídos ${registos.length}, esperados ${marcadores}`
  );

  // 2 — totais válidos
  const semTotal = registos.filter((r) => !(typeof r.totalTaxas === 'number' && r.totalTaxas > 0));
  verificar(semTotal.length === 0, 'todos os registos têm Total das Taxas',
    semTotal.map((r) => r.numeroDU).join(', '));

  // 3 e 4 — o total está na linha certa e não é uma coluna vizinha
  let foraDaLinha = 0;
  let colunaTrocada = 0;
  let semLinha = 0;
  for (const r of registos) {
    const linha = linhas.find((l) => new RegExp(`(^|\\s)${r.numeroDU}\\s+${r.dataReg}\\s`).test(l));
    if (!linha) { semLinha++; continue; }
    if (!new RegExp(`(^|\\s)${r.totalTaxas}(\\s|$)`).test(linha)) foraDaLinha++;
    if (String(r.totalTaxas) === String(r.numeroDU)) colunaTrocada++;
    // a coluna seguinte ao total é "L" + nº de liquidação
    const m = new RegExp(`${r.totalTaxas}\\s+([A-Z])\\s+(\\d+)`).exec(linha);
    if (m && String(r.totalTaxas) === m[2]) colunaTrocada++;
  }
  verificar(semLinha === 0, 'cada registo foi localizado na página de origem', `${semLinha} sem linha`);
  verificar(foraDaLinha === 0, 'cada total aparece na linha original do processo', `${foraDaLinha} fora`);
  verificar(colunaTrocada === 0, 'o total não é o Nº do DU nem o nº de liquidação');

  // 5 — datas coerentes com o nome do ficheiro
  if (dataIso) {
    const diferentes = registos.filter((r) => r.dataReg !== dataIso);
    verificar(diferentes.length === 0, 'Data de Reg. coincide com a data do nome do ficheiro',
      [...new Set(diferentes.map((r) => r.dataReg))].join(', '));
  }
  verificar(!!periodo, 'Período identificado no nome do ficheiro');

  // 6 — DU sem repetições
  const unicos = new Set(registos.map((r) => r.numeroDU));
  verificar(unicos.size === registos.length, 'Nº do DU sem repetições',
    `${registos.length - unicos.size} repetido(s)`);

  console.log(`   ${registos.length} registos · período ${periodo || '?'} · `
    + `total ${registos.reduce((a, r) => a + (r.totalTaxas || 0), 0).toLocaleString('pt-PT')}`);
}

/* ---- 8 — agrupamento ---------------------------------------------- */
console.log('\n▸ Agrupamento em blocos de 3 dias seguidos');
const grupos = agruparPorDiasConsecutivos(ficheiros, 3);
for (const g of grupos) {
  console.log(`   ${rotuloDoGrupo(g)} → ${g.ficheiros.map((f) => f.nome).join(', ')}`);
  verificar(g.datas.length <= 3, 'no máximo 3 dias por grupo', `${g.datas.length} dias`);
  for (let i = 1; i < g.datas.length; i++) {
    verificar(
      dayNumber(g.datas[i]) === dayNumber(g.datas[i - 1]) + 1,
      'os dias de cada grupo são consecutivos',
      `${g.datas[i - 1]} → ${g.datas[i]}`
    );
  }
}
const agrupados = grupos.flatMap((g) => g.ficheiros);
verificar(agrupados.length === ficheiros.length, 'cada ficheiro entra exactamente num grupo');

/* ---- 7 — conteúdo do Excel ---------------------------------------- */
console.log('\n▸ Ficheiros Excel gerados');
for (const g of [...grupos, { datas: [], ficheiros, rotulo: 'Consolidado', consolidado: true }]) {
  const grupo = { ...g, rotulo: g.rotulo || rotuloDoGrupo(g) };
  const buffer = await criarWorkbook(ExcelJS, grupo, { estado: 'Pago', consolidado: g.consolidado });

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.getWorksheet('Relatório');
  console.log(`   ${grupo.rotulo}`);

  // localiza o cabeçalho da tabela de dados
  let linhaCab = 0;
  for (let r = 1; r <= ws.rowCount; r++) {
    if (ws.getCell(r, 1).value === COLUNAS[0] && ws.getCell(r, 2).value === COLUNAS[1]) { linhaCab = r; break; }
  }
  verificar(linhaCab > 0, 'tabela de dados encontrada');
  verificar(
    COLUNAS.every((c, i) => ws.getCell(linhaCab, i + 1).value === c),
    'as 7 colunas estão pela ordem pedida',
    COLUNAS.map((c, i) => ws.getCell(linhaCab, i + 1).value).join(' | ')
  );
  verificar(
    ws.getCell(linhaCab, COLUNAS.length + 1).value === COLUNA_ORIGEM,
    'coluna de origem (PDF) a seguir às 7 pedidas'
  );

  const esperados = grupo.ficheiros.flatMap((f) => f.registos.map((r) => ({
    periodo: f.periodo, dataReg: r.dataReg, total: r.totalTaxas, du: String(r.numeroDU),
    ficheiro: f.nome,
  })));

  const lidos = [];
  for (let r = linhaCab + 1; r <= ws.rowCount; r++) {
    const du = ws.getCell(r, 5).value;
    if (du === null || du === undefined || du === '') continue;
    const data = ws.getCell(r, 2).value;
    lidos.push({
      periodo: ws.getCell(r, 1).value,
      dataReg: data instanceof Date ? data.toISOString().slice(0, 10) : String(data),
      total: ws.getCell(r, 3).value,
      du: String(du),
      estado: ws.getCell(r, 4).value,
      tecnico: ws.getCell(r, 6).value,
      tipo: ws.getCell(r, 7).value,
      ficheiro: ws.getCell(r, 8).value,
    });
  }

  verificar(lidos.length === esperados.length, 'o Excel tem todos os registos',
    `${lidos.length} de ${esperados.length}`);

  const chave = (o) => `${o.periodo}|${o.dataReg}|${o.total}|${o.du}|${o.ficheiro}`;
  const conjunto = new Set(lidos.map(chave));
  const emFalta = esperados.filter((e) => !conjunto.has(chave(e)));
  verificar(emFalta.length === 0, 'todos os registos do PDF estão no Excel com os mesmos valores',
    emFalta.slice(0, 3).map(chave).join(' ; '));

  // cada PDF tem de contribuir com todas as suas linhas, uma a uma
  for (const f of grupo.ficheiros) {
    const doFicheiro = lidos.filter((l) => l.ficheiro === f.nome);
    verificar(
      doFicheiro.length === f.registos.length,
      `todas as linhas de ${f.nome} estão no Excel`,
      `${doFicheiro.length} de ${f.registos.length}`
    );
    console.log(`      ${f.nome}: ${doFicheiro.length} linha(s)`);
    const dusExcel = new Set(doFicheiro.map((l) => l.du));
    const emFaltaDU = f.registos.filter((r) => !dusExcel.has(String(r.numeroDU)));
    verificar(emFaltaDU.length === 0, `nenhum processo de ${f.nome} em falta`,
      emFaltaDU.map((r) => r.numeroDU).join(', '));
  }

  // uma linha por processo: nada agregado nem repetido
  const chavesUnicas = new Set(lidos.map((l) => `${l.ficheiro}|${l.du}`));
  verificar(chavesUnicas.size === lidos.length,
    'uma linha por processo, sem repetições nem agregações',
    `${lidos.length} linhas, ${chavesUnicas.size} processos distintos`);

  verificar(lidos.every((l) => l.estado === 'Pago'), 'Estado [4] preenchido com "Pago"');

  // estatísticas por Estado e por Tipo, à direita do resumo
  let linhaEstados = 0;
  for (let r = 1; r < linhaCab; r++) {
    if (ws.getCell(r, 5).value === 'Estado [4]') { linhaEstados = r; break; }
  }
  verificar(linhaEstados > 0, 'bloco de estados à direita do resumo');
  verificar(ws.getCell(linhaEstados, 1).value === 'Indicador',
    'estados na mesma linha de cabeçalho do resumo');
  verificar(ws.getCell(linhaEstados, 9).value === 'Tipo [7]',
    'bloco de tipos à direita dos estados', String(ws.getCell(linhaEstados, 9).value));

  if (linhaEstados) {
    const rotulos = [ws.getCell(linhaEstados + 1, 5).value, ws.getCell(linhaEstados + 2, 5).value];
    verificar(rotulos[0] === 'Pago' && rotulos[1] === 'Não Pago',
      'estados Pago e Não Pago listados', rotulos.join(' | '));
    verificar(
      [1, 2].every((i) => {
        const n = ws.getCell(linhaEstados + i, 6).value;
        const t = ws.getCell(linhaEstados + i, 7).value;
        return n && n.formula && n.formula.startsWith('COUNTIF')
          && t && t.formula && t.formula.startsWith('SUMIF');
      }),
      'estados contados e somados por fórmula'
    );
    verificar(
      [1, 2].every((i) => {
        const tipo = ws.getCell(linhaEstados + i, 9).value;
        const n = ws.getCell(linhaEstados + i, 10).value;
        return tipo && tipo.formula && tipo.formula.includes('MATCH')
          && n && n.formula && n.formula.includes('COUNTIF');
      }),
      'tipos distintos listados e contados por fórmula'
    );
  }

  // tabela de técnicos com metade das linhas de cada lado
  let linhaTec = 0;
  for (let r = 1; r < linhaCab; r++) {
    if (ws.getCell(r, 1).value === 'Técnico [6]') { linhaTec = r; break; }
  }
  verificar(linhaTec > 0, 'tabela de técnicos encontrada');
  verificar(ws.getCell(linhaTec, 6).value === 'Técnico [6]',
    'técnicos com uma metade de cada lado', String(ws.getCell(linhaTec, 6).value));

  // o topo fica fixo, incluindo o cabeçalho das colunas
  const vista = (ws.views || [])[0] || {};
  verificar(vista.state === 'frozen' && vista.ySplit === linhaCab,
    'topo e cabeçalho dos dados congelados', JSON.stringify(vista));

  const dv = ws.getCell(linhaCab + 1, 4).dataValidation;
  verificar(
    dv && dv.type === 'list' && /Pago/.test(String(dv.formulae[0])) && /Não Pago/.test(String(dv.formulae[0])),
    'coluna Estado com lista pendente Pago / Não Pago',
    JSON.stringify(dv)
  );
  verificar(lidos.every((l) => l.tecnico === null || l.tecnico === undefined), 'Técnico [6] em branco');
  verificar(lidos.every((l) => l.tipo && typeof l.tipo === 'object' && l.tipo.formula),
    'Tipo [7] calculado por fórmula a partir do Técnico');

  const ordenado = lidos.every((l, i) => i === 0 || lidos[i - 1].dataReg <= l.dataReg);
  verificar(ordenado, 'registos ordenados por data');
}

console.log(`\n${verificacoes - falhas}/${verificacoes} verificações passaram.`);
if (falhas) {
  console.log(`${falhas} FALHA(S).`);
  process.exit(1);
}
console.log('Tudo certo.');
