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
 *   9. O resumo cabe em duas linhas (nº de processos e total das taxas), com uma
 *      coluna por indicador — total, períodos, atribuição, estados e tipos — e é
 *      a única parte congelada da folha; a tabela de técnicos tem metade das
 *      linhas de cada lado.
 *  10. Os grupos de 3 dias cobrem cada ficheiro exactamente uma vez, e o modo de
 *      um ficheiro por dia dá um grupo por cada dia.
 *  11. Um relatório preenchido com técnicos pode ser relido sem perder nada:
 *      mesmas linhas, mesma ordem, técnicos, estados e tipos preservados.
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
import { lerRelatorio } from '../assets/leitor-relatorio.js';

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

const porDia = agruparPorDiasConsecutivos(ficheiros, 1);
const diasDistintos = new Set(ficheiros.map((f) => f.dataIso)).size;
verificar(porDia.length === diasDistintos, 'modo "um ficheiro por dia" dá um grupo por dia',
  `${porDia.length} grupos para ${diasDistintos} dias`);
verificar(porDia.every((g) => g.datas.length === 1), 'cada grupo diário tem um só dia');
verificar(porDia.flatMap((g) => g.ficheiros).length === ficheiros.length,
  'nenhum ficheiro se perde no modo diário');

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

  // resumo compacto: cabeçalho + exactamente duas linhas de valores
  let linhaResumoN = 0;
  for (let r = 1; r < linhaCab; r++) {
    if (ws.getCell(r, 1).value === 'Nº de Processos') { linhaResumoN = r; break; }
  }
  verificar(linhaResumoN > 0, 'resumo com a linha "Nº de Processos"');
  const linhaResumoT = linhaResumoN + 1;
  const linhaResumoCab = linhaResumoN - 1;
  verificar(ws.getCell(linhaResumoT, 1).value === 'Total das Taxas',
    'resumo ocupa apenas duas linhas de valores', String(ws.getCell(linhaResumoT, 1).value));
  verificar(ws.getCell(linhaResumoCab, 1).value === 'Indicador',
    'cabeçalho do resumo por cima das duas linhas');

  const cabResumo = [];
  for (let c = 2; c <= 30; c++) {
    const v = ws.getCell(linhaResumoCab, c).value;
    if (v === null || v === undefined) break;
    cabResumo.push(typeof v === 'object' && v.formula ? `f:${v.formula}` : v);
  }
  verificar(cabResumo[0] === 'Total de processos', 'primeira coluna do resumo é o total');
  verificar(cabResumo.includes('Pago') && cabResumo.includes('Não Pago'),
    'estados no resumo', cabResumo.join(' | '));
  verificar(cabResumo.filter((v) => String(v).includes('MATCH')).length >= 3,
    'colunas de tipos no resumo', String(cabResumo.filter((v) => String(v).includes('MATCH')).length));
  verificar(cabResumo.includes('Outros tipos'), 'coluna de segurança "Outros tipos"');

  verificar(
    [linhaResumoN, linhaResumoT].every((r) => cabResumo.every((_, i) => {
      const v = ws.getCell(r, 2 + i).value;
      return v && typeof v === 'object' && v.formula;
    })),
    'todas as células do resumo são fórmulas'
  );

  // só o resumo fica congelado
  const vista = (ws.views || [])[0] || {};
  verificar(vista.state === 'frozen' && vista.ySplit === linhaResumoT,
    'congelamento apenas até à segunda linha do resumo', JSON.stringify(vista));
  verificar(vista.ySplit < linhaCab, 'a tabela de dados acompanha a rolagem');
  // A célula activa tem de ficar abaixo do corte, senão o painel de baixo rola
  // até ela e o resumo aparece repetido.
  verificar(vista.activeCell === `A${linhaResumoT + 1}` && vista.topLeftCell === `A${linhaResumoT + 1}`,
    'painel inferior começa abaixo da zona congelada',
    `activeCell=${vista.activeCell} topLeftCell=${vista.topLeftCell}`);

  // tabela de técnicos com metade das linhas de cada lado
  let linhaTec = 0;
  for (let r = 1; r < linhaCab; r++) {
    if (ws.getCell(r, 1).value === 'Técnico [6]') { linhaTec = r; break; }
  }
  verificar(linhaTec > 0, 'tabela de técnicos encontrada');
  verificar(ws.getCell(linhaTec, 6).value === 'Técnico [6]',
    'técnicos com uma metade de cada lado', String(ws.getCell(linhaTec, 6).value));

  /* ---- 11 — ida e volta: preencher e reler ------------------------ */
  {
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const folha = wb2.getWorksheet('Relatório');
    let rTec = 0;
    for (let r = 1; r < linhaCab; r++) {
      if (folha.getCell(r, 1).value === 'Técnico [6]') { rTec = r + 1; break; }
    }
    folha.getCell(rTec, 1).value = 'João';
    folha.getCell(rTec, 2).value = 'A';
    folha.getCell(rTec, 6).value = 'Ana';
    folha.getCell(rTec, 7).value = 'B';

    const equipa = ['João', 'Ana'];
    let i = 0;
    for (let r = linhaCab + 1; r <= folha.rowCount; r++) {
      if (folha.getCell(r, 5).value === null || folha.getCell(r, 5).value === undefined) continue;
      folha.getCell(r, 6).value = equipa[i % 2];
      if (i % 3 === 0) folha.getCell(r, 4).value = 'Não Pago';
      i++;
    }

    const relido = await lerRelatorio(ExcelJS, await wb2.xlsx.writeBuffer(), 'preenchido.xlsx');
    verificar(relido.registos.length === lidos.length,
      'relatório preenchido relido com todas as linhas',
      `${relido.registos.length} de ${lidos.length}`);
    verificar(
      relido.registos.every((r, k) => r.numeroDU === lidos[k].du),
      'ordem preservada ao reler o relatório'
    );
    verificar(relido.tecnicos.length === 2
      && relido.tecnicos.some((t) => t.nome === 'João' && t.tipo === 'A')
      && relido.tecnicos.some((t) => t.nome === 'Ana' && t.tipo === 'B'),
    'mapa de técnicos relido das duas metades',
    relido.tecnicos.map((t) => `${t.nome}=${t.tipo}`).join(', '));
    verificar(relido.registos.every((r) => r.tecnico === 'João' || r.tecnico === 'Ana'),
      'técnico de cada linha preservado');
    verificar(
      relido.registos.every((r) => r.tipo === (r.tecnico === 'João' ? 'A' : 'B')),
      'tipo recalculado a partir do técnico'
    );
    verificar(relido.registos.filter((r) => r.estado === 'Não Pago').length === Math.ceil(lidos.length / 3),
      'estados preservados ao reler',
      String(relido.registos.filter((r) => r.estado === 'Não Pago').length));
    verificar(relido.registos.every((r) => r.totalTaxas === lidos.find((l) => l.du === r.numeroDU).total),
      'totais preservados ao reler');
  }

  const dv = ws.getCell(linhaCab + 1, 4).dataValidation;
  verificar(
    dv && dv.type === 'list' && /Pago/.test(String(dv.formulae[0])) && /Não Pago/.test(String(dv.formulae[0])),
    'coluna Estado com lista pendente Pago / Não Pago',
    JSON.stringify(dv)
  );
  verificar(lidos.every((l) => l.tecnico === null || l.tecnico === undefined), 'Técnico [6] em branco');
  verificar(lidos.every((l) => l.tipo && typeof l.tipo === 'object' && l.tipo.formula),
    'Tipo [7] calculado por fórmula a partir do Técnico');

  // a ordem tem de ser exactamente a dos PDFs, sem reordenações
  const sequenciaEsperada = grupo.ficheiros.flatMap((f) => f.registos.map((r) => String(r.numeroDU)));
  const sequenciaLida = lidos.map((l) => l.du);
  const primeiraDiferenca = sequenciaEsperada.findIndex((v, i) => v !== sequenciaLida[i]);
  verificar(
    primeiraDiferenca === -1 && sequenciaLida.length === sequenciaEsperada.length,
    'linhas na ordem exacta dos PDFs',
    primeiraDiferenca >= 0
      ? `posição ${primeiraDiferenca + 1}: esperado ${sequenciaEsperada[primeiraDiferenca]}, `
        + `lido ${sequenciaLida[primeiraDiferenca]}`
      : ''
  );

  // separador de milhares e alinhamento do cabeçalho
  const fmt = ws.getCell(linhaResumoN, 2).numFmt || '';
  verificar(/^\[\$-4\d\d\]/.test(fmt), 'números com separador de milhares fixo', fmt);
  verificar((ws.getCell(1, 1).alignment || {}).horizontal === 'left'
    && (ws.getCell(2, 1).alignment || {}).horizontal === 'left',
  'título e subtítulo alinhados à esquerda');
}

console.log(`\n${verificacoes - falhas}/${verificacoes} verificações passaram.`);
if (falhas) {
  console.log(`${falhas} FALHA(S).`);
  process.exit(1);
}
console.log('Tudo certo.');
