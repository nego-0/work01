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
 *   7. O Excel gerado contém exactamente os mesmos registos, nas 7 colunas
 *      pedidas e por ordem ascendente de data, com Estado preenchido, Técnico
 *      vazio e Tipo em fórmula.
 *   8. Cada PDF contribui com todas as suas linhas: uma linha por processo, sem
 *      agregações nem repetições, e cada linha identifica o PDF de origem.
 *   9. O resumo cabe em duas linhas (nº de processos e total das taxas), com uma
 *      coluna por indicador — total, períodos, atribuição, estados e tipos — e é
 *      a única parte congelada da folha; os técnicos ocupam três tabelas
 *      coladas lado a lado.
 *  10. Os grupos de 3 dias cobrem cada ficheiro exactamente uma vez, e o modo de
 *      um ficheiro por dia dá um grupo por cada dia.
 *  11. Um relatório preenchido com técnicos pode ser relido sem perder nada:
 *      mesmas linhas, mesma ordem, técnicos, estados e tipos preservados.
 *  12. A junção de relatórios apanha as repetições (mesmo Nº do DU na mesma
 *      data), mantém a versão pedida por cada política e não duplica linhas.
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
import { lerRelatorio, juntarRegistos, chaveRegisto } from '../assets/leitor-relatorio.js';

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

  // o resumo conta muito além dos dados (linha 999), para quem acrescentar
  // linhas à mão não ter de esticar as fórmulas.
  const fTotal = String((ws.getCell(linhaResumoN, 2).value || {}).formula || '');
  const mLimite = fTotal.match(/:\$?[A-Z]\$?(\d+)\)/);
  verificar(mLimite && Number(mLimite[1]) >= 999,
    'o resumo conta até à linha 999 (dados acrescentados à mão contam sozinhos)', fTotal);
  // "Por atribuir" usa SUMPRODUCT com máscara (du<>"")*(tecnico=""): as muitas
  // linhas em branco entre os dados e a 999 não podem contar como por atribuir.
  const formulasResumoN = [];
  for (let c = 2; c <= 40; c++) {
    const v = ws.getCell(linhaResumoN, c).value;
    if (v === null || v === undefined) break;
    formulasResumoN.push(String(v.formula || ''));
  }
  verificar(
    formulasResumoN.some((f) => /SUMPRODUCT/.test(f) && /<>""/.test(f) && /=""/.test(f)),
    'contagem "por atribuir" imune às linhas em branco até 999 (SUMPRODUCT)'
  );
  // as linhas livres até 999 já trazem a fórmula do Tipo, para o cálculo automático
  const fTipoLivre = String((ws.getCell(999, 7).value || {}).formula || '');
  verificar(/VLOOKUP/.test(fTipoLivre),
    'linhas livres até 999 já trazem a fórmula do Tipo', fTipoLivre);

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

  // três tabelas de técnicos coladas lado a lado (A, F, K)
  let linhaTec = 0;
  for (let r = 1; r < linhaCab; r++) {
    if (ws.getCell(r, 1).value === 'Técnico [6]') { linhaTec = r; break; }
  }
  verificar(linhaTec > 0, 'tabela de técnicos encontrada');
  verificar(
    [1, 6, 11].every((c) => ws.getCell(linhaTec, c).value === 'Técnico [6]'),
    'três tabelas de técnicos lado a lado (A, F, K)',
    [1, 6, 11].map((c) => ws.getCell(linhaTec, c).value).join(' | ')
  );

  // o Tipo procura o técnico nas três tabelas por INTERVALO — é isto que deixa
  // inserir linhas nas tabelas sem partir a lógica.
  const fTipo = String((ws.getCell(linhaCab + 1, 7).value || {}).formula || '');
  verificar(
    (fTipo.match(/VLOOKUP/g) || []).length === 3
      && /\$A\$\d+:\$B\$\d+/.test(fTipo)
      && /\$F\$\d+:\$G\$\d+/.test(fTipo)
      && /\$K\$\d+:\$L\$\d+/.test(fTipo),
    'Tipo procurado nas três tabelas por intervalo', fTipo
  );

  // as contagens por técnico usam o mesmo intervalo de dados nas três tabelas,
  // por isso nenhum processo pode ser contado a dois técnicos.
  const rangesTec = new Set();
  for (const c of [1, 6, 11]) {
    const f = String((ws.getCell(linhaTec + 1, c + 2).value || {}).formula || '');
    const m = f.match(/COUNTIF\((\$F\$\d+:\$F\$\d+)/);
    if (m) rangesTec.add(m[1]);
  }
  verificar(rangesTec.size === 1,
    'as três tabelas contam sobre o mesmo intervalo de dados',
    [...rangesTec].join(' / '));

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
    folha.getCell(rTec, 11).value = 'Rui';
    folha.getCell(rTec, 12).value = 'C';

    const equipa = ['João', 'Ana', 'Rui'];
    let i = 0;
    for (let r = linhaCab + 1; r <= folha.rowCount; r++) {
      if (folha.getCell(r, 5).value === null || folha.getCell(r, 5).value === undefined) continue;
      folha.getCell(r, 6).value = equipa[i % 3];
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
    verificar(relido.tecnicos.length === 3
      && relido.tecnicos.some((t) => t.nome === 'João' && t.tipo === 'A')
      && relido.tecnicos.some((t) => t.nome === 'Ana' && t.tipo === 'B')
      && relido.tecnicos.some((t) => t.nome === 'Rui' && t.tipo === 'C'),
    'mapa de técnicos relido das três tabelas',
    relido.tecnicos.map((t) => `${t.nome}=${t.tipo}`).join(', '));
    verificar(relido.registos.every((r) => ['João', 'Ana', 'Rui'].includes(r.tecnico)),
      'técnico de cada linha preservado');
    const tipoEsperado = { João: 'A', Ana: 'B', Rui: 'C' };
    verificar(
      relido.registos.every((r) => r.tipo === tipoEsperado[r.tecnico]),
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

  // a ordem tem de ser ascendente por Data de Reg.
  const datasLidas = lidos.map((l) => l.dataReg);
  const foraDeOrdem = datasLidas.findIndex((d, i) => i > 0 && datasLidas[i - 1] > d);
  verificar(
    foraDeOrdem === -1,
    'linhas por ordem ascendente de Data de Reg.',
    foraDeOrdem >= 0 ? `${datasLidas[foraDeOrdem - 1]} antes de ${datasLidas[foraDeOrdem]}` : ''
  );
  // os mesmos processos, só reordenados: mesmo conjunto de (data|du)
  const espSeq = grupo.ficheiros
    .flatMap((f) => f.registos.map((r) => `${r.dataReg}|${r.numeroDU}`)).sort();
  const lidSeq = lidos.map((l) => `${l.dataReg}|${l.du}`).sort();
  verificar(
    espSeq.length === lidSeq.length && espSeq.every((v, i) => v === lidSeq[i]),
    'a reordenação não perde nem troca nenhum processo de dia'
  );
  // dentro de cada dia, mantém-se a ordem original dos PDFs (ordenação estável)
  const seqPorDia = (pares) => {
    const m = new Map();
    for (const [data, du] of pares) {
      if (!m.has(data)) m.set(data, []);
      m.get(data).push(String(du));
    }
    return m;
  };
  const espDia = seqPorDia(grupo.ficheiros.flatMap((f) => f.registos.map((r) => [r.dataReg, r.numeroDU])));
  const lidDia = seqPorDia(lidos.map((l) => [l.dataReg, l.du]));
  let ordemEstavel = true;
  for (const [data, seq] of espDia) {
    const lida = lidDia.get(data) || [];
    if (seq.length !== lida.length || seq.some((v, i) => v !== lida[i])) ordemEstavel = false;
  }
  verificar(ordemEstavel, 'dentro de cada dia mantém-se a ordem original dos PDFs');

  // separador de milhares e alinhamento do cabeçalho
  const fmt = ws.getCell(linhaResumoN, 2).numFmt || '';
  verificar(/^\[\$-4\d\d\]/.test(fmt), 'números com separador de milhares fixo', fmt);
  verificar((ws.getCell(1, 1).alignment || {}).horizontal === 'left'
    && (ws.getCell(2, 1).alignment || {}).horizontal === 'left',
  'título e subtítulo alinhados à esquerda');
}

/* ---- 12 — junção com redundâncias --------------------------------- */
console.log('\n▸ Junção de relatórios');
{
  const primeiro = ficheiros[0];
  const segundo = ficheiros[1];

  // O mesmo relatório duas vezes: tudo repetido, nada acrescentado.
  const comoEstao = primeiro.registos.map((r) => ({ ...r, tecnico: '', estado: 'Pago' }));
  const preenchidos = primeiro.registos.map((r) => ({ ...r, tecnico: 'Ana', tipo: 'A', estado: 'Não Pago' }));

  const iguais = juntarRegistos([
    { origem: 'a.xlsx', registos: comoEstao },
    { origem: 'b.xlsx', registos: comoEstao },
  ], 'completo');
  verificar(iguais.registos.length === comoEstao.length,
    'o mesmo relatório duas vezes não duplica linhas',
    `${iguais.registos.length} de ${comoEstao.length}`);
  verificar(iguais.duplicados.length === comoEstao.length && iguais.duplicados.every((d) => d.iguais),
    'todas as repetições são detectadas e marcadas como iguais');

  // Política "a mais preenchida": ganha a versão com técnico.
  const completo = juntarRegistos([
    { origem: 'vazio.xlsx', registos: comoEstao },
    { origem: 'cheio.xlsx', registos: preenchidos },
  ], 'completo');
  verificar(completo.registos.every((r) => r.tecnico === 'Ana'),
    'política "mais preenchida" fica com a versão que tem técnico');
  verificar(completo.duplicados.every((d) => !d.iguais),
    'repetições com dados diferentes são assinaladas');

  // Política "a que já estava": ganha a primeira fonte, mesmo sendo mais vazia.
  const daBase = juntarRegistos([
    { origem: 'vazio.xlsx', registos: comoEstao },
    { origem: 'cheio.xlsx', registos: preenchidos },
  ], 'base');
  verificar(daBase.registos.every((r) => !r.tecnico), 'política "a que já estava" mantém a primeira');

  // Política "a nova": ganha a última fonte.
  const daNova = juntarRegistos([
    { origem: 'cheio.xlsx', registos: preenchidos },
    { origem: 'vazio.xlsx', registos: comoEstao },
  ], 'novo');
  verificar(daNova.registos.every((r) => !r.tecnico), 'política "a acrescentada" substitui a anterior');

  // Relatórios sem processos em comum somam-se sem repetições.
  const somados = juntarRegistos([
    { origem: 'a.xlsx', registos: primeiro.registos },
    { origem: 'b.xlsx', registos: segundo.registos },
  ], 'completo');
  verificar(somados.registos.length === primeiro.registos.length + segundo.registos.length,
    'relatórios de dias diferentes somam-se todos',
    `${somados.registos.length} de ${primeiro.registos.length + segundo.registos.length}`);
  verificar(somados.duplicados.length === 0, 'sem repetições entre dias diferentes');
  verificar(
    somados.registos.every((r, i) => r.numeroDU
      === [...primeiro.registos, ...segundo.registos][i].numeroDU),
    'a ordem das fontes é respeitada na junção'
  );

  // A chave é o Nº do DU dentro da data: o mesmo DU noutro dia é outro processo.
  const outroDia = primeiro.registos.map((r) => ({ ...r, dataReg: '2026-12-31' }));
  const separados = juntarRegistos([
    { origem: 'a.xlsx', registos: primeiro.registos },
    { origem: 'c.xlsx', registos: outroDia },
  ], 'completo');
  verificar(separados.duplicados.length === 0 && separados.registos.length === primeiro.registos.length * 2,
    'o mesmo Nº do DU noutra data conta como processo diferente');
  verificar(chaveRegisto(primeiro.registos[0]) !== chaveRegisto(outroDia[0]),
    'a chave inclui a data de registo');

  console.log(`   ${primeiro.registos.length} + ${segundo.registos.length} linhas usadas nos cenários`);
}

console.log(`\n${verificacoes - falhas}/${verificacoes} verificações passaram.`);
if (falhas) {
  console.log(`${falhas} FALHA(S).`);
  process.exit(1);
}
console.log('Tudo certo.');
