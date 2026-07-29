/**
 * report.js — construção dos ficheiros Excel (.xlsx).
 *
 * A biblioteca ExcelJS é injectada como argumento para o módulo funcionar tanto
 * no browser (global `ExcelJS`) como em Node (import 'exceljs').
 *
 * Estrutura de cada folha "Relatório":
 *   1. Resumo (totais e contagens por período)
 *   2. Técnicos — área editável que define o Tipo [7] em função do Técnico [6]
 *      e mostra, por técnico, o nº de processos e o total das taxas
 *   3. Estatísticas por Tipo
 *   4. Dados: Período | Data de Reg. | Total das Taxas | Estado | Nº do DU |
 *             Técnico | Tipo
 *
 * Tudo é feito com fórmulas nativas do Excel: assim que o utilizador escrever o
 * nome do técnico numa linha de dados, o Tipo é preenchido automaticamente e as
 * estatísticas no topo actualizam-se sozinhas.
 */

import { formatDatePt, dayNumber } from './parser.js';

export const COLUNAS = [
  'Período',
  'Data de Reg.',
  'Total das Taxas',
  'Estado',
  'Nº do DU',
  'Técnico',
  'Tipo',
];

const COR_TITULO = 'FF1F3864';
const COR_SECCAO = 'FF2E75B6';
const COR_CABECALHO = 'FFD9E2F3';
const COR_INPUT = 'FFFFF2CC';
const COR_CALC = 'FFF2F2F2';
const BORDA = { style: 'thin', color: { argb: 'FFBFBFBF' } };

/** Serial do Excel para uma data ISO (independente do fuso horário). */
function excelSerial(iso) {
  return dayNumber(iso) + 25569;
}

function aplicarBorda(row, deCol, ateCol) {
  for (let c = deCol; c <= ateCol; c++) {
    row.getCell(c).border = { top: BORDA, left: BORDA, bottom: BORDA, right: BORDA };
  }
}

function tituloSeccao(ws, linha, texto, nota) {
  const cell = ws.getCell(linha, 1);
  cell.value = texto;
  cell.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_SECCAO } };
  cell.alignment = { vertical: 'middle' };
  ws.mergeCells(linha, 1, linha, 7);
  ws.getRow(linha).height = 20;
  if (nota) {
    const c = ws.getCell(linha + 1, 1);
    c.value = nota;
    c.font = { italic: true, size: 9, color: { argb: 'FF595959' } };
    ws.mergeCells(linha + 1, 1, linha + 1, 7);
    return linha + 2;
  }
  return linha + 1;
}

function linhaCabecalho(ws, linha, valores) {
  const row = ws.getRow(linha);
  valores.forEach((v, i) => {
    const cell = row.getCell(i + 1);
    cell.value = v;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CABECALHO } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  aplicarBorda(row, 1, valores.length);
  return linha + 1;
}

/**
 * Cria a folha "Relatório" com estatísticas + dados.
 *
 * @param {object} ws            worksheet ExcelJS
 * @param {object} params
 * @param {string} params.titulo
 * @param {string} params.subtitulo
 * @param {Array}  params.registos  [{periodo, dataReg, totalTaxas, numeroDU}]
 * @param {Array}  params.tecnicos  [{nome, tipo}] pré-preenchidos (pode ser [])
 * @param {string} params.estado    valor por omissão do campo Estado [4]
 * @param {number} params.linhasTecnicos nº de linhas da área de configuração
 */
function construirFolhaRelatorio(ws, params) {
  const { titulo, subtitulo, registos, tecnicos = [], estado = 'Pago' } = params;
  const nLinhasTec = Math.max(params.linhasTecnicos || 12, tecnicos.length + 4);

  ws.columns = [
    { width: 16 }, { width: 14 }, { width: 18 }, { width: 12 },
    { width: 14 }, { width: 26 }, { width: 16 },
  ];

  /* ---- Título ---------------------------------------------------- */
  ws.mergeCells(1, 1, 1, 7);
  const t = ws.getCell(1, 1);
  t.value = titulo;
  t.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_TITULO } };
  t.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 26;

  ws.mergeCells(2, 1, 2, 7);
  const s = ws.getCell(2, 1);
  s.value = subtitulo;
  s.font = { size: 10, color: { argb: 'FF404040' } };
  s.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  ws.getRow(2).height = 30;

  /* ---- Reserva de posições (os dados ficam no fim da folha) ------- */
  const periodos = [...new Set(registos.map((r) => r.periodo).filter(Boolean))];

  let linha = 4;
  const linhaResumoTitulo = linha;
  linha = tituloSeccao(ws, linha, 'RESUMO');
  const linhaResumoCab = linha;
  linha = linhaCabecalho(ws, linha, ['Indicador', 'Nº de Processos', 'Total das Taxas']);
  const primResumo = linha;
  const nResumo = 2 + periodos.length + (periodos.length ? 1 : 0);
  linha += nResumo;
  linha += 1; // espaço

  const linhaTecTitulo = linha;
  linha = tituloSeccao(
    ws,
    linha,
    'TÉCNICOS — DEFINIR O TIPO [7] EM FUNÇÃO DO TÉCNICO [6]',
    'Escreva o nome do técnico e o tipo correspondente nas células amarelas. '
      + 'Na tabela de dados basta escolher o técnico: o tipo é preenchido automaticamente e '
      + 'as estatísticas actualizam-se.'
  );
  const linhaTecCab = linha;
  linha = linhaCabecalho(ws, linha, ['Técnico [6]', 'Tipo [7]', 'Nº de Processos', 'Total das Taxas']);
  const primTec = linha;
  const ultTec = primTec + nLinhasTec - 1;
  linha = ultTec + 1;
  linha += 1;

  const linhaTipoTitulo = linha;
  linha = tituloSeccao(ws, linha, 'ESTATÍSTICAS POR TIPO [7]');
  const linhaTipoCab = linha;
  linha = linhaCabecalho(ws, linha, ['Tipo [7]', 'Nº de Processos', 'Total das Taxas']);
  const primTipo = linha;
  const ultTipo = primTipo + nLinhasTec - 1;
  linha = ultTipo + 1;
  linha += 1;

  const linhaDadosTitulo = linha;
  linha = tituloSeccao(ws, linha, 'DADOS');
  const linhaDadosCab = linha;
  linha = linhaCabecalho(ws, linha, COLUNAS);
  const primDados = linha;
  const ultDados = primDados + Math.max(registos.length, 1) - 1;

  /* Intervalos usados pelas fórmulas ------------------------------- */
  const R = {
    periodo: `$A$${primDados}:$A$${ultDados}`,
    total: `$C$${primDados}:$C$${ultDados}`,
    du: `$E$${primDados}:$E$${ultDados}`,
    tecnico: `$F$${primDados}:$F$${ultDados}`,
    tipo: `$G$${primDados}:$G$${ultDados}`,
    mapa: `$A$${primTec}:$B$${ultTec}`,
    mapaTec: `$A$${primTec}:$A$${ultTec}`,
  };

  /* ---- Secção 1: Resumo ------------------------------------------ */
  const resumo = [];
  resumo.push({
    rotulo: 'Total de processos',
    n: `COUNTA(${R.du})`,
    soma: `SUM(${R.total})`,
    destaque: true,
  });
  for (const p of periodos) {
    resumo.push({
      rotulo: p,
      n: `COUNTIF(${R.periodo},"${p}")`,
      soma: `SUMIF(${R.periodo},"${p}",${R.total})`,
    });
  }
  if (periodos.length) {
    resumo.push({
      rotulo: 'Processos com técnico atribuído',
      n: `COUNTA(${R.du})-COUNTBLANK(${R.tecnico})`,
      soma: `SUMIF(${R.tecnico},"<>",${R.total})`,
    });
  }
  resumo.push({
    rotulo: 'Processos por atribuir',
    n: `COUNTBLANK(${R.tecnico})`,
    soma: `SUM(${R.total})-SUMIF(${R.tecnico},"<>",${R.total})`,
  });

  resumo.forEach((item, i) => {
    const row = ws.getRow(primResumo + i);
    row.getCell(1).value = item.rotulo;
    row.getCell(2).value = { formula: item.n };
    row.getCell(3).value = { formula: item.soma };
    row.getCell(2).numFmt = '#,##0';
    row.getCell(3).numFmt = '#,##0';
    if (item.destaque) {
      for (let c = 1; c <= 3; c++) row.getCell(c).font = { bold: true };
    }
    for (let c = 1; c <= 3; c++) {
      row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CALC } };
    }
    aplicarBorda(row, 1, 3);
  });

  /* ---- Secção 2: Técnicos (mapa + estatísticas) ------------------ */
  for (let i = 0; i < nLinhasTec; i++) {
    const r = primTec + i;
    const row = ws.getRow(r);
    const cfg = tecnicos[i];

    row.getCell(1).value = cfg ? cfg.nome : null;
    row.getCell(2).value = cfg ? cfg.tipo : null;
    for (const c of [1, 2]) {
      row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_INPUT } };
    }

    row.getCell(3).value = { formula: `IF($A${r}="","",COUNTIF(${R.tecnico},$A${r}))` };
    row.getCell(4).value = { formula: `IF($A${r}="","",SUMIF(${R.tecnico},$A${r},${R.total}))` };
    row.getCell(3).numFmt = '#,##0';
    row.getCell(4).numFmt = '#,##0';
    aplicarBorda(row, 1, 4);
  }

  /* ---- Secção 3: Estatísticas por tipo --------------------------- */
  for (let i = 0; i < nLinhasTec; i++) {
    const r = primTipo + i;
    const rMapa = primTec + i;
    const row = ws.getRow(r);

    // Lista de tipos distintos: mostra o valor apenas na 1ª ocorrência.
    row.getCell(1).value = {
      formula:
        `IF($B${rMapa}="","",IF(COUNTIF($B$${primTec}:$B${rMapa},$B${rMapa})=1,$B${rMapa},""))`,
    };
    row.getCell(2).value = { formula: `IF($A${r}="","",COUNTIF(${R.tipo},$A${r}))` };
    row.getCell(3).value = { formula: `IF($A${r}="","",SUMIF(${R.tipo},$A${r},${R.total}))` };
    row.getCell(2).numFmt = '#,##0';
    row.getCell(3).numFmt = '#,##0';
    for (let c = 1; c <= 3; c++) {
      row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CALC } };
    }
    aplicarBorda(row, 1, 3);
  }

  /* ---- Secção 4: Dados ------------------------------------------- */
  registos.forEach((reg, i) => {
    const r = primDados + i;
    const row = ws.getRow(r);

    row.getCell(1).value = reg.periodo || '';
    if (reg.dataReg) {
      row.getCell(2).value = excelSerial(reg.dataReg);
      row.getCell(2).numFmt = 'dd/mm/yyyy';
      row.getCell(2).alignment = { horizontal: 'center' };
    }
    row.getCell(3).value = reg.totalTaxas == null ? '' : reg.totalTaxas;
    row.getCell(3).numFmt = '#,##0';
    row.getCell(4).value = estado;
    row.getCell(5).value = /^\d+$/.test(String(reg.numeroDU)) && !/^0\d/.test(String(reg.numeroDU))
      ? Number(reg.numeroDU)
      : String(reg.numeroDU);
    row.getCell(5).alignment = { horizontal: 'center' };
    row.getCell(6).value = null; // Técnico [6] — em branco, a preencher
    row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_INPUT } };
    row.getCell(7).value = {
      formula: `IF($F${r}="","",IFERROR(VLOOKUP($F${r},${R.mapa},2,FALSE),""))`,
    };
    aplicarBorda(row, 1, 7);
  });

  /* ---- Acabamentos ----------------------------------------------- */
  // Lista pendente na coluna Técnico, alimentada pela área de configuração.
  ws.dataValidations.add(`F${primDados}:F${ultDados}`, {
    type: 'list',
    allowBlank: true,
    formulae: [R.mapaTec],
    showErrorMessage: false,
  });

  ws.autoFilter = {
    from: { row: linhaDadosCab, column: 1 },
    to: { row: ultDados, column: 7 },
  };
  ws.views = [{ state: 'frozen', ySplit: linhaDadosCab }];

  return {
    linhaResumoTitulo, linhaResumoCab, linhaTecTitulo, linhaTecCab,
    linhaTipoTitulo, linhaTipoCab, linhaDadosTitulo, linhaDadosCab,
    primDados, ultDados, primTec, ultTec, primTipo, ultTipo,
  };
}

/** Folha auxiliar com o resumo de cada PDF de origem. */
function construirFolhaOrigem(ws, ficheiros) {
  ws.columns = [
    { width: 44 }, { width: 14 }, { width: 14 }, { width: 18 }, { width: 18 },
  ];
  ws.mergeCells(1, 1, 1, 5);
  const t = ws.getCell(1, 1);
  t.value = 'FICHEIROS PDF DE ORIGEM';
  t.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_SECCAO } };
  t.alignment = { vertical: 'middle' };
  ws.getRow(1).height = 20;

  linhaCabecalho(ws, 2, ['Ficheiro', 'Período', 'Data', 'Nº de Processos', 'Total das Taxas']);

  ficheiros.forEach((f, i) => {
    const row = ws.getRow(3 + i);
    row.getCell(1).value = f.nome;
    row.getCell(2).value = f.periodo || '(não identificado)';
    row.getCell(3).value = f.dataIso ? formatDatePt(f.dataIso) : '(não identificada)';
    row.getCell(3).alignment = { horizontal: 'center' };
    row.getCell(4).value = f.registos.length;
    row.getCell(4).numFmt = '#,##0';
    row.getCell(5).value = f.registos.reduce((a, r) => a + (r.totalTaxas || 0), 0);
    row.getCell(5).numFmt = '#,##0';
    aplicarBorda(row, 1, 5);
  });

  const rTotal = 3 + ficheiros.length;
  const row = ws.getRow(rTotal);
  row.getCell(1).value = 'TOTAL';
  row.getCell(4).value = { formula: `SUM(D3:D${rTotal - 1})` };
  row.getCell(5).value = { formula: `SUM(E3:E${rTotal - 1})` };
  row.getCell(4).numFmt = '#,##0';
  row.getCell(5).numFmt = '#,##0';
  for (let c = 1; c <= 5; c++) row.getCell(c).font = { bold: true };
  aplicarBorda(row, 1, 5);
}

/**
 * Constrói o workbook de um grupo (3 dias seguidos) ou do consolidado.
 *
 * @param {object} ExcelJS         biblioteca ExcelJS
 * @param {object} grupo           { rotulo, ficheiros: [{nome, periodo, dataIso, registos}] }
 * @param {object} opcoes          { tecnicos, estado, linhasTecnicos, consolidado }
 * @returns {Promise<ArrayBuffer>}
 */
export async function criarWorkbook(ExcelJS, grupo, opcoes = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Gerador de Relatórios';
  wb.created = new Date();

  const registos = [];
  for (const f of grupo.ficheiros) {
    for (const r of f.registos) {
      registos.push({
        periodo: f.periodo,
        dataReg: r.dataReg,
        totalTaxas: r.totalTaxas,
        numeroDU: r.numeroDU,
      });
    }
  }
  registos.sort(
    (a, b) => (a.dataReg < b.dataReg ? -1 : a.dataReg > b.dataReg ? 1 : 0)
      || String(a.numeroDU).localeCompare(String(b.numeroDU), 'pt', { numeric: true })
  );

  const nomes = grupo.ficheiros.map((f) => f.nome).join(' · ');
  const ws = wb.addWorksheet('Relatório', {
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  construirFolhaRelatorio(ws, {
    titulo: opcoes.consolidado
      ? 'Relatório Consolidado — Todos os Períodos'
      : `Relatório ${grupo.rotulo}`,
    subtitulo: `${grupo.ficheiros.length} ficheiro(s) PDF: ${nomes}\nGerado em ${new Date().toLocaleString('pt-PT')}`,
    registos,
    tecnicos: opcoes.tecnicos || [],
    estado: opcoes.estado || 'Pago',
    linhasTecnicos: opcoes.linhasTecnicos || 12,
  });

  construirFolhaOrigem(wb.addWorksheet('Ficheiros de Origem'), grupo.ficheiros);

  return wb.xlsx.writeBuffer();
}
