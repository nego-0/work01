/**
 * report.js — construção dos ficheiros Excel (.xlsx).
 *
 * A biblioteca ExcelJS é injectada como argumento para o módulo funcionar tanto
 * no browser (global `ExcelJS`) como em Node (import 'exceljs').
 *
 * Estrutura de cada folha "Relatório":
 *   1. Resumo, com as estatísticas por Estado [4] e por Tipo [7] à direita
 *   2. Técnicos — área editável que define o Tipo [7] em função do Técnico [6]
 *      e mostra, por técnico, o nº de processos e o total das taxas, em duas
 *      metades lado a lado
 *   3. Dados: Período | Data de Reg. | Total das Taxas | Estado | Nº do DU |
 *             Técnico | Tipo | Ficheiro (PDF)
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

/** Coluna extra, depois das 7 pedidas: de que PDF veio cada linha. */
export const COLUNA_ORIGEM = 'Ficheiro (PDF)';

/** Nº total de colunas da tabela de dados. */
const N_COLUNAS = COLUNAS.length + 1;

/* Colunas onde começa cada bloco de estatísticas / da metade direita dos
   técnicos, e a área auxiliar escondida que junta os dois lados do mapa. */
const COL_ESTADO = 5;    // E — estatísticas por Estado [4]
const COL_TIPO = 9;      // I — estatísticas por Tipo [7]
const COL_TEC_DIR = 6;   // F — metade direita da tabela de técnicos
const COL_AUX = 13;      // M — mapa auxiliar (escondido)

const COR_TITULO = 'FF1F3864';
const COR_SECCAO = 'FF2E75B6';
const COR_CABECALHO = 'FFD9E2F3';
const COR_INPUT = 'FFFFF2CC';
const COR_CALC = 'FFF2F2F2';
const BORDA = { style: 'thin', color: { argb: 'FFBFBFBF' } };

/** Letra da coluna (1 -> 'A'). Só é usado até à coluna Z. */
function letraCol(n) {
  return String.fromCharCode(64 + n);
}

/** Serial do Excel para uma data ISO (independente do fuso horário). */
function excelSerial(iso) {
  return dayNumber(iso) + 25569;
}

function aplicarBorda(row, deCol, ateCol) {
  for (let c = deCol; c <= ateCol; c++) {
    row.getCell(c).border = { top: BORDA, left: BORDA, bottom: BORDA, right: BORDA };
  }
}

/**
 * Faixa de secção. A nota, quando existe, fica na mesma linha do título para
 * não gastar altura — o topo da folha tem de caber no ecrã.
 */
function tituloSeccao(ws, linha, texto, nota, ateCol = N_COLUNAS) {
  const cell = ws.getCell(linha, 1);
  const negrito = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  cell.value = nota
    ? { richText: [
        { font: negrito, text: texto },
        { font: { size: 10, italic: true, color: { argb: 'FFD6E4F5' } }, text: `   ${nota}` },
      ] }
    : texto;
  cell.font = negrito;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_SECCAO } };
  cell.alignment = { vertical: 'middle' };
  ws.mergeCells(linha, 1, linha, ateCol);
  ws.getRow(linha).height = 20;
  return linha + 1;
}

function linhaCabecalho(ws, linha, valores, colInicial = 1) {
  const row = ws.getRow(linha);
  valores.forEach((v, i) => {
    const cell = row.getCell(colInicial + i);
    cell.value = v;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CABECALHO } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  aplicarBorda(row, colInicial, colInicial + valores.length - 1);
  return linha + 1;
}

/**
 * Cria a folha "Relatório" com estatísticas + dados.
 *
 * Disposição:
 *
 *   RESUMO        | ESTADOS [4]      | TIPOS [7]        <- linhas 4..12
 *   TÉCNICOS [6] (6 linhas à esquerda + 6 à direita)    <- linhas 14..21
 *   DADOS (uma linha por processo)                      <- a partir da linha 25
 *
 * Tudo o que está acima da tabela de dados, incluindo o cabeçalho das colunas,
 * fica congelado: as estatísticas continuam à vista enquanto se percorre a
 * tabela.
 *
 * @param {object} ws            worksheet ExcelJS
 * @param {object} params
 * @param {string} params.titulo
 * @param {string} params.subtitulo
 * @param {Array}  params.registos  [{periodo, dataReg, totalTaxas, numeroDU}]
 * @param {Array}  params.tecnicos  [{nome, tipo}] pré-preenchidos (pode ser [])
 * @param {string} params.estado    valor por omissão do campo Estado [4]
 * @param {number} params.linhasTecnicos nº total de linhas de técnicos (metade
 *                                       de cada lado)
 */
function construirFolhaRelatorio(ws, params) {
  const { titulo, subtitulo, registos, tecnicos = [], estado = 'Pago' } = params;

  // Técnicos: metade das linhas de cada lado.
  const totalTec = Math.max(params.linhasTecnicos || 12, tecnicos.length, 2);
  const metade = Math.ceil(totalTec / 2);
  const nTec = metade * 2;

  ws.columns = [
    { width: 22 }, { width: 14 }, { width: 18 }, { width: 14 }, // A-D
    { width: 14 }, { width: 22 }, { width: 16 }, { width: 22 }, // E-H
    { width: 16 }, { width: 16 }, { width: 18 },                // I-K (tipos)
    { width: 2 },                                               // L  (separador)
    { width: 18 }, { width: 14 }, { width: 8 },                 // M-O (auxiliar)
  ];
  for (const c of [COL_AUX, COL_AUX + 1, COL_AUX + 2]) ws.getColumn(c).hidden = true;

  /* ---- Título ---------------------------------------------------- */
  ws.mergeCells(1, 1, 1, COL_TIPO + 2);
  const t = ws.getCell(1, 1);
  t.value = titulo;
  t.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_TITULO } };
  t.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 26;

  ws.mergeCells(2, 1, 2, COL_TIPO + 2);
  const s = ws.getCell(2, 1);
  s.value = subtitulo;
  s.font = { size: 10, color: { argb: 'FF404040' } };
  s.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  ws.getRow(2).height = 30;

  /* ---- Posições --------------------------------------------------- */
  const periodos = [...new Set(registos.map((r) => r.periodo).filter(Boolean))];
  // Estados possíveis: o valor por omissão mais os restantes da lista.
  const estados = [...new Set([estado, 'Pago', 'Não Pago'].filter(Boolean))];
  const nResumoLinhas = 2 + periodos.length + (periodos.length ? 1 : 0);
  const nTipos = Math.max(nResumoLinhas - 1, 4); // + a linha "Outros tipos"

  let linha = 4;
  const linhaResumoTitulo = linha;
  linha = tituloSeccao(ws, linha, 'RESUMO', null, COL_TIPO + 2);
  const linhaResumoCab = linha;
  linhaCabecalho(ws, linha, ['Indicador', 'Nº de Processos', 'Total das Taxas']);
  linhaCabecalho(ws, linha, ['Estado [4]', 'Nº de Processos', 'Total das Taxas'], COL_ESTADO);
  linha = linhaCabecalho(ws, linha, ['Tipo [7]', 'Nº de Processos', 'Total das Taxas'], COL_TIPO);
  const primResumo = linha;
  linha += Math.max(nResumoLinhas, estados.length, nTipos + 1);
  linha += 1; // espaço

  const linhaTecTitulo = linha;
  linha = tituloSeccao(
    ws,
    linha,
    'TÉCNICOS [6] — DEFINIR O TIPO [7] DE CADA UM',
    'escreva o nome e o tipo nas células amarelas; na tabela de dados basta escolher '
      + 'o técnico e o tipo aparece sozinho',
    COL_TIPO + 2
  );
  const linhaTecCab = linha;
  linhaCabecalho(ws, linha, ['Técnico [6]', 'Tipo [7]', 'Nº de Processos', 'Total das Taxas']);
  linha = linhaCabecalho(
    ws, linha, ['Técnico [6]', 'Tipo [7]', 'Nº de Processos', 'Total das Taxas'], COL_TEC_DIR
  );
  const primTec = linha;
  const ultTec = primTec + metade - 1;
  linha = ultTec + 2;

  const linhaDadosTitulo = linha;
  linha = tituloSeccao(
    ws,
    linha,
    'DADOS',
    'uma linha por processo extraído dos PDFs; indique o técnico responsável em cada '
      + 'linha — a última coluna mostra de que PDF veio',
    COL_TIPO + 2
  );
  const linhaDadosCab = linha;
  linha = linhaCabecalho(ws, linha, [...COLUNAS, COLUNA_ORIGEM]);
  const primDados = linha;
  const ultDados = primDados + Math.max(registos.length, 1) - 1;

  /* Intervalos usados pelas fórmulas ------------------------------- */
  const auxPrim = 1;
  const auxUlt = nTec;
  const A = letraCol(COL_AUX);          // nomes (auxiliar)
  const B = letraCol(COL_AUX + 1);      // tipos (auxiliar)
  const C = letraCol(COL_AUX + 2);      // ordem do 1º aparecimento

  const R = {
    periodo: `$A$${primDados}:$A$${ultDados}`,
    total: `$C$${primDados}:$C$${ultDados}`,
    estado: `$D$${primDados}:$D$${ultDados}`,
    du: `$E$${primDados}:$E$${ultDados}`,
    tecnico: `$F$${primDados}:$F$${ultDados}`,
    tipo: `$G$${primDados}:$G$${ultDados}`,
    mapa: `$${A}$${auxPrim}:$${B}$${auxUlt}`,
    mapaNomes: `$${A}$${auxPrim}:$${A}$${auxUlt}`,
    mapaTipos: `$${B}$${auxPrim}:$${B}$${auxUlt}`,
    mapaOrdem: `$${C}$${auxPrim}:$${C}$${auxUlt}`,
  };

  /* ---- Área auxiliar (escondida): junta os dois lados num só mapa - */
  for (let i = 0; i < nTec; i++) {
    const r = auxPrim + i;
    const rMapa = primTec + (i % metade);
    const colNome = i < metade ? 'A' : letraCol(COL_TEC_DIR);
    const colTipo = i < metade ? 'B' : letraCol(COL_TEC_DIR + 1);
    const row = ws.getRow(r);
    row.getCell(COL_AUX).value = { formula: `IF($${colNome}${rMapa}="","",$${colNome}${rMapa})` };
    row.getCell(COL_AUX + 1).value = { formula: `IF($${colTipo}${rMapa}="","",$${colTipo}${rMapa})` };
    // Numera cada tipo na 1ª vez que aparece, para a lista de tipos distintos.
    row.getCell(COL_AUX + 2).value = {
      formula: i === 0
        ? `IF($${B}${r}="","",1)`
        : `IF($${B}${r}="","",IF(COUNTIF($${B}$${auxPrim}:$${B}${r},$${B}${r})=1,`
          + `MAX($${C}$${auxPrim}:$${C}${r - 1})+1,""))`,
    };
  }

  /* ---- Resumo ----------------------------------------------------- */
  const resumo = [
    {
      rotulo: 'Total de processos',
      n: `COUNTA(${R.du})`,
      soma: `SUM(${R.total})`,
      destaque: true,
    },
    ...periodos.map((p) => ({
      rotulo: p,
      n: `COUNTIF(${R.periodo},"${p}")`,
      soma: `SUMIF(${R.periodo},"${p}",${R.total})`,
    })),
  ];
  if (periodos.length) {
    resumo.push({
      rotulo: 'Com técnico atribuído',
      n: `COUNTA(${R.du})-COUNTBLANK(${R.tecnico})`,
      soma: `SUMIF(${R.tecnico},"<>",${R.total})`,
    });
  }
  resumo.push({
    rotulo: 'Por atribuir',
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
    if (item.destaque) for (let c = 1; c <= 3; c++) row.getCell(c).font = { bold: true };
    for (let c = 1; c <= 3; c++) {
      row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CALC } };
    }
    aplicarBorda(row, 1, 3);
  });

  /* ---- Estatísticas por Estado ------------------------------------ */
  const letraEstado = letraCol(COL_ESTADO);
  estados.forEach((nome, i) => {
    const r = primResumo + i;
    const row = ws.getRow(r);
    row.getCell(COL_ESTADO).value = nome;
    row.getCell(COL_ESTADO + 1).value = { formula: `COUNTIF(${R.estado},$${letraEstado}${r})` };
    row.getCell(COL_ESTADO + 2).value = {
      formula: `SUMIF(${R.estado},$${letraEstado}${r},${R.total})`,
    };
    row.getCell(COL_ESTADO + 1).numFmt = '#,##0';
    row.getCell(COL_ESTADO + 2).numFmt = '#,##0';
    for (let c = COL_ESTADO; c <= COL_ESTADO + 2; c++) {
      row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CALC } };
    }
    aplicarBorda(row, COL_ESTADO, COL_ESTADO + 2);
  });

  /* ---- Estatísticas por Tipo -------------------------------------- */
  const letraTipo = letraCol(COL_TIPO);
  for (let i = 0; i < nTipos; i++) {
    const r = primResumo + i;
    const row = ws.getRow(r);
    // i-ésimo tipo distinto definido na tabela de técnicos.
    row.getCell(COL_TIPO).value = {
      formula: `IFERROR(INDEX(${R.mapaTipos},MATCH(${i + 1},${R.mapaOrdem},0)),"")`,
    };
    row.getCell(COL_TIPO + 1).value = {
      formula: `IF($${letraTipo}${r}="","",COUNTIF(${R.tipo},$${letraTipo}${r}))`,
    };
    row.getCell(COL_TIPO + 2).value = {
      formula: `IF($${letraTipo}${r}="","",SUMIF(${R.tipo},$${letraTipo}${r},${R.total}))`,
    };
    row.getCell(COL_TIPO + 1).numFmt = '#,##0';
    row.getCell(COL_TIPO + 2).numFmt = '#,##0';
    for (let c = COL_TIPO; c <= COL_TIPO + 2; c++) {
      row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CALC } };
    }
    aplicarBorda(row, COL_TIPO, COL_TIPO + 2);
  }

  // Rede de segurança: o que sobra caso haja mais tipos do que linhas.
  const rOutros = primResumo + nTipos;
  const rowOutros = ws.getRow(rOutros);
  const letraN = letraCol(COL_TIPO + 1);
  const letraT = letraCol(COL_TIPO + 2);
  rowOutros.getCell(COL_TIPO).value = 'Outros tipos';
  rowOutros.getCell(COL_TIPO + 1).value = {
    formula: `COUNTIF(${R.tipo},"?*")-SUM($${letraN}$${primResumo}:$${letraN}$${rOutros - 1})`,
  };
  rowOutros.getCell(COL_TIPO + 2).value = {
    formula: `SUMIF(${R.tipo},"?*",${R.total})`
      + `-SUM($${letraT}$${primResumo}:$${letraT}$${rOutros - 1})`,
  };
  rowOutros.getCell(COL_TIPO + 1).numFmt = '#,##0';
  rowOutros.getCell(COL_TIPO + 2).numFmt = '#,##0';
  for (let c = COL_TIPO; c <= COL_TIPO + 2; c++) {
    rowOutros.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CALC } };
    rowOutros.getCell(c).font = { italic: true, color: { argb: 'FF595959' } };
  }
  aplicarBorda(rowOutros, COL_TIPO, COL_TIPO + 2);

  /* ---- Técnicos: 6 linhas de cada lado ---------------------------- */
  for (let i = 0; i < metade; i++) {
    const r = primTec + i;
    const row = ws.getRow(r);

    [{ col: 1, cfg: tecnicos[i] }, { col: COL_TEC_DIR, cfg: tecnicos[metade + i] }]
      .forEach(({ col, cfg }) => {
        const letra = letraCol(col);
        row.getCell(col).value = cfg ? cfg.nome : null;
        row.getCell(col + 1).value = cfg ? cfg.tipo : null;
        for (const c of [col, col + 1]) {
          row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_INPUT } };
        }
        row.getCell(col + 2).value = {
          formula: `IF($${letra}${r}="","",COUNTIF(${R.tecnico},$${letra}${r}))`,
        };
        row.getCell(col + 3).value = {
          formula: `IF($${letra}${r}="","",SUMIF(${R.tecnico},$${letra}${r},${R.total}))`,
        };
        row.getCell(col + 2).numFmt = '#,##0';
        row.getCell(col + 3).numFmt = '#,##0';
        aplicarBorda(row, col, col + 3);
      });
  }

  /* ---- Dados ------------------------------------------------------ */
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
    row.getCell(4).alignment = { horizontal: 'center' };
    row.getCell(5).value = /^\d+$/.test(String(reg.numeroDU)) && !/^0\d/.test(String(reg.numeroDU))
      ? Number(reg.numeroDU)
      : String(reg.numeroDU);
    row.getCell(5).alignment = { horizontal: 'center' };
    row.getCell(6).value = null; // Técnico [6] — em branco, a preencher
    row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_INPUT } };
    row.getCell(7).value = {
      formula: `IF($F${r}="","",IFERROR(VLOOKUP($F${r},${R.mapa},2,FALSE),""))`,
    };
    row.getCell(8).value = reg.ficheiro || '';
    row.getCell(8).font = { size: 9, color: { argb: 'FF595959' } };
    aplicarBorda(row, 1, N_COLUNAS);
  });

  /* ---- Acabamentos ----------------------------------------------- */
  // Lista pendente na coluna Técnico, com os nomes dos dois lados.
  ws.dataValidations.add(`F${primDados}:F${ultDados}`, {
    type: 'list',
    allowBlank: true,
    formulae: [R.mapaNomes],
    showErrorMessage: false,
  });

  // Lista pendente na coluna Estado, para marcar um processo como não pago.
  ws.dataValidations.add(`D${primDados}:D${ultDados}`, {
    type: 'list',
    allowBlank: true,
    formulae: [`"${estados.join(',')}"`],
    showErrorMessage: false,
  });

  ws.autoFilter = {
    from: { row: linhaDadosCab, column: 1 },
    to: { row: ultDados, column: N_COLUNAS },
  };

  // Tudo o que está acima dos dados fica fixo, incluindo o cabeçalho das
  // colunas: as estatísticas acompanham a leitura da tabela.
  // O zoom a 90% dá folga para o topo congelado caber mesmo em ecrãs baixos.
  ws.views = [{
    state: 'frozen',
    ySplit: linhaDadosCab,
    topLeftCell: `A${primDados}`,
    zoomScale: 90,
    zoomScaleNormal: 90,
  }];

  return {
    linhaResumoTitulo, linhaResumoCab, linhaTecTitulo, linhaTecCab,
    linhaDadosTitulo, linhaDadosCab,
    primDados, ultDados, primTec, ultTec, metade,
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
        ficheiro: f.nome,
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
