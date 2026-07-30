/**
 * report.js — construção dos ficheiros Excel (.xlsx).
 *
 * A biblioteca ExcelJS é injectada como argumento para o módulo funcionar tanto
 * no browser (global `ExcelJS`) como em Node (import 'exceljs').
 *
 * Estrutura de cada folha "Relatório":
 *   1. Resumo compacto — duas linhas (nº de processos e total das taxas) com uma
 *      coluna por indicador: total, cada período, atribuição, estados e tipos.
 *      É a única parte congelada da folha.
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

/** Coluna onde começa a metade direita da tabela de técnicos. */
const COL_TEC_DIR = 6;   // F

/** Colunas reservadas no resumo para os tipos distintos. */
const N_COLUNAS_TIPO = 5;

/** Primeira coluna do mapa auxiliar escondido (nome, tipo, ordem). */
const COL_AUX = 20;      // T

const COR_TITULO = 'FF1F3864';
const COR_SECCAO = 'FF2E75B6';
const COR_CABECALHO = 'FFD9E2F3';
const COR_INPUT = 'FFFFF2CC';
const COR_CALC = 'FFF2F2F2';
const BORDA = { style: 'thin', color: { argb: 'FFBFBFBF' } };

/**
 * Formato dos números. O prefixo de idioma obriga a folha de cálculo a usar o
 * ponto como separador de milhares, seja quais forem as definições regionais de
 * quem abre o ficheiro. Usa-se 416 (português do Brasil) e não 816 (Portugal)
 * porque a convenção de Portugal é separar os milhares por espaço.
 */
const FMT_NUMERO = '[$-416]#,##0';
const FMT_DATA = '[$-416]dd/mm/yyyy';

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
 *   RESUMO — uma coluna por indicador, duas linhas de valores:     <- linhas 3..6
 *            "Nº de Processos" e "Total das Taxas"; inclui o total,
 *            cada período, a atribuição, os estados e os tipos.
 *   TÉCNICOS [6] — 6 linhas à esquerda + 6 à direita               <- a partir da 8
 *   DADOS — uma linha por processo                                 <- a partir da 19
 *
 * Só o resumo fica congelado (as suas duas linhas de valores mais o cabeçalho);
 * o resto acompanha a rolagem.
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

  const periodos = [...new Set(registos.map((r) => r.periodo).filter(Boolean))];
  // Estados possíveis: o valor por omissão mais os restantes da lista.
  const estados = [...new Set([estado, 'Pago', 'Não Pago'].filter(Boolean))];
  const nTipos = N_COLUNAS_TIPO;

  /* ---- Linhas ----------------------------------------------------- */
  const linhaResumoTitulo = 3;
  const linhaResumoCab = 4;
  const linhaResumoN = 5;      // Nº de Processos
  const linhaResumoT = 6;      // Total das Taxas
  const ultimaFixa = linhaResumoT;

  const linhaTecTitulo = ultimaFixa + 2;
  const linhaTecCab = linhaTecTitulo + 1;
  const primTec = linhaTecCab + 1;
  const ultTec = primTec + metade - 1;

  const linhaDadosTitulo = ultTec + 2;
  const linhaDadosCab = linhaDadosTitulo + 1;
  const primDados = linhaDadosCab + 1;
  const ultDados = primDados + Math.max(registos.length, 1) - 1;

  /* ---- Colunas ---------------------------------------------------- */
  // As 8 primeiras servem a tabela de dados; as seguintes só o resumo.
  const nColResumo = 1 + 1 + periodos.length + 2 + estados.length + nTipos + 1;
  const larguras = [24, 14, 18, 14, 14, 24, 16, 22];
  while (larguras.length < nColResumo) larguras.push(13);
  ws.columns = [
    ...larguras.map((width) => ({ width })),
    ...Array.from({ length: COL_AUX - larguras.length - 1 }, () => ({ width: 9 })),
    { width: 18 }, { width: 14 }, { width: 8 }, // área auxiliar
  ];
  for (const c of [COL_AUX, COL_AUX + 1, COL_AUX + 2]) ws.getColumn(c).hidden = true;

  /* ---- Título ----------------------------------------------------- */
  ws.mergeCells(1, 1, 1, nColResumo);
  const t = ws.getCell(1, 1);
  t.value = titulo;
  t.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_TITULO } };
  t.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(1).height = 24;

  ws.mergeCells(2, 1, 2, nColResumo);
  const s = ws.getCell(2, 1);
  s.value = subtitulo;
  s.font = { size: 10, color: { argb: 'FF404040' } };
  s.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  ws.getRow(2).height = 26;

  /* Intervalos usados pelas fórmulas -------------------------------- */
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
    mapa: `$${A}$1:$${B}$${nTec}`,
    mapaNomes: `$${A}$1:$${A}$${nTec}`,
    mapaTipos: `$${B}$1:$${B}$${nTec}`,
    mapaOrdem: `$${C}$1:$${C}$${nTec}`,
  };

  /* ---- Área auxiliar (escondida): junta os dois lados num só mapa -- */
  for (let i = 0; i < nTec; i++) {
    const r = 1 + i;
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
        : `IF($${B}${r}="","",IF(COUNTIF($${B}$1:$${B}${r},$${B}${r})=1,`
          + `MAX($${C}$1:$${C}${r - 1})+1,""))`,
    };
  }

  /* ---- Resumo compacto: uma coluna por indicador ------------------ */
  tituloSeccao(
    ws,
    linhaResumoTitulo,
    'RESUMO',
    'sempre à vista — acompanha o preenchimento da tabela de dados',
    nColResumo
  );

  const colunas = [];
  colunas.push({
    titulo: 'Total de processos',
    n: `COUNTA(${R.du})`,
    soma: `SUM(${R.total})`,
    grupo: 'total',
  });
  for (const p of periodos) {
    colunas.push({
      titulo: p,
      n: `COUNTIF(${R.periodo},"${p}")`,
      soma: `SUMIF(${R.periodo},"${p}",${R.total})`,
      grupo: 'periodo',
    });
  }
  // SUMPRODUCT em vez de SUMIF(...,"<>",...): o critério "<>" não é
  // interpretado da mesma maneira por todas as folhas de cálculo.
  const somaComTecnico = `SUMPRODUCT((${R.tecnico}<>"")*${R.total})`;
  colunas.push({
    titulo: 'Com técnico',
    n: `COUNTA(${R.du})-COUNTBLANK(${R.tecnico})`,
    soma: somaComTecnico,
    grupo: 'atribuicao',
  });
  colunas.push({
    titulo: 'Por atribuir',
    n: `COUNTBLANK(${R.tecnico})`,
    soma: `SUM(${R.total})-${somaComTecnico}`,
    grupo: 'atribuicao',
  });
  for (const e of estados) {
    colunas.push({ titulo: e, criterio: R.estado, grupo: 'estado' });
  }
  // Índice, dentro de `colunas`, da primeira coluna de tipo.
  const primColTipo = 1 + periodos.length + 2 + estados.length;
  for (let i = 0; i < nTipos; i++) {
    colunas.push({
      tituloFormula: `IFERROR(INDEX(${R.mapaTipos},MATCH(${i + 1},${R.mapaOrdem},0)),"")`,
      criterio: R.tipo,
      grupo: 'tipo',
    });
  }
  colunas.push({ titulo: 'Outros tipos', grupo: 'tipo', outros: true });

  const FUNDOS = {
    total: 'FFBDD7EE',
    periodo: COR_CABECALHO,
    atribuicao: 'FFEDEDED',
    estado: 'FFE2EFDA',
    tipo: 'FFFCE4D6',
  };

  ws.getCell(linhaResumoCab, 1).value = 'Indicador';
  ws.getCell(linhaResumoN, 1).value = 'Nº de Processos';
  ws.getCell(linhaResumoT, 1).value = 'Total das Taxas';
  for (const r of [linhaResumoCab, linhaResumoN, linhaResumoT]) {
    const cell = ws.getCell(r, 1);
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CABECALHO } };
    cell.alignment = { vertical: 'middle' };
  }
  ws.getRow(linhaResumoCab).height = 28;

  colunas.forEach((col, i) => {
    const c = 2 + i;
    const letra = letraCol(c);
    const primTipoLetra = letraCol(2 + primColTipo);
    const ultTipoLetra = letraCol(2 + primColTipo + nTipos - 1);

    const cabecalho = ws.getCell(linhaResumoCab, c);
    cabecalho.value = col.tituloFormula ? { formula: col.tituloFormula } : col.titulo;
    cabecalho.font = { bold: true, size: 10 };
    cabecalho.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: FUNDOS[col.grupo] } };
    cabecalho.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

    let fN;
    let fT;
    if (col.outros) {
      // Rede de segurança: tipos que não couberam nas colunas acima.
      fN = `SUMPRODUCT((${R.tipo}<>"")*1)`
        + `-SUM($${primTipoLetra}$${linhaResumoN}:$${ultTipoLetra}$${linhaResumoN})`;
      fT = `SUMPRODUCT((${R.tipo}<>"")*${R.total})`
        + `-SUM($${primTipoLetra}$${linhaResumoT}:$${ultTipoLetra}$${linhaResumoT})`;
    } else if (col.criterio) {
      fN = `IF(${letra}$${linhaResumoCab}="","",COUNTIF(${col.criterio},${letra}$${linhaResumoCab}))`;
      fT = `IF(${letra}$${linhaResumoCab}="","",`
        + `SUMIF(${col.criterio},${letra}$${linhaResumoCab},${R.total}))`;
    } else {
      fN = col.n;
      fT = col.soma;
    }

    [[linhaResumoN, fN], [linhaResumoT, fT]].forEach(([r, formula]) => {
      const cell = ws.getCell(r, c);
      cell.value = { formula };
      cell.numFmt = FMT_NUMERO;
      cell.alignment = { horizontal: 'right' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_CALC } };
      if (col.grupo === 'total') cell.font = { bold: true };
      if (col.outros) cell.font = { italic: true, color: { argb: 'FF595959' } };
    });
  });

  for (const r of [linhaResumoCab, linhaResumoN, linhaResumoT]) {
    aplicarBorda(ws.getRow(r), 1, 1 + colunas.length);
  }

  /* ---- Técnicos: metade das linhas de cada lado ------------------- */
  tituloSeccao(
    ws,
    linhaTecTitulo,
    'TÉCNICOS [6] — DEFINIR O TIPO [7] DE CADA UM',
    'escreva o nome e o tipo nas células amarelas; na tabela de dados basta escolher '
      + 'o técnico e o tipo aparece sozinho',
    nColResumo
  );
  linhaCabecalho(ws, linhaTecCab, ['Técnico [6]', 'Tipo [7]', 'Nº de Processos', 'Total das Taxas']);
  linhaCabecalho(
    ws, linhaTecCab, ['Técnico [6]', 'Tipo [7]', 'Nº de Processos', 'Total das Taxas'], COL_TEC_DIR
  );

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
        row.getCell(col + 2).numFmt = FMT_NUMERO;
        row.getCell(col + 3).numFmt = FMT_NUMERO;
        aplicarBorda(row, col, col + 3);
      });
  }

  /* ---- Dados ------------------------------------------------------ */
  tituloSeccao(
    ws,
    linhaDadosTitulo,
    'DADOS',
    'uma linha por processo extraído dos PDFs; indique o técnico responsável em cada '
      + 'linha — a última coluna mostra de que PDF veio',
    nColResumo
  );
  linhaCabecalho(ws, linhaDadosCab, [...COLUNAS, COLUNA_ORIGEM]);

  registos.forEach((reg, i) => {
    const r = primDados + i;
    const row = ws.getRow(r);

    row.getCell(1).value = reg.periodo || '';
    if (reg.dataReg) {
      row.getCell(2).value = excelSerial(reg.dataReg);
      row.getCell(2).numFmt = FMT_DATA;
      row.getCell(2).alignment = { horizontal: 'center' };
    }
    row.getCell(3).value = reg.totalTaxas == null ? null : reg.totalTaxas;
    row.getCell(3).numFmt = FMT_NUMERO;
    row.getCell(4).value = reg.estado || estado;
    row.getCell(4).alignment = { horizontal: 'center' };
    row.getCell(5).value = /^\d+$/.test(String(reg.numeroDU)) && !/^0\d/.test(String(reg.numeroDU))
      ? Number(reg.numeroDU)
      : String(reg.numeroDU);
    row.getCell(5).alignment = { horizontal: 'center' };
    // Técnico [6] — em branco, a preencher; ou já preenchido, quando a origem é
    // um relatório anterior.
    row.getCell(6).value = reg.tecnico || null;
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

  // Só o resumo fica fixo; os técnicos e os dados acompanham a rolagem.
  // A célula activa tem de ficar abaixo do corte: se ficasse em A1 — dentro da
  // zona congelada — alguns visualizadores rolam o painel de baixo até lá e o
  // resumo aparece duas vezes.
  ws.views = [{
    state: 'frozen',
    ySplit: ultimaFixa,
    topLeftCell: `A${ultimaFixa + 1}`,
    activeCell: `A${ultimaFixa + 1}`,
  }];

  return {
    linhaResumoTitulo, linhaResumoCab, linhaResumoN, linhaResumoT, ultimaFixa,
    linhaTecTitulo, linhaTecCab, linhaDadosTitulo, linhaDadosCab,
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
    row.getCell(4).numFmt = FMT_NUMERO;
    row.getCell(5).value = f.registos.reduce((a, r) => a + (r.totalTaxas || 0), 0);
    row.getCell(5).numFmt = FMT_NUMERO;
    aplicarBorda(row, 1, 5);
  });

  const rTotal = 3 + ficheiros.length;
  const row = ws.getRow(rTotal);
  row.getCell(1).value = 'TOTAL';
  row.getCell(4).value = { formula: `SUM(D3:D${rTotal - 1})` };
  row.getCell(5).value = { formula: `SUM(E3:E${rTotal - 1})` };
  row.getCell(4).numFmt = FMT_NUMERO;
  row.getCell(5).numFmt = FMT_NUMERO;
  for (let c = 1; c <= 5; c++) row.getCell(c).font = { bold: true };
  aplicarBorda(row, 1, 5);
}

/**
 * Folha com os processos que apareceram em mais do que um relatório, o que foi
 * mantido e o que foi posto de lado.
 */
function construirFolhaRedundancias(ws, duplicados, politica = '') {
  ws.columns = [
    { width: 12 }, { width: 13 }, { width: 13 }, { width: 26 }, { width: 22 },
    { width: 12 }, { width: 16 }, { width: 26 }, { width: 22 }, { width: 12 }, { width: 16 },
  ];

  ws.mergeCells(1, 1, 1, 11);
  const t = ws.getCell(1, 1);
  t.value = 'PROCESSOS REPETIDOS';
  t.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_SECCAO } };
  t.alignment = { vertical: 'middle' };
  ws.getRow(1).height = 20;

  ws.mergeCells(2, 1, 2, 11);
  const nota = ws.getCell(2, 1);
  nota.value = politica
    ? `Em caso de repetição foi mantida ${politica}.`
    : 'Processos encontrados em mais do que um relatório.';
  nota.font = { italic: true, size: 9, color: { argb: 'FF595959' } };

  ws.mergeCells(3, 4, 3, 7);
  ws.mergeCells(3, 8, 3, 11);
  for (const [col, texto, cor] of [[4, 'MANTIDO', 'FFE2EFDA'], [8, 'POSTO DE LADO', 'FFFCE4D6']]) {
    const c = ws.getCell(3, col);
    c.value = texto;
    c.font = { bold: true };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cor } };
    c.alignment = { horizontal: 'center' };
  }

  linhaCabecalho(ws, 4, [
    'Nº do DU', 'Data de Reg.', 'Situação',
    'Relatório', 'Técnico', 'Estado', 'Total das Taxas',
    'Relatório', 'Técnico', 'Estado', 'Total das Taxas',
  ]);

  duplicados.forEach((d, i) => {
    const row = ws.getRow(5 + i);
    row.getCell(1).value = /^\d+$/.test(String(d.numeroDU)) ? Number(d.numeroDU) : d.numeroDU;
    row.getCell(2).value = d.dataReg ? excelSerial(d.dataReg) : '';
    row.getCell(2).numFmt = FMT_DATA;
    row.getCell(3).value = d.iguais ? 'Dados iguais' : 'Dados diferentes';
    if (!d.iguais) row.getCell(3).font = { bold: true, color: { argb: 'FFB3261E' } };

    [[4, d.mantido], [8, d.descartado]].forEach(([col, versao]) => {
      row.getCell(col).value = versao.origemRelatorio || versao.ficheiro || '';
      row.getCell(col).font = { size: 9, color: { argb: 'FF595959' } };
      row.getCell(col + 1).value = versao.tecnico || '';
      row.getCell(col + 2).value = versao.estado || '';
      row.getCell(col + 3).value = versao.totalTaxas == null ? null : versao.totalTaxas;
      row.getCell(col + 3).numFmt = FMT_NUMERO;
    });
    aplicarBorda(row, 1, 11);
  });

  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + duplicados.length, column: 11 } };
  ws.views = [{ state: 'frozen', ySplit: 4, topLeftCell: 'A5', activeCell: 'A5' }];
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

  // Quando os registos já vêm juntos (junção de relatórios, com as
  // redundâncias resolvidas), usam-se tal como estão.
  const registos = opcoes.registos || grupo.ficheiros.flatMap((f) => f.registos.map((r) => ({
    periodo: r.periodo || f.periodo,
    dataReg: r.dataReg,
    totalTaxas: r.totalTaxas,
    numeroDU: r.numeroDU,
    estado: r.estado,
    tecnico: r.tecnico,
    ficheiro: r.ficheiro || f.nome,
  })));
  // Sem ordenação: as linhas ficam na ordem em que aparecem nos PDFs.

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

  if (opcoes.duplicados && opcoes.duplicados.length) {
    construirFolhaRedundancias(wb.addWorksheet('Redundâncias'), opcoes.duplicados, opcoes.politica);
  }

  return wb.xlsx.writeBuffer();
}
