/**
 * parser.js — extracção de dados dos PDFs "Document List" (ASYCUDAWorld).
 *
 * O módulo é puro (sem DOM) para poder ser usado tanto no browser como em Node
 * (ver tools/extract-cli.mjs). A única dependência externa é o objecto
 * `textContent` devolvido pelo pdf.js.
 */

/* ------------------------------------------------------------------ */
/* Utilitários                                                         */
/* ------------------------------------------------------------------ */

/** Remove acentos e passa a minúsculas. */
export function normalize(str) {
  return String(str)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

/**
 * Converte o texto de um montante em número.
 * Aceita "86642", "86.642,00", "86,642.00", "86 642,00".
 */
export function parseAmount(text) {
  if (text == null) return null;
  let s = String(text).replace(/[^\d.,-]/g, '').trim();
  if (!s) return null;

  const hasDot = s.includes('.');
  const hasComma = s.includes(',');

  if (hasDot && hasComma) {
    // O separador decimal é o que aparece mais à direita.
    const decSep = s.lastIndexOf(',') > s.lastIndexOf('.') ? ',' : '.';
    const thoSep = decSep === ',' ? '.' : ',';
    s = s.split(thoSep).join('');
    s = s.replace(decSep, '.');
  } else if (hasComma) {
    // "1,234" com grupos de 3 é milhar; caso contrário é decimal.
    s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.split(',').join('') : s.replace(',', '.');
  } else if (hasDot) {
    s = /^-?\d{1,3}(\.\d{3})+$/.test(s) ? s.split('.').join('') : s;
  }

  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

/** "2026-07-01" -> {iso, dia, mes, ano}; devolve null se não for data. */
export function parseIsoDate(text) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(text || '').trim());
  if (!m) return null;
  return { iso: m[0], ano: +m[1], mes: +m[2], dia: +m[3] };
}

/** Nº de dias desde a epoch (para comparar dias consecutivos). */
export function dayNumber(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return Math.floor(Date.UTC(y, m - 1, d) / 86400000);
}

/** "2026-07-01" -> "01.07.2026" */
export function formatDatePt(iso) {
  const [y, m, d] = iso.split('-');
  return `${d}.${m}.${y}`;
}

/* ------------------------------------------------------------------ */
/* [1] Período e data a partir do nome do ficheiro                      */
/* ------------------------------------------------------------------ */

const PERIODOS = [
  { canonico: 'Madrugada', re: /madrugada/ },
  { canonico: 'Manhã', re: /manh(a|ã|_|\b)/ },
  { canonico: 'Tarde', re: /tarde/ },
  { canonico: 'Noite', re: /noite/ },
];

/**
 * Extrai o período [1] e a data do nome do ficheiro.
 * Ex.: "Tarde 22.07.2026.pdf", "01.07.2026_Manhã.pdf", "2026-07-03 Madrugada.pdf"
 */
export function parseFileName(fileName) {
  const base = String(fileName).replace(/\.pdf$/i, '');
  const norm = normalize(base);

  let periodo = '';
  for (const p of PERIODOS) {
    if (p.re.test(norm)) { periodo = p.canonico; break; }
  }

  let dataIso = '';
  let m = /(\d{4})[.\-_/](\d{1,2})[.\-_/](\d{1,2})/.exec(base);
  if (m) {
    dataIso = `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  } else {
    m = /(\d{1,2})[.\-_/](\d{1,2})[.\-_/](\d{4})/.exec(base);
    if (m) {
      dataIso = `${m[3]}-${String(m[2]).padStart(2, '0')}-${String(m[1]).padStart(2, '0')}`;
    }
  }

  return { periodo, dataIso };
}

/* ------------------------------------------------------------------ */
/* Normalização geométrica do texto do PDF                              */
/* ------------------------------------------------------------------ */

/**
 * Converte os itens de `page.getTextContent()` em itens com coordenadas
 * "visuais" (x da esquerda para a direita, top de cima para baixo), mesmo
 * quando o texto está desenhado rodado 90º — o caso destes relatórios.
 */
export function itemsFromTextContent(textContent) {
  const out = [];
  for (const it of textContent.items) {
    if (!it.str || !it.str.trim()) continue;
    const [a, b, , , e, f] = it.transform;
    const theta = Math.atan2(b, a);
    const cos = Math.cos(theta);
    const sin = Math.sin(theta);
    const x = e * cos + f * sin;      // avanço no sentido da leitura
    const top = -(-e * sin + f * cos); // profundidade da página (cresce para baixo)
    out.push({ text: it.str.trim(), x, top, width: it.width || 0, textoOriginal: it.str });
  }
  return out;
}

/**
 * Divide cada item em palavras, estimando a posição de cada uma pela largura
 * média do caractere.
 *
 * É necessário porque um único item de texto pode atravessar duas colunas
 * quando o conteúdo de uma delas transborda — por exemplo, o nome do
 * destinatário colado ao total das taxas. Sem esta divisão, o total ficaria
 * preso ao nome e seria perdido.
 */
export function splitItems(items) {
  const out = [];
  for (const it of items) {
    const original = it.textoOriginal ?? it.text;
    const palavras = [...original.matchAll(/\S+/g)];
    if (palavras.length > 1 && it.width > 0) {
      const larguraChar = it.width / original.length;
      for (const p of palavras) {
        out.push({
          text: p[0],
          x: it.x + larguraChar * p.index,
          top: it.top,
          width: larguraChar * p[0].length,
        });
      }
    } else {
      out.push({ text: it.text, x: it.x, top: it.top, width: it.width });
    }
  }
  return out;
}

/** Agrupa itens em linhas pela coordenada `top`. */
export function groupIntoLines(items, tolerance = 3) {
  const lines = [];
  for (const it of [...items].sort((p, q) => p.top - q.top || p.x - q.x)) {
    const last = lines[lines.length - 1];
    if (last && Math.abs(last.top - it.top) <= tolerance) {
      last.items.push(it);
      last.top = (last.top * (last.items.length - 1) + it.top) / last.items.length;
    } else {
      lines.push({ top: it.top, items: [it] });
    }
  }
  for (const l of lines) l.items.sort((p, q) => p.x - q.x);
  return lines;
}

/* ------------------------------------------------------------------ */
/* Leitura da tabela                                                    */
/* ------------------------------------------------------------------ */

const HEADER_DU = /^n\s*[°ºo]?\s*do\s*du$/;
const HEADER_DATA_REG = /^data\s*de\s*reg/;
const HEADER_TOTAL = /^total\s*das/;

/**
 * Localiza a linha de cabeçalho da tabela e devolve as âncoras (x) de cada
 * coluna. O cabeçalho repete-se em todas as páginas e as posições variam
 * ligeiramente de ficheiro para ficheiro, por isso é detectado dinamicamente.
 */
function findHeader(lines) {
  for (let i = 0; i < lines.length; i++) {
    const texts = lines[i].items.map((it) => normalize(it.text));
    const iDu = texts.findIndex((t) => HEADER_DU.test(t));
    const iData = texts.findIndex((t) => HEADER_DATA_REG.test(t));
    const iTotal = texts.findIndex((t) => HEADER_TOTAL.test(t));
    if (iDu >= 0 && iData >= 0 && iTotal >= 0) {
      const anchors = lines[i].items.map((it) => it.x);
      return { lineIndex: i, top: lines[i].top, anchors, iDu, iData, iTotal };
    }
  }
  return null;
}

/** Texto da coluna `idx` presente numa linha. */
function cellText(line, anchors, idx, pad = 4) {
  const from = anchors[idx] - pad;
  const to = idx + 1 < anchors.length ? anchors[idx + 1] - pad : Infinity;
  return line.items
    .filter((it) => it.x >= from && it.x < to)
    .map((it) => it.text)
    .join(' ')
    .trim();
}

/**
 * Extrai as linhas de dados de uma página.
 * Devolve [{ numeroDU, dataReg, totalTaxas, totalTaxasTexto }].
 */
export function extractRowsFromPage(items) {
  // O cabeçalho é procurado nos itens originais ("N° do DU", "Data de Reg.",
  // "Total das"); os dados são lidos com os itens divididos em palavras.
  const header = findHeader(groupIntoLines(items));
  if (!header) return [];

  const { anchors, iDu, iData, iTotal } = header;
  const rows = [];

  for (const line of groupIntoLines(splitItems(items))) {
    if (line.top <= header.top + 1) continue;  // título, filtros e cabeçalho

    const dataTxt = cellText(line, anchors, iData);
    const data = parseIsoDate(dataTxt);
    if (!data) continue;                       // linhas de continuação / rodapé

    const duTxt = cellText(line, anchors, iDu);
    if (!/^[\w#/-]+$/.test(duTxt)) continue;   // sem Nº do DU => não é registo

    const totalTxt = cellText(line, anchors, iTotal);
    // O total é a última palavra numérica da coluna: se o nome do destinatário
    // transbordou, as palavras do nome ficam antes do valor.
    const numerico = totalTxt.split(/\s+/).filter((t) => /^[\d.,]+$/.test(t)).pop();
    const totalTaxas = parseAmount(numerico);

    // Assinala uma leitura duvidosa, para a interface poder avisar e deixar
    // corrigir à mão. Não deita a linha fora — mostra-a com o problema.
    const problema = totalTaxas == null || totalTaxas <= 0
      ? 'Total das Taxas não reconhecido'
      : null;

    rows.push({
      numeroDU: duTxt,
      dataReg: data.iso,
      totalTaxas,
      totalTaxasTexto: totalTxt,
      ...(problema ? { problema } : {}),
    });
  }

  return rows;
}

/**
 * Percorre todas as páginas de um documento pdf.js e devolve os registos.
 * `pdfDoc` é o resultado de `pdfjsLib.getDocument(...).promise`.
 */
export async function extractRowsFromDocument(pdfDoc, onProgress) {
  const rows = [];
  for (let p = 1; p <= pdfDoc.numPages; p++) {
    const page = await pdfDoc.getPage(p);
    const textContent = await page.getTextContent();
    rows.push(...extractRowsFromPage(itemsFromTextContent(textContent)));
    page.cleanup();
    if (onProgress) onProgress(p, pdfDoc.numPages);
  }
  return rows;
}

/* ------------------------------------------------------------------ */
/* Agrupamento em blocos de 3 dias seguidos                             */
/* ------------------------------------------------------------------ */

/**
 * Agrupa os ficheiros em blocos de até 3 dias consecutivos.
 * Ficheiros do mesmo dia ficam sempre no mesmo bloco.
 *
 * @param {Array<{dataIso:string}>} ficheiros
 * @returns {Array<{datas:string[], ficheiros:object[]}>}
 */
export function agruparPorDiasConsecutivos(ficheiros, maxDias = 3) {
  const porData = new Map();
  for (const f of ficheiros) {
    const chave = f.dataIso || 'sem-data';
    if (!porData.has(chave)) porData.set(chave, []);
    porData.get(chave).push(f);
  }

  const datas = [...porData.keys()].filter((d) => d !== 'sem-data').sort();
  const grupos = [];
  let actual = null;

  for (const data of datas) {
    const consecutiva =
      actual && dayNumber(data) === dayNumber(actual.datas[actual.datas.length - 1]) + 1;
    if (!actual || !consecutiva || actual.datas.length >= maxDias) {
      actual = { datas: [], ficheiros: [] };
      grupos.push(actual);
    }
    actual.datas.push(data);
    actual.ficheiros.push(...porData.get(data));
  }

  if (porData.has('sem-data')) {
    grupos.push({ datas: [], ficheiros: porData.get('sem-data') });
  }

  return grupos;
}

/** Nome do ficheiro Excel de um grupo. */
export function nomeDoGrupo(grupo) {
  if (!grupo.datas.length) return 'Relatorio_sem_data';
  const ini = formatDatePt(grupo.datas[0]).replace(/\./g, '-');
  if (grupo.datas.length === 1) return `Relatorio_${ini}`;
  const fim = formatDatePt(grupo.datas[grupo.datas.length - 1]).replace(/\./g, '-');
  return `Relatorio_${ini}_a_${fim}`;
}

/** Rótulo legível de um grupo. */
export function rotuloDoGrupo(grupo) {
  if (!grupo.datas.length) return 'Sem data no nome do ficheiro';
  if (grupo.datas.length === 1) return formatDatePt(grupo.datas[0]);
  return `${formatDatePt(grupo.datas[0])} a ${formatDatePt(grupo.datas[grupo.datas.length - 1])}`;
}
