/**
 * leitor-relatorio.js — lê um relatório Excel já gerado (e já preenchido com os
 * técnicos) para se poder consolidar vários deles sem voltar aos PDFs.
 *
 * Devolve os registos com o Técnico [6] e o Estado [4] tal como estão no
 * ficheiro, mais o mapa Técnico → Tipo da área de configuração. O Tipo não é
 * lido da fórmula: é recalculado a partir desse mapa, para não depender de o
 * Excel ter deixado o valor em cache.
 *
 * A biblioteca ExcelJS é injectada, tal como em report.js.
 */

import { COLUNAS } from './report.js';

/** Converte o que estiver na célula da data para ISO (yyyy-mm-dd). */
function dataParaIso(valor) {
  if (valor instanceof Date) return valor.toISOString().slice(0, 10);
  if (typeof valor === 'number') {
    const ms = (valor - 25569) * 86400000;
    return new Date(ms).toISOString().slice(0, 10);
  }
  const texto = String(valor ?? '').trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(texto);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})/.exec(texto);
  if (m) return `${m[3]}-${String(m[2]).padStart(2, '0')}-${String(m[1]).padStart(2, '0')}`;
  return '';
}

/** Texto de uma célula, seja ela valor simples, fórmula ou texto formatado. */
function texto(celula) {
  const v = celula ? celula.value : null;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((p) => p.text).join('').trim();
    if ('result' in v) return v.result === null || v.result === undefined ? '' : String(v.result).trim();
    if ('text' in v) return String(v.text).trim();
    return '';
  }
  return String(v).trim();
}

function numero(celula) {
  const v = celula ? celula.value : null;
  if (typeof v === 'number') return v;
  if (v && typeof v === 'object' && typeof v.result === 'number') return v.result;
  const n = Number(String(v ?? '').replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}\b)/g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

/**
 * Lê um relatório gerado por esta aplicação.
 *
 * @param {object} ExcelJS
 * @param {ArrayBuffer} buffer
 * @param {string} nomeFicheiro
 * @returns {Promise<{nome, tecnicos, registos}>}
 */
export async function lerRelatorio(ExcelJS, buffer, nomeFicheiro = '') {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const ws = wb.getWorksheet('Relatório') || wb.worksheets[0];
  if (!ws) throw new Error('O ficheiro não tem nenhuma folha.');

  /* ---- Cabeçalho da tabela de dados ------------------------------- */
  let linhaCab = 0;
  for (let r = 1; r <= ws.rowCount; r++) {
    if (texto(ws.getCell(r, 1)) === COLUNAS[0] && texto(ws.getCell(r, 2)) === COLUNAS[1]) {
      linhaCab = r;
      break;
    }
  }
  if (!linhaCab) {
    throw new Error('Não parece um relatório desta aplicação (falta a tabela de dados).');
  }

  /* ---- Mapa Técnico → Tipo ---------------------------------------- */
  const tecnicos = [];
  const vistos = new Set();
  let linhaTec = 0;
  for (let r = 1; r < linhaCab; r++) {
    if (texto(ws.getCell(r, 1)) === 'Técnico [6]') { linhaTec = r; break; }
  }
  if (linhaTec) {
    for (let r = linhaTec + 1; r < linhaCab; r++) {
      // As faixas de secção ocupam a largura toda; a tabela de técnicos acaba aí.
      if (ws.getCell(r, 1).isMerged) break;
      // As duas metades da tabela: à esquerda (A,B) e à direita (F,G).
      for (const [colNome, colTipo] of [[1, 2], [6, 7]]) {
        const nome = texto(ws.getCell(r, colNome));
        if (!nome || nome === 'Técnico [6]' || vistos.has(nome.toLowerCase())) continue;
        vistos.add(nome.toLowerCase());
        tecnicos.push({ nome, tipo: texto(ws.getCell(r, colTipo)) });
      }
    }
  }
  const tipoDe = new Map(tecnicos.map((t) => [t.nome.toLowerCase(), t.tipo]));

  /* ---- Registos ---------------------------------------------------- */
  const registos = [];
  for (let r = linhaCab + 1; r <= ws.rowCount; r++) {
    const du = texto(ws.getCell(r, 5));
    if (!du) continue;
    const tecnico = texto(ws.getCell(r, 6));
    registos.push({
      periodo: texto(ws.getCell(r, 1)),
      dataReg: dataParaIso(ws.getCell(r, 2).value),
      totalTaxas: numero(ws.getCell(r, 3)),
      estado: texto(ws.getCell(r, 4)),
      numeroDU: du,
      tecnico,
      tipo: tecnico ? (tipoDe.get(tecnico.toLowerCase()) || '') : '',
      // Preserva o PDF de origem indicado no relatório; se faltar, fica o nome
      // do próprio relatório.
      ficheiro: texto(ws.getCell(r, 8)) || nomeFicheiro,
    });
  }

  if (!registos.length) throw new Error('O relatório não tem linhas de dados.');

  return { nome: nomeFicheiro, tecnicos, registos };
}

/** Junta os mapas de técnicos de vários relatórios, sem repetir nomes. */
export function juntarTecnicos(relatorios) {
  const juntos = [];
  const vistos = new Set();
  for (const rel of relatorios) {
    for (const t of rel.tecnicos || []) {
      const chave = t.nome.toLowerCase();
      if (vistos.has(chave)) continue;
      vistos.add(chave);
      juntos.push(t);
    }
  }
  return juntos;
}
