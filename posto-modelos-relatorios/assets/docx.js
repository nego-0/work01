/*
 * Motor de geração de .docx (OOXML) sobre o JSZip.
 * Um .docx é um ZIP de XML; este módulo constrói parágrafos e tabelas e
 * empacota tudo num documento Word válido. Corre no navegador e no Node.
 */
(function (global, factory) {
  const mod = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = mod;
  global.PMR = global.PMR || {};
  global.PMR.docx = mod;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

  // Largura útil da página A4 com margens de 1134 twips (≈2 cm) de cada lado.
  const LARGURA_UTIL = 9638;

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Uma "run" (troço de texto com formatação). o = {bold,italic,underline,size(pt),color,font}
  function runXml(text, o) {
    o = o || {};
    const rpr = [];
    if (o.font) rpr.push(`<w:rFonts w:ascii="${esc(o.font)}" w:hAnsi="${esc(o.font)}"/>`);
    if (o.bold) rpr.push('<w:b/>');
    if (o.italic) rpr.push('<w:i/>');
    if (o.color) rpr.push(`<w:color w:val="${o.color}"/>`);
    if (o.size) rpr.push(`<w:sz w:val="${o.size * 2}"/>`); // pt -> meios-pontos
    if (o.underline) rpr.push('<w:u w:val="single"/>');
    const rprXml = rpr.length ? `<w:rPr>${rpr.join('')}</w:rPr>` : '';
    // quebras de linha internas viram <w:br/>
    const partes = String(text == null ? '' : text).split('\n');
    const tXml = partes.map((p, i) =>
      (i ? '<w:br/>' : '') + `<w:t xml:space="preserve">${esc(p)}</w:t>`).join('');
    return `<w:r>${rprXml}${tXml}</w:r>`;
  }

  // Parágrafo. runs = array de {text, ...formato}. popt = {align,before,after,line}
  function paragraph(runs, popt) {
    popt = popt || {};
    const ppr = [];
    const sp = [];
    if (popt.before != null) sp.push(`w:before="${popt.before}"`);
    if (popt.after != null) sp.push(`w:after="${popt.after}"`);
    if (popt.line != null) sp.push(`w:line="${popt.line}" w:lineRule="auto"`);
    if (sp.length) ppr.push(`<w:spacing ${sp.join(' ')}/>`);
    if (popt.align) ppr.push(`<w:jc w:val="${popt.align}"/>`);
    const pprXml = ppr.length ? `<w:pPr>${ppr.join('')}</w:pPr>` : '';
    const rx = (runs || []).map((r) => runXml(r.text, r)).join('');
    return `<w:p>${pprXml}${rx}</w:p>`;
  }

  // Parágrafo. p('texto', {bold, align, after, ...}) para uma só run, ou
  // p([{text,bold}, {text,...}], {align, ...}) para várias runs numa linha.
  function p(text, opt) {
    opt = opt || {};
    if (Array.isArray(text)) return paragraph(text, opt);
    return paragraph([Object.assign({ text: text }, opt)], opt);
  }

  function espaco(after) { return p('', { after: after == null ? 120 : after }); }

  /*
   * Tabela.
   *   rows: array de linhas; cada linha é um array de células.
   *   célula: string, ou {text, bold, italic, align, valign, fill, colspan, size, color}
   *   opt: {widths:[dxa...], total, align, border:false, header:0}
   */
  function table(rows, opt) {
    opt = opt || {};
    const total = opt.total || LARGURA_UTIL;
    const ncol = opt.widths
      ? opt.widths.length
      : Math.max.apply(null, rows.map((r) =>
          r.reduce((a, c) => a + ((c && c.colspan) || 1), 0)));
    let widths = opt.widths;
    if (!widths) {
      const w = Math.floor(total / ncol);
      widths = new Array(ncol).fill(w);
    }
    const grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
    const bd = opt.border === false ? '' :
      `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
        .map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('')}</w:tblBorders>`;
    const tblPr = `<w:tblPr><w:tblW w:w="${total}" w:type="dxa"/>` +
      `${opt.align ? `<w:jc w:val="${opt.align}"/>` : ''}${bd}<w:tblLook w:val="0000"/></w:tblPr>`;

    const trs = rows.map((row, ri) => {
      let ci = 0;
      const cells = row.map((cell) => {
        const c = (cell && typeof cell === 'object') ? cell : { text: cell };
        const span = c.colspan || 1;
        let w = 0;
        for (let k = 0; k < span && ci < widths.length; k++) { w += widths[ci++]; }
        const tcpr = [`<w:tcW w:w="${w}" w:type="dxa"/>`];
        if (span > 1) tcpr.push(`<w:gridSpan w:val="${span}"/>`);
        if (c.fill) tcpr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${c.fill}"/>`);
        tcpr.push(`<w:vAlign w:val="${c.valign || 'center'}"/>`);
        const para = paragraph(
          [{ text: c.text == null ? '' : c.text, bold: c.bold, italic: c.italic, size: c.size, color: c.color }],
          { align: c.align || 'left' });
        return `<w:tc><w:tcPr>${tcpr.join('')}</w:tcPr>${para}</w:tc>`;
      }).join('');
      const cabec = (opt.header && ri < opt.header) ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
      return `<w:tr>${cabec}${cells}</w:tr>`;
    }).join('');

    // O Word exige um parágrafo a seguir a cada tabela.
    return `<w:tbl>${tblPr}${grid}${trs}</w:tbl><w:p/>`;
  }

  function documento(blocos) {
    const body = (blocos || []).join('');
    const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" ' +
      'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<w:document xmlns:w="${NS}"><w:body>${body}${sect}</w:body></w:document>`;
  }

  const CONTENT_TYPES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>';

  const RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  const DOC_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:styles xmlns:w="${NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>' +
    '</w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '</w:styles>';

  // Devolve um objecto JSZip pronto a gerar. JSZip é passado como argumento
  // (global no navegador, importado no Node).
  function empacotar(JSZip, documentXml) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', CONTENT_TYPES);
    zip.folder('_rels').file('.rels', RELS);
    const w = zip.folder('word');
    w.file('document.xml', documentXml);
    w.file('styles.xml', STYLES);
    w.folder('_rels').file('document.xml.rels', DOC_RELS);
    return zip;
  }

  // Gera o .docx e desencadeia a transferência (só no navegador).
  async function baixar(JSZip, documentXml, nome) {
    const zip = empacotar(JSZip, documentXml);
    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = nome.endsWith('.docx') ? nome : nome + '.docx';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 0);
  }

  // '2026-07-06' -> '06 de Julho de 2026'
  function porExtenso(iso) {
    if (!iso) return '';
    const partes = iso.split('-').map(Number);
    const y = partes[0], m = partes[1], d = partes[2];
    return `${String(d).padStart(2, '0')} de ${MESES[m - 1]} de ${y}`;
  }

  // '2026-07-06' -> 'dia 06 de Julho'
  function diaExtenso(iso) {
    const partes = iso.split('-').map(Number);
    return `${String(partes[2]).padStart(2, '0')} de ${MESES[partes[1] - 1]}`;
  }

  function proximoDia(iso) {
    const d = new Date(iso + 'T00:00:00');
    d.setDate(d.getDate() + 1);
    return d.toISOString().slice(0, 10);
  }

  return {
    esc, runXml, paragraph, p, espaco, table, documento, empacotar, baixar,
    porExtenso, diaExtenso, proximoDia, MESES, LARGURA_UTIL
  };
});
