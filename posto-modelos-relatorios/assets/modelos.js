/*
 * Os cinco modelos. Cada função recebe dados e a instituição e devolve
 * { nome, xml } — o nome do ficheiro e o corpo do documento .docx.
 * O bloco institucional e as linhas de assinatura são reaproveitados por todos.
 */
(function (global, factory) {
  const mod = factory(global);
  if (typeof module !== 'undefined' && module.exports) module.exports = mod;
  global.PMR = global.PMR || {};
  global.PMR.modelos = mod;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (G) {
  'use strict';

  function dx() { return G.PMR.docx; }

  // 12345 -> "12.345"
  function milhar(n) {
    const s = String(Math.round(Number(n) || 0));
    return s.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }

  // '2026-07-06' -> '06.07.2026'
  function dataCurta(iso) {
    if (!iso) return '';
    const [y, m, d] = iso.split('-');
    return `${d}.${m}.${y}`;
  }

  function ultimoDia(mes /* 'YYYY-MM' */) {
    const [y, m] = mes.split('-').map(Number);
    const dia = new Date(y, m, 0).getDate();
    return `${mes}-${String(dia).padStart(2, '0')}`;
  }

  function nomeMesAno(mes) {
    const { MESES } = dx();
    const [y, m] = mes.split('-').map(Number);
    return `${MESES[m - 1]} de ${y}`;
  }

  // --- blocos reutilizáveis -------------------------------------------------

  const LINHA = '_________________________________';

  function vistoChefePosto(inst) {
    const { p } = dx();
    return [
      p('Visto pelo Chefe do Posto', { align: 'right', after: 0 }),
      p(LINHA, { align: 'right', after: 0 }),
      p(inst.chefePostoCurto || inst.chefePosto, { align: 'right', after: 240 })
    ];
  }

  function assinatura(rotulo, nome, align) {
    const { p } = dx();
    align = align || 'left';
    const b = [p(rotulo, { align: align, before: 240, after: 0 }),
      p(LINHA, { align: align, after: 0 })];
    if (nome) b.push(p(nome, { align: align, bold: true, after: 120 }));
    return b;
  }

  function cabecalhoCentral(inst, extra) {
    const { p } = dx();
    const b = [
      p(inst.regiao, { align: 'center', bold: true, after: 0, size: 12 }),
      p(inst.posto, { align: 'center', bold: true, after: 120, size: 12 })
    ];
    (extra || []).forEach((linha) => b.push(
      p(linha, { align: 'center', bold: true, after: 0, size: 12 })));
    return b;
  }

  function rodapeInstituicao(inst) {
    const { p } = dx();
    return [
      dx().espaco(240),
      p(inst.postoCurto + '.', { align: 'center', italic: true, size: 8, after: 0 }),
      p([{ text: inst.morada }].map((r) => r.text).join(''),
        { align: 'center', italic: true, size: 8, after: 0 }),
      p(`Website: ${inst.website}  |  E-mail: ${inst.email}`,
        { align: 'center', italic: true, size: 8, after: 0 })
    ];
  }

  // Linha "nada a assinalar" para tabelas vazias.
  function linhaVazia(ncol) {
    const c = [{ text: '-', align: 'center' }, { text: '0', align: 'center' }];
    while (c.length < ncol) c.push({ text: '----------', align: 'center' });
    return c;
  }

  // =========================================================================
  // TRB — Termo de Recepção de Bagagens (entregue à companhia aérea)
  // =========================================================================
  function termoRecepcao(inst, dados) {
    const { p, espaco, porExtenso, documento } = dx();
    const bags = dados.bagagens || [];
    const b = [];
    b.push(p(inst.postoCurto, { align: 'center', bold: true, after: 240, size: 11 }));
    b.push(p('TERMO DE RECEPÇÃO', { align: 'center', bold: true, size: 14, after: 240 }));

    const ref = bags[0] || {};
    const plural = bags.length > 1;
    const intro = plural
      ? `No âmbito das nossas actividades, foram retidas ${bags.length} bagagens com ` +
        `interesse aduaneiro, com os seguintes nomes e n.os de etiqueta:`
      : `No âmbito das nossas actividades, foi retida 1 bagagem com interesse ` +
        `aduaneiro, proveniente de ${ref.proveniencia || '—'}, do vôo ${ref.voo || '—'}, ` +
        `do dia ${porExtenso(ref.dataVoo) || '—'}, com o seguinte nome e n.º de etiqueta:`;
    b.push(p(intro, { align: 'both', after: 200 }));

    bags.forEach((bag) => {
      const linha = plural
        ? `${bag.passageiro || '—'} – ${bag.etiqueta || '—'}  (${bag.proveniencia || '—'}, ` +
          `vôo ${bag.voo || '—'}, ${porExtenso(bag.dataVoo) || '—'})`
        : `${bag.passageiro || '—'} – ${bag.etiqueta || '—'}`;
      b.push(p(linha, { align: 'center', bold: true, after: 200 }));
    });

    b.push(p('Por ser verdade, lavrou-se o presente termo de recepção que vai ' +
      'devidamente assinado pelos responsáveis das áreas.', { align: 'both', after: 240 }));

    b.push.apply(b, assinatura('O Chefe de Turno AGT', dados.chefeAgt || ''));
    b.push.apply(b, assinatura('O Chefe de Turno PFA', dados.chefePfa || ''));
    b.push.apply(b, assinatura('Pela Companhia Aérea', dados.companhia || ''));

    b.push(espaco(240));
    b.push(p(`${inst.cidade}, aos ${porExtenso(dados.dataTermo || (bags[0] && bags[0].dataVoo))}`,
      { align: 'left' }));

    return { nome: `TRB_${(ref.passageiro || 'bagagem').replace(/\s+/g, '_')}`, xml: documento(b) };
  }

  // =========================================================================
  // CKL — Check-list de materiais e equipamentos (fim de turno)
  // =========================================================================
  function checklist(inst, dados) {
    const { p, table, espaco, documento } = dx();
    const b = [];
    b.push.apply(b, vistoChefePosto(inst));
    b.push(p('CHECK-LIST | MATERIAIS E EQUIPAMENTOS DE SERVIÇO',
      { align: 'center', bold: true, size: 13, after: 200 }));
    b.push(p(`TURNO: ${dados.turno || ''}                                     ` +
      `DATA: ${dataCurta(dados.data)}`, { after: 200 }));

    const largs = [4200, 900, 1800, 2738];
    const linhas = [[
      { text: 'DESCRIÇÃO', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'QTD', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'ESTADO', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'OBSERVAÇÃO', bold: true, align: 'center', fill: 'D9D9D9' }
    ]];
    (dados.itens || []).forEach((it) => linhas.push([
      { text: it.descricao || '' },
      { text: it.qtd || '', align: 'center' },
      { text: it.estado || '', align: 'center' },
      { text: it.observacao || '' }
    ]));
    b.push(table(linhas, { widths: largs, header: 1 }));

    b.push(p([{ text: 'Ocorrência Diária: ', bold: true },
      { text: dados.ocorrencia || 'Sem ocorrências a assinalar.' }], { before: 120, after: 240 }));

    b.push.apply(b, assinatura('Ass. Chefe de Turno | entreguei:', dados.chefeEntrega || ''));
    b.push.apply(b, assinatura('Ass. Chefe de Turno | recebi:', dados.chefeRecebe || ''));
    return { nome: `CKL_Turno_${dados.turno || ''}_${dados.data || ''}`, xml: documento(b) };
  }

  // =========================================================================
  // RGA-24 — Relatório-síntese das últimas 24 h (gestão aeroportuária)
  // =========================================================================
  function relatorio24(inst, ag) {
    const { p, table, espaco, porExtenso, diaExtenso, documento } = dx();
    const ano = (ag.dia || '').slice(0, 4);
    const b = [];
    b.push.apply(b, vistoChefePosto(inst));
    b.push(p([{ text: 'Assunto: ', bold: true },
      { text: `Relatório Síntese das Apreensões realizadas no Período de ` +
        `${diaExtenso(ag.dia)} à ${diaExtenso(ag.diaSeguinte)} de ${ano}.`, bold: true }],
      { after: 200 }));
    b.push(p(`O ${inst.postoCurto}, vem por meio deste relatar as principais ocorrências ` +
      `registadas no período compreendido entre às 07h00 do dia ${diaExtenso(ag.dia)} até às ` +
      `07h00 do dia ${diaExtenso(ag.diaSeguinte)} do ano ${ano}, o seguinte:`,
      { align: 'both', after: 200 }));

    // Secção 1 — Valores
    b.push(p('Apreensão de Valores', { bold: true, after: 80 }));
    const hV = [[
      { text: 'N.º', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Passageiros Envolvidos', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Nacionalidade', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Destino', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'N.º do Voo', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Valor', bold: true, align: 'center', fill: 'D9D9D9' }
    ]];
    if (ag.valores.length) {
      ag.valores.forEach((a, i) => hV.push([
        { text: String(i + 1).padStart(2, '0'), align: 'center' },
        { text: a.passageiro || '' },
        { text: a.nacionalidade || '', align: 'center' },
        { text: a.origemDestino || '', align: 'center' },
        { text: a.voo || '', align: 'center' },
        { text: a.valor || '', align: 'center' }
      ]));
    } else { hV.push(linhaVazia(6)); }
    b.push(table(hV, { widths: [700, 2600, 1700, 1500, 1200, 1938], header: 1 }));
    b.push(p([{ text: 'Apreensão de Valores com maior Realce: ', bold: true },
      { text: ag.realces.length ? ag.realces.join('; ') : 'Nada a realçar.' }], { after: 160 }));

    // Secção 2 — Mercadoria sujeita a direitos
    b.push(p('Apreensão de Mercadoria Sujeita a Pagamento de Direitos', { bold: true, after: 80 }));
    b.push(seccaoMercadoria(inst, ag.mercadoria));
    b.push(p([{ text: 'Apreensão de Mercadorias com maior Realce: ', bold: true },
      { text: 'Nada a realçar.' }], { after: 160 }));

    // Secção 3 — Mercadoria proibida / restrita
    b.push(p('Apreensão de Mercadoria Proibida - Restrita', { bold: true, after: 80 }));
    b.push(seccaoMercadoria(inst, ag.proibida));
    b.push(p([{ text: 'Apreensão de Mercadorias com maior Realce: ', bold: true },
      { text: 'Nada a realçar.' }], { after: 200 }));

    b.push(p('Nada mais a relatar de momento, reiteramos cordiais saudações.',
      { align: 'both', after: 200 }));
    b.push(p(`${inst.cidade}, ${porExtenso(ag.diaSeguinte)}.`, { after: 200 }));
    b.push.apply(b, assinatura('O CHEFE DE TURNO', ag.chefeTurno || '', 'center'));
    return { nome: `RGA24_${ag.dia}`, xml: documento(b) };
  }

  function seccaoMercadoria(inst, lista) {
    const { table } = dx();
    const h = [[
      { text: 'N.º', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Passageiros Envolvidos', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Nacionalidade', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Proveniência', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'N.º do Voo', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Tipo de Mercadoria', bold: true, align: 'center', fill: 'D9D9D9' }
    ]];
    if (lista && lista.length) {
      lista.forEach((a, i) => h.push([
        { text: String(i + 1).padStart(2, '0'), align: 'center' },
        { text: a.passageiro || '' },
        { text: a.nacionalidade || '', align: 'center' },
        { text: a.origemDestino || '', align: 'center' },
        { text: a.voo || '', align: 'center' },
        { text: a.descricao || '' }
      ]));
    } else { h.push(linhaVazia(6)); }
    return table(h, { widths: [700, 2400, 1600, 1600, 1200, 2138], header: 1 });
  }

  // =========================================================================
  // RMT — Relatório mensal do lado Terra
  // =========================================================================
  function relatorioTerra(inst, ag) {
    const { p, table, porExtenso, documento } = dx();
    const fim = ultimoDia(ag.mes);
    const b = [];
    b.push.apply(b, vistoChefePosto(inst));
    b.push.apply(b, cabecalhoCentral(inst));
    b.push(dx().espaco(120));
    b.push(p(`RELATÓRIO DE ACTIVIDADES DURANTE O MÊS DE ${nomeMesAno(ag.mes).toUpperCase()}`,
      { align: 'center', bold: true, size: 12, after: 200 }));

    b.push(p(`No período compreendido entre os dias 01 à ${fim.slice(8)} de ${nomeMesAno(ag.mes)}, ` +
      `o Turno ${ag.turno || ''} efectuou o seguinte movimento:`, { align: 'both', after: 120 }));
    b.push(p(`- Chegaram ao nosso Aeroporto ${milhar(ag.voos)} Voos com ` +
      `${milhar(ag.passageiros)} passageiros;`, { after: 200 }));

    b.push(p('APREENSÕES', { bold: true, after: 60 }));
    b.push(p('Realçamos que no período em causa foram feitas as seguintes apreensões:',
      { align: 'both', after: 80 }));
    const a = ag.apreensoes;
    const totApre = a.moeda + a.diversa + a.proibida + a.dutyFree;
    b.push(tabelaContagem([
      ['01', '(Moeda)', a.moeda], ['02', 'Mercadoria Diversa', a.diversa],
      ['03', 'Mercadoria Proibida', a.proibida], ['04', 'Duty Free', a.dutyFree]
    ], totApre));

    b.push(p('SEPARADOS DE BAGAGEM', { bold: true, after: 60 }));
    b.push(p('Realçamos que no período em causa foram feitos os seguintes separados:',
      { align: 'both', after: 80 }));
    const s = ag.separados;
    const totSep = s.comercial + s.naoComercial + s.regPosterior;
    b.push(tabelaContagem([
      ['01', 'Comercial', s.comercial], ['02', 'Não Comercial', s.naoComercial],
      ['03', 'Regularização a Posterior', s.regPosterior]
    ], totSep));

    b.push(p('IMPORTAÇÕES E EXPORTAÇÕES TEMPORÁRIAS', { bold: true, after: 60 }));
    b.push(p('Registou-se ainda as seguintes importações e exportações temporárias:',
      { align: 'both', after: 80 }));
    const t = ag.temporarias;
    const totTmp = t.impTemp + t.expTemp + t.termoResp + t.termoComp;
    b.push(tabelaContagem([
      ['01', 'Importação Temporária', t.impTemp], ['02', 'Exportação Temporária', t.expTemp],
      ['04', 'Termo de Responsabilidade', t.termoResp], ['05', 'Termo de Compromisso', t.termoComp]
    ], totTmp));

    b.push(p(`${inst.cidade}, ${porExtenso(fim)}.`, { after: 200 }));
    b.push.apply(b, assinatura('O CHEFE DO TURNO', ag.chefeTurno || '', 'center'));
    b.push.apply(b, rodapeInstituicao(inst));
    return { nome: `RMT_${ag.mes}`, xml: documento(b) };
  }

  function tabelaContagem(linhas, total) {
    const { table } = dx();
    const h = [[
      { text: 'N.º', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Descrição', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Total', bold: true, align: 'center', fill: 'D9D9D9' },
      { text: 'Observação', bold: true, align: 'center', fill: 'D9D9D9' }
    ]];
    linhas.forEach((l) => h.push([
      { text: l[0], align: 'center' }, { text: l[1] },
      { text: milhar(l[2]), align: 'center' }, { text: '-', align: 'center' }
    ]));
    h.push([{ text: 'TOTAL GERAL', bold: true, colspan: 2 },
      { text: milhar(total), bold: true, align: 'center' }, { text: '-', align: 'center' }]);
    return table(h, { widths: [900, 4500, 1600, 2638], header: 1 });
  }

  // =========================================================================
  // RMA — Relatório mensal do lado Ar (grupo Scanner)
  // =========================================================================
  function relatorioAr(inst, equipa, ag) {
    const { p, table, porExtenso, documento } = dx();
    const fim = ultimoDia(ag.mes);
    const b = [];
    b.push.apply(b, vistoChefePosto(inst));
    b.push.apply(b, cabecalhoCentral(inst, [inst.sufixoAr || 'DE ICOLO E BENGO']));
    b.push(p(`Grupo Scanner – Turno - ${ag.turno || ''}`, { align: 'center', bold: true, after: 60 }));
    b.push(p(`RELATÓRIO REFERENTE AO MÊS DE ${nomeMesAno(ag.mes).toUpperCase()}`,
      { align: 'center', bold: true, size: 12, after: 160 }));

    b.push(p([{ text: 'Área de Operação: ', bold: true },
      { text: 'Aeroporto Internacional António Agostinho Neto' }], { after: 80 }));
    b.push(p([{ text: 'Nome do pessoal em serviço: ', bold: true },
      { text: (equipa || []).join(', ') + '.' }], { align: 'both', after: 80 }));
    b.push(p([{ text: 'Caso significante durante o mês: ', bold: true },
      { text: ag.casoSignificante || 'Somos a realçar que as nossas actividades foram ' +
        'realizadas nas seguintes áreas: Placa, Cabines, Contentorização.' }],
      { align: 'both', after: 200 }));

    b.push(table([
      [
        { text: 'Número de Voos Processados', bold: true, align: 'center', fill: 'D9D9D9' },
        { text: 'Número de Passageiros', bold: true, align: 'center', fill: 'D9D9D9' },
        { text: 'N.º de Volumes Inspecionados na Contentorização', bold: true, align: 'center', fill: 'D9D9D9' },
        { text: 'Apreensões', bold: true, align: 'center', fill: 'D9D9D9' },
        { text: 'Visitas ao Catering', bold: true, align: 'center', fill: 'D9D9D9' },
        { text: 'Visitas a Bordo', bold: true, align: 'center', fill: 'D9D9D9' }
      ],
      [
        { text: milhar(ag.voos), align: 'center' },
        { text: milhar(ag.passageiros), align: 'center' },
        { text: milhar(ag.volumes), align: 'center' },
        { text: milhar(ag.apreensoes), align: 'center' },
        { text: milhar(ag.visitasCatering), align: 'center' },
        { text: milhar(ag.visitasBordo), align: 'center' }
      ]
    ], { header: 1 }));

    b.push(p('DADOS DA FURGONETA', { bold: true, before: 120, after: 60 }));
    const f = ag.furgoneta;
    b.push(table([
      [
        { text: 'Volumes Scaneados', bold: true, align: 'center', fill: 'D9D9D9' },
        { text: 'Volumes Inspecionados', bold: true, align: 'center', fill: 'D9D9D9' },
        { text: 'Apreensões', bold: true, align: 'center', fill: 'D9D9D9' }
      ],
      [
        { text: milhar(f.scaneados), align: 'center' },
        { text: milhar(f.inspecionados), align: 'center' },
        { text: milhar(f.apreensoes), align: 'center' }
      ]
    ], { header: 1 }));

    b.push(p(`Icolo e Bengo, ${porExtenso(fim)}.`, { before: 160, after: 200 }));
    b.push.apply(b, assinatura('O CHEFE DO TURNO', ag.chefeTurno || '', 'center'));
    return { nome: `RMA_${ag.mes}`, xml: documento(b) };
  }

  return {
    termoRecepcao, checklist, relatorio24, relatorioTerra, relatorioAr,
    milhar, dataCurta, ultimoDia, nomeMesAno
  };
});
