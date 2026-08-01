/*
 * Interface. Liga os formulários aos módulos: capta os registos de turno,
 * agrega-os nos relatórios e gera os .docx. Sem servidor: tudo no navegador.
 */
(function () {
  'use strict';
  const { docx, dados, registos, agregacao, modelos } = window.PMR;
  const JSZip = window.JSZip;
  let D = dados.carregar();

  // ---- ferramentas de DOM -------------------------------------------------
  function h(tag, attrs) {
    const e = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      const v = attrs[k];
      if (v == null) continue;
      if (k === 'class') e.className = v;
      else if (k === 'html') e.innerHTML = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function')
        e.addEventListener(k.slice(2).toLowerCase(), v);
      else e.setAttribute(k, v);
    }
    for (let i = 2; i < arguments.length; i++) {
      let kids = arguments[i];
      if (!Array.isArray(kids)) kids = [kids];
      kids.forEach((kid) => {
        if (kid == null || kid === false) return;
        e.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
      });
    }
    return e;
  }
  function campo(rot, input, cls) {
    return h('div', { class: 'campo ' + (cls || '') }, h('label', null, rot), input);
  }
  function inp(value, attrs) {
    return h('input', Object.assign({ value: value == null ? '' : value }, attrs || {}));
  }
  function num(value, attrs) {
    return inp(value == null ? 0 : value, Object.assign({ type: 'number', min: '0' }, attrs || {}));
  }
  function sel(valores, atual, attrs) {
    const s = h('select', attrs || {});
    valores.forEach((v) => {
      const o = h('option', { value: v }, v);
      if (v === atual) o.selected = true;
      s.append(o);
    });
    return s;
  }
  function limpar(el) { while (el.firstChild) el.removeChild(el.firstChild); }

  async function gerar(doc) {
    await docx.baixar(JSZip, doc.xml, doc.nome);
  }

  // =======================================================================
  // Navegação
  // =======================================================================
  const PAINEIS = [
    ['registos', 'Registos de Turno'],
    ['relatorios', 'Relatórios'],
    ['trb', 'Termo de Recepção'],
    ['ckl', 'Check-list'],
    ['dados', 'Dados-mestre']
  ];
  function mostrar(id) {
    document.querySelectorAll('.painel').forEach((p) =>
      p.classList.toggle('activo', p.id === 'painel-' + id));
    document.querySelectorAll('nav button').forEach((b) =>
      b.classList.toggle('activo', b.dataset.p === id));
    const render = { registos: renderRegistos, relatorios: renderRelatorios,
      trb: renderTRB, ckl: renderCKL, dados: renderDados };
    if (render[id]) render[id]();
  }

  // =======================================================================
  // Painel: Registos de Turno
  // =======================================================================
  function renderRegistos() {
    const raiz = document.getElementById('painel-registos');
    limpar(raiz);
    raiz.append(
      h('div', { class: 'cartao' },
        h('h2', null, 'Registos de Turno'),
        h('p', { class: 'ajuda' }, 'Introduza aqui, uma vez por turno, o movimento do posto. ' +
          'É desta fonte que os relatórios saem sozinhos, sem reescrever nada.'),
        h('div', { class: 'linha-botoes' },
          h('button', { class: 'acao', onclick: () => abrirEditor(registos.novoRegisto()) },
            '+ Novo registo'))));

    const lista = registos.listar();
    const tabela = h('div', { class: 'cartao' }, h('h3', null, 'Registos guardados'));
    if (!lista.length) {
      tabela.append(h('p', { class: 'vazio' }, 'Ainda não há registos.'));
    } else {
      const t = h('table', { class: 'lista' },
        h('tr', null, h('th', null, 'Data'), h('th', null, 'Turno'), h('th', null, 'Lado'),
          h('th', null, 'Voos'), h('th', null, 'Pass.'), h('th', null, 'Apre.'),
          h('th', null, 'Bag.'), h('th', null, '')));
      lista.forEach((r) => {
        t.append(h('tr', null,
          h('td', null, r.data), h('td', null, r.turno),
          h('td', null, h('span', { class: 'pill' }, r.lado)),
          h('td', null, r.voos || 0), h('td', null, r.passageiros || 0),
          h('td', null, (r.apreensoes || []).length), h('td', null, (r.bagagens || []).length),
          h('td', null,
            h('button', { class: 'sec', onclick: () => abrirEditor(clone(r)) }, 'Abrir'),
            ' ',
            h('button', { class: 'ligereza', onclick: () => {
              if (confirm('Apagar este registo?')) { registos.apagar(r.id); renderRegistos(); }
            } }, 'Apagar'))));
      });
      tabela.append(t);
    }
    raiz.append(tabela);
  }

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function abrirEditor(reg) {
    const raiz = document.getElementById('painel-registos');
    limpar(raiz);
    const chefes = D.chefesTurno;

    const cabec = h('div', { class: 'grelha g4' },
      campo('Data', inp(reg.data, { type: 'date', oninput: (e) => (reg.data = e.target.value) })),
      campo('Turno', sel(D.turnos, reg.turno, { oninput: (e) => (reg.turno = e.target.value) })),
      campo('Lado', sel(D.lados, reg.lado, { oninput: (e) => (reg.lado = e.target.value) })),
      campo('Chefe de Turno', chefeInput(reg)));

    const movimento = h('div', { class: 'grelha g4' },
      campo('Voos', num(reg.voos, { oninput: (e) => (reg.voos = +e.target.value) })),
      campo('Passageiros', num(reg.passageiros, { oninput: (e) => (reg.passageiros = +e.target.value) })),
      campo('Volumes (contentorização)', num(reg.volumes, { oninput: (e) => (reg.volumes = +e.target.value) })),
      campo('Visitas ao Catering', num(reg.visitasCatering, { oninput: (e) => (reg.visitasCatering = +e.target.value) })));
    const movimento2 = h('div', { class: 'grelha g4' },
      campo('Visitas a Bordo', num(reg.visitasBordo, { oninput: (e) => (reg.visitasBordo = +e.target.value) })),
      campo('Furgoneta — Scaneados', num(reg.furgoneta.scaneados, { oninput: (e) => (reg.furgoneta.scaneados = +e.target.value) })),
      campo('Furgoneta — Inspecionados', num(reg.furgoneta.inspecionados, { oninput: (e) => (reg.furgoneta.inspecionados = +e.target.value) })),
      campo('Furgoneta — Apreensões', num(reg.furgoneta.apreensoes, { oninput: (e) => (reg.furgoneta.apreensoes = +e.target.value) })));

    const sep = h('div', { class: 'grelha g3' },
      campo('Separados — Comercial', num(reg.separados.comercial, { oninput: (e) => (reg.separados.comercial = +e.target.value) })),
      campo('Separados — Não Comercial', num(reg.separados.naoComercial, { oninput: (e) => (reg.separados.naoComercial = +e.target.value) })),
      campo('Separados — Reg. a Posterior', num(reg.separados.regPosterior, { oninput: (e) => (reg.separados.regPosterior = +e.target.value) })));
    const tmp = h('div', { class: 'grelha g4' },
      campo('Importação Temporária', num(reg.temporarias.impTemp, { oninput: (e) => (reg.temporarias.impTemp = +e.target.value) })),
      campo('Exportação Temporária', num(reg.temporarias.expTemp, { oninput: (e) => (reg.temporarias.expTemp = +e.target.value) })),
      campo('Termo de Responsabilidade', num(reg.temporarias.termoResp, { oninput: (e) => (reg.temporarias.termoResp = +e.target.value) })),
      campo('Termo de Compromisso', num(reg.temporarias.termoComp, { oninput: (e) => (reg.temporarias.termoComp = +e.target.value) })));

    const caixaApre = h('div');
    const caixaBag = h('div');
    function pintarApre() {
      limpar(caixaApre);
      reg.apreensoes.forEach((a, i) => caixaApre.append(blocoApreensao(a, i, reg, pintarApre)));
    }
    function pintarBag() {
      limpar(caixaBag);
      reg.bagagens.forEach((b, i) => caixaBag.append(blocoBagagem(b, i, reg, pintarBag)));
    }
    pintarApre(); pintarBag();

    const editor = h('div', { class: 'cartao' },
      h('h2', null, 'Registo de turno'),
      cabec,
      h('h3', null, 'Movimento'), movimento, movimento2,
      h('h3', null, 'Separados de bagagem (lado Terra)'), sep,
      h('h3', null, 'Importações / Exportações temporárias (lado Terra)'), tmp,
      h('h3', null, 'Apreensões'), caixaApre,
      h('button', { class: 'sec', onclick: () => {
        reg.apreensoes.push({ tipo: D.tiposApreensao[0], passageiro: '', nacionalidade: '',
          origemDestino: '', voo: '', valor: '', descricao: '', realce: false, realceNota: '' });
        pintarApre();
      } }, '+ Apreensão'),
      h('h3', null, 'Bagagens retidas (geram Termo de Recepção)'), caixaBag,
      h('button', { class: 'sec', onclick: () => {
        reg.bagagens.push({ proveniencia: '', voo: '', dataVoo: reg.data, passageiro: '',
          etiqueta: '', companhia: '' });
        pintarBag();
      } }, '+ Bagagem'),
      h('div', { class: 'linha-botoes' },
        h('button', { class: 'acao', onclick: () => { registos.guardar(reg); renderRegistos(); } }, 'Guardar registo'),
        h('button', { class: 'sec', onclick: renderRegistos }, 'Cancelar')));
    raiz.append(editor);
  }

  function chefeInput(reg) {
    const dl = h('datalist', { id: 'chefes-dl' });
    D.chefesTurno.forEach((c) => dl.append(h('option', { value: c })));
    const i = inp(reg.chefeTurno, { list: 'chefes-dl', oninput: (e) => (reg.chefeTurno = e.target.value) });
    const wrap = h('div', null, i, dl);
    return wrap;
  }

  function blocoApreensao(a, i, reg, repinta) {
    return h('div', { class: 'sub' },
      h('button', { class: 'ligereza remover', onclick: () => { reg.apreensoes.splice(i, 1); repinta(); } }, '✕ remover'),
      h('div', { class: 'grelha g3' },
        campo('Tipo', sel(D.tiposApreensao, a.tipo, { oninput: (e) => (a.tipo = e.target.value) })),
        campo('Passageiro', inp(a.passageiro, { oninput: (e) => (a.passageiro = e.target.value) })),
        campo('Nacionalidade', inp(a.nacionalidade, { oninput: (e) => (a.nacionalidade = e.target.value) }))),
      h('div', { class: 'grelha g3' },
        campo('Proveniência / Destino', inp(a.origemDestino, { oninput: (e) => (a.origemDestino = e.target.value) })),
        campo('N.º do Voo', inp(a.voo, { oninput: (e) => (a.voo = e.target.value) })),
        campo('Valor (se moeda)', inp(a.valor, { oninput: (e) => (a.valor = e.target.value) }))),
      h('div', { class: 'grelha g2' },
        campo('Tipo de mercadoria / descrição', inp(a.descricao, { oninput: (e) => (a.descricao = e.target.value) })),
        campo('Maior realce (nota)', inp(a.realceNota, {
          oninput: (e) => { a.realceNota = e.target.value; a.realce = !!e.target.value; } }))));
  }

  function blocoBagagem(b, i, reg, repinta) {
    return h('div', { class: 'sub' },
      h('button', { class: 'ligereza remover', onclick: () => { reg.bagagens.splice(i, 1); repinta(); } }, '✕ remover'),
      h('div', { class: 'grelha g3' },
        campo('Proveniência', inp(b.proveniencia, { oninput: (e) => (b.proveniencia = e.target.value) })),
        campo('N.º do Voo', inp(b.voo, { oninput: (e) => (b.voo = e.target.value) })),
        campo('Data do Voo', inp(b.dataVoo, { type: 'date', oninput: (e) => (b.dataVoo = e.target.value) }))),
      h('div', { class: 'grelha g3' },
        campo('Passageiro', inp(b.passageiro, { oninput: (e) => (b.passageiro = e.target.value) })),
        campo('N.º de Etiqueta', inp(b.etiqueta, { oninput: (e) => (b.etiqueta = e.target.value) })),
        campo('Companhia Aérea', inp(b.companhia, { oninput: (e) => (b.companhia = e.target.value) }))),
      h('button', { class: 'sec', onclick: () => gerar(modelos.termoRecepcao(D.instituicao, {
        chefeAgt: reg.chefeTurno, chefePfa: '', companhia: b.companhia,
        dataTermo: b.dataVoo || reg.data, bagagens: [b]
      })) }, 'Gerar TRB desta bagagem'));
  }

  // =======================================================================
  // Painel: Relatórios (agregados dos registos)
  // =======================================================================
  function renderRelatorios() {
    const raiz = document.getElementById('painel-relatorios');
    limpar(raiz);
    const lista = registos.listar();
    const hoje = new Date().toISOString().slice(0, 10);
    const mesActual = hoje.slice(0, 7);

    const dataR = inp(hoje, { type: 'date' });
    const mesT = inp(mesActual, { type: 'month' });
    const mesA = inp(mesActual, { type: 'month' });

    raiz.append(
      h('div', { class: 'cartao' },
        h('h2', null, 'Relatórios automáticos'),
        h('p', { class: 'ajuda' }, 'Os relatórios são calculados a partir dos registos de turno. ' +
          'Não há nada a reescrever: escolha o período e gere.'),
        h('p', { class: 'aviso' }, `Há ${lista.length} registo(s) guardado(s) para agregar.`)),

      h('div', { class: 'cartao' },
        h('h3', null, 'RGA-24 — Relatório das últimas 24 h (gestão aeroportuária)'),
        h('div', { class: 'grelha g2' }, campo('Dia (07h → 07h do dia seguinte)', dataR)),
        h('div', { class: 'linha-botoes' },
          h('button', { class: 'acao', onclick: () =>
            gerar(modelos.relatorio24(D.instituicao, agregacao.rga24(registos.listar(), dataR.value))) },
            'Gerar RGA-24'))),

      h('div', { class: 'cartao' },
        h('h3', null, 'RMT — Relatório mensal do lado Terra'),
        h('div', { class: 'grelha g2' }, campo('Mês', mesT)),
        h('div', { class: 'linha-botoes' },
          h('button', { class: 'acao', onclick: () =>
            gerar(modelos.relatorioTerra(D.instituicao, agregacao.rmt(registos.listar(), mesT.value))) },
            'Gerar RMT'))),

      h('div', { class: 'cartao' },
        h('h3', null, 'RMA — Relatório mensal do lado Ar (grupo Scanner)'),
        h('div', { class: 'grelha g2' }, campo('Mês', mesA)),
        h('div', { class: 'linha-botoes' },
          h('button', { class: 'acao', onclick: () =>
            gerar(modelos.relatorioAr(D.instituicao, D.equipa, agregacao.rma(registos.listar(), mesA.value))) },
            'Gerar RMA'))));
  }

  // =======================================================================
  // Painel: TRB avulso
  // =======================================================================
  function renderTRB() {
    const raiz = document.getElementById('painel-trb');
    limpar(raiz);
    const b = { proveniencia: '', voo: '', dataVoo: new Date().toISOString().slice(0, 10),
      passageiro: '', etiqueta: '', companhia: '' };
    const d = { chefeAgt: '', chefePfa: '', dataTermo: new Date().toISOString().slice(0, 10) };
    raiz.append(h('div', { class: 'cartao' },
      h('h2', null, 'Termo de Recepção de Bagagens'),
      h('p', { class: 'ajuda' }, 'Emissão avulsa. Para ligar a bagagem a um turno, use antes o separador Registos.'),
      h('div', { class: 'grelha g3' },
        campo('Proveniência', inp('', { oninput: (e) => (b.proveniencia = e.target.value) })),
        campo('N.º do Voo', inp('', { oninput: (e) => (b.voo = e.target.value) })),
        campo('Data do Voo', inp(b.dataVoo, { type: 'date', oninput: (e) => (b.dataVoo = e.target.value) }))),
      h('div', { class: 'grelha g3' },
        campo('Passageiro', inp('', { oninput: (e) => (b.passageiro = e.target.value) })),
        campo('N.º de Etiqueta', inp('', { oninput: (e) => (b.etiqueta = e.target.value) })),
        campo('Companhia Aérea', inp('', { oninput: (e) => (d.companhia = e.target.value) }))),
      h('div', { class: 'grelha g3' },
        campo('Chefe de Turno AGT', inp('', { oninput: (e) => (d.chefeAgt = e.target.value) })),
        campo('Chefe de Turno PFA', inp('', { oninput: (e) => (d.chefePfa = e.target.value) })),
        campo('Data do Termo', inp(d.dataTermo, { type: 'date', oninput: (e) => (d.dataTermo = e.target.value) }))),
      h('div', { class: 'linha-botoes' },
        h('button', { class: 'acao', onclick: () =>
          gerar(modelos.termoRecepcao(D.instituicao, Object.assign({ bagagens: [b] }, d))) },
          'Gerar TRB'))));
  }

  // =======================================================================
  // Painel: CKL
  // =======================================================================
  function renderCKL() {
    const raiz = document.getElementById('painel-ckl');
    limpar(raiz);
    const dadosCkl = {
      turno: D.turnos[1], data: new Date().toISOString().slice(0, 10),
      itens: D.itensCkl.map((i) => Object.assign({ estado: 'Operacional', observacao: '' }, i)),
      ocorrencia: '', chefeEntrega: '', chefeRecebe: ''
    };
    const corpo = h('table', { class: 'lista' },
      h('tr', null, h('th', null, 'Descrição'), h('th', null, 'Qtd'),
        h('th', null, 'Estado'), h('th', null, 'Observação')));
    dadosCkl.itens.forEach((it) => {
      corpo.append(h('tr', null,
        h('td', null, it.descricao),
        h('td', null, inp(it.qtd, { style: 'width:60px', oninput: (e) => (it.qtd = e.target.value) })),
        h('td', null, sel(D.estados, it.estado, { oninput: (e) => (it.estado = e.target.value) })),
        h('td', null, inp(it.observacao, { oninput: (e) => (it.observacao = e.target.value) }))));
    });
    raiz.append(h('div', { class: 'cartao' },
      h('h2', null, 'Check-list de materiais e equipamentos'),
      h('p', { class: 'ajuda' }, 'O inventário-base já vem preenchido; só falta o estado, as observações e a ocorrência.'),
      h('div', { class: 'grelha g2' },
        campo('Turno', sel(D.turnos, dadosCkl.turno, { oninput: (e) => (dadosCkl.turno = e.target.value) })),
        campo('Data', inp(dadosCkl.data, { type: 'date', oninput: (e) => (dadosCkl.data = e.target.value) }))),
      h('h3', null, 'Itens'), corpo,
      campo('Ocorrência diária', h('textarea', { oninput: (e) => (dadosCkl.ocorrencia = e.target.value) })),
      h('div', { class: 'grelha g2' },
        campo('Chefe de Turno (entreguei)', inp('', { oninput: (e) => (dadosCkl.chefeEntrega = e.target.value) })),
        campo('Chefe de Turno (recebi)', inp('', { oninput: (e) => (dadosCkl.chefeRecebe = e.target.value) }))),
      h('div', { class: 'linha-botoes' },
        h('button', { class: 'acao', onclick: () => gerar(modelos.checklist(D.instituicao, dadosCkl)) },
          'Gerar Check-list'))));
  }

  // =======================================================================
  // Painel: Dados-mestre
  // =======================================================================
  function renderDados() {
    const raiz = document.getElementById('painel-dados');
    limpar(raiz);
    const inst = D.instituicao;
    const listaTxt = (arr) => arr.join('\n');
    const eqTA = h('textarea', { html: listaTxt(D.equipa) });
    const chTA = h('textarea', { html: listaTxt(D.chefesTurno) });
    const itTA = h('textarea', { html: D.itensCkl.map((i) => `${i.descricao} | ${i.qtd}`).join('\n') });

    function guardar() {
      inst.chefePostoCurto = qc.value; inst.cidade = qcid.value;
      inst.website = qw.value; inst.email = qe.value; inst.morada = qm.value;
      D.equipa = eqTA.value.split('\n').map((s) => s.trim()).filter(Boolean);
      D.chefesTurno = chTA.value.split('\n').map((s) => s.trim()).filter(Boolean);
      D.itensCkl = itTA.value.split('\n').map((s) => s.trim()).filter(Boolean).map((l) => {
        const [descricao, qtd] = l.split('|').map((x) => (x || '').trim());
        return { descricao, qtd: qtd || '-' };
      });
      dados.guardar(D);
      alert('Dados-mestre guardados neste navegador.');
    }

    const qc = inp(inst.chefePostoCurto);
    const qcid = inp(inst.cidade);
    const qw = inp(inst.website);
    const qe = inp(inst.email);
    const qm = inp(inst.morada);

    raiz.append(h('div', { class: 'cartao' },
      h('h2', null, 'Dados-mestre'),
      h('p', { class: 'ajuda' }, 'Valores que mudam raramente. Ficam guardados neste navegador e alimentam todos os documentos.'),
      h('div', { class: 'grelha g2' },
        campo('Chefe do Posto', qc), campo('Cidade (Terra/Luanda)', qcid)),
      h('div', { class: 'grelha g3' },
        campo('Website', qw), campo('E-mail', qe), campo('Morada (rodapé)', qm)),
      h('h3', null, 'Equipa (um nome por linha)'), eqTA,
      h('h3', null, 'Chefes de Turno (um por linha)'), chTA,
      h('h3', null, 'Itens da Check-list (descrição | quantidade)'), itTA,
      h('div', { class: 'linha-botoes' },
        h('button', { class: 'acao', onclick: guardar }, 'Guardar dados-mestre'),
        h('button', { class: 'sec', onclick: () => { D = dados.repor(); renderDados(); } }, 'Repor predefinições'))));

    raiz.append(h('div', { class: 'cartao' },
      h('h3', null, 'Cópia de segurança dos registos'),
      h('p', { class: 'ajuda' }, 'Os registos ficam neste computador. Exporte para os levar para outro, ou para guardar.'),
      h('div', { class: 'linha-botoes' },
        h('button', { class: 'sec', onclick: exportarRegistos }, 'Exportar registos (JSON)'),
        importarBotao())));
  }

  function exportarRegistos() {
    const blob = new Blob([JSON.stringify(registos.listar(), null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'registos-posto.json';
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 0);
  }
  function importarBotao() {
    const f = h('input', { type: 'file', accept: '.json', style: 'display:none',
      onchange: (e) => {
        const file = e.target.files[0]; if (!file) return;
        const fr = new FileReader();
        fr.onload = () => {
          try {
            const arr = JSON.parse(fr.result);
            if (Array.isArray(arr)) { registos.gravarTodos(arr); alert(arr.length + ' registo(s) importado(s).'); }
          } catch (err) { alert('Ficheiro inválido.'); }
        };
        fr.readAsText(file);
      } });
    const b = h('button', { class: 'sec', onclick: () => f.click() }, 'Importar registos (JSON)');
    return h('span', null, b, f);
  }

  // =======================================================================
  // Arranque
  // =======================================================================
  function iniciar() {
    const nav = document.getElementById('nav');
    PAINEIS.forEach(([id, rot]) => {
      const b = h('button', { 'data-p': id, onclick: () => mostrar(id) }, rot);
      nav.append(b);
      document.querySelector('main').append(h('section', { id: 'painel-' + id, class: 'painel' }));
    });
    mostrar('registos');
  }
  document.addEventListener('DOMContentLoaded', iniciar);
})();
