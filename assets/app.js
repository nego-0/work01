/**
 * app.js — interface do gerador de relatórios.
 */

import * as pdfjsLib from '../vendor/pdf.min.mjs';
import {
  parseFileName,
  extractRowsFromDocument,
  agruparPorDiasConsecutivos,
  nomeDoGrupo,
  rotuloDoGrupo,
  formatDatePt,
  parseAmount,
} from './parser.js';
import { criarWorkbook } from './report.js';
import { lerRelatorio, juntarTecnicos, juntarRegistos, POLITICAS } from './leitor-relatorio.js';

if (pdfjsLib.GlobalWorkerOptions && !pdfjsLib.GlobalWorkerOptions.workerSrc
    && typeof globalThis.pdfjsWorker === 'undefined') {
  // Só na versão em vários ficheiros: o single-file traz o worker embutido.
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdf.worker.min.mjs', import.meta.url).href;
}

/**
 * Técnicos que aparecem por omissão na área de configuração. Ficam guardados no
 * navegador assim que forem alterados; o botão "Repor lista" traz esta de volta.
 */
const TECNICOS_PREDEFINIDOS = [
  'Abednego Agostinho',
  'Raquel Machado',
  'Leopoldo Maiato',
  'Ruth Contreiras',
  'Gisela Antonio',
  'Vania Chungo',
  'Jeronimo dos Santos',
  'Mario Massanga',
  'Altair Pereira',
  'Marcolino da Silva',
  'Carolina Costa',
  'Constancia Cortez',
  'Felson Jorge',
  'Alfredo Jose',
  'Jesse Martins',
  'Virgilio da Conceição',
].map((nome) => ({ nome, tipo: '' }));

const CHAVE_GUARDADA = 'gerador-relatorios:tecnicos';

const $ = (id) => document.getElementById(id);
const seleccionado = (nome) => document.querySelector(`input[name="${nome}"]:checked`).value;

const estado = {
  fonte: 'pdf',    // 'pdf' | 'relatorio' | 'acrescentar'
  base: null,      // consolidado a acrescer, quando a fonte é 'acrescentar'
  ficheiros: [],   // { id, nome, periodo, dataIso, file, registos, tecnicos, erro }
  gerados: [],     // { nome, blob, grupo }
};

/** As fontes que não são PDF juntam relatórios e verificam redundâncias. */
const juntaRelatorios = () => estado.fonte !== 'pdf';

/* ------------------------------------------------------------------ */
/* Origem                                                              */
/* ------------------------------------------------------------------ */

const zona = $('zona');
const inputFicheiros = $('input-ficheiros');

function extensaoAceite() {
  return estado.fonte === 'pdf' ? /\.pdf$/i : /\.xlsx$/i;
}

function aplicarFonte() {
  estado.fonte = seleccionado('fonte');
  const ehPdf = estado.fonte === 'pdf';
  const ehAcrescentar = estado.fonte === 'acrescentar';

  inputFicheiros.accept = ehPdf ? 'application/pdf,.pdf' : '.xlsx';
  $('zona-nota').innerHTML = ehPdf
    ? 'O nome deve conter a data e o período — ex.: <code>Manhã 21.07.2026.pdf</code>, '
      + '<code>01.07.2026_Tarde.pdf</code>, <code>2026-07-23 Madrugada.pdf</code>'
    : 'Relatórios <code>.xlsx</code> gerados por esta aplicação. Os técnicos e os estados que '
      + 'lá estiverem preenchidos são mantidos.';

  $('area-base').hidden = !ehAcrescentar;
  $('area-politica').hidden = ehPdf;

  // A partir de relatórios só faz sentido consolidar.
  for (const radio of document.querySelectorAll('input[name="modo"]')) {
    radio.disabled = !ehPdf;
    if (!ehPdf) radio.checked = radio.value === 'consolidado';
  }
  document.querySelectorAll('input[name="modo"]')[0].closest('.opcoes')
    .classList.toggle('desactivada', !ehPdf);
  $('gerar-consolidado').disabled = !ehPdf;
  $('campo-consolidado').classList.toggle('desactivada', !ehPdf);

  estado.ficheiros = [];
  estado.base = null;
  desenharBase();
  desenharFicheiros();
}

for (const radio of document.querySelectorAll('input[name="fonte"]')) {
  radio.addEventListener('change', aplicarFonte);
}

zona.addEventListener('click', () => inputFicheiros.click());
zona.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inputFicheiros.click(); }
});
inputFicheiros.addEventListener('change', () => {
  adicionarFicheiros([...inputFicheiros.files]);
  inputFicheiros.value = '';
});

for (const ev of ['dragenter', 'dragover']) {
  zona.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.add('activa'); });
}
for (const ev of ['dragleave', 'drop']) {
  zona.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.remove('activa'); });
}
zona.addEventListener('drop', (e) => {
  adicionarFicheiros([...(e.dataTransfer?.files || [])]);
});

$('btn-limpar').addEventListener('click', () => {
  estado.ficheiros = [];
  desenharFicheiros();
});

/* ---- Consolidado a acrescer -------------------------------------- */

const zonaBase = $('zona-base');
const inputBase = $('input-base');

zonaBase.addEventListener('click', () => inputBase.click());
zonaBase.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inputBase.click(); }
});
inputBase.addEventListener('change', () => {
  definirBase(inputBase.files[0]);
  inputBase.value = '';
});
for (const ev of ['dragenter', 'dragover']) {
  zonaBase.addEventListener(ev, (e) => { e.preventDefault(); zonaBase.classList.add('activa'); });
}
for (const ev of ['dragleave', 'drop']) {
  zonaBase.addEventListener(ev, (e) => { e.preventDefault(); zonaBase.classList.remove('activa'); });
}
zonaBase.addEventListener('drop', (e) => definirBase((e.dataTransfer?.files || [])[0]));

async function definirBase(file) {
  if (!file || !/\.xlsx$/i.test(file.name)) return;
  const item = {
    id: `base${++contador}`, nome: file.name, file,
    registos: null, tecnicos: [], erro: null, aLer: true, ehBase: true,
  };
  estado.base = item;
  desenharBase();
  await lerRelatorioExcel(item);
  desenharBase();
  desenharFicheiros();
}

function desenharBase() {
  const alvo = $('info-base');
  alvo.innerHTML = '';
  const b = estado.base;
  alvo.hidden = !b;
  if (!b) return;

  const div = document.createElement('div');
  div.className = 'ficheiro' + (b.erro ? ' com-erro' : '');
  const info = document.createElement('div');
  info.className = 'ficheiro-info';
  info.innerHTML = `<div class="ficheiro-nome">📗 ${b.nome}</div>`;
  const meta = document.createElement('div');
  meta.className = 'ficheiro-meta';
  meta.textContent = b.aLer ? 'A ler…'
    : b.erro ? b.erro
      : `${b.registos.length} registo(s) · `
        + `${b.registos.filter((r) => r.tecnico).length} com técnico · `
        + `${b.tecnicos.length} técnico(s) no mapa`;
  info.appendChild(meta);

  const remover = document.createElement('button');
  remover.type = 'button';
  remover.className = 'btn-remover';
  remover.title = 'Remover';
  remover.textContent = '✕';
  remover.addEventListener('click', () => {
    estado.base = null;
    desenharBase();
    desenharFicheiros();
  });

  div.append(info, remover);
  alvo.appendChild(div);
}

let contador = 0;

async function adicionarFicheiros(files) {
  const aceite = extensaoAceite();
  for (const file of files) {
    if (!aceite.test(file.name)) continue;
    if (estado.ficheiros.some((f) => f.nome === file.name && f.file.size === file.size)) continue;

    const { periodo, dataIso } = parseFileName(file.name);
    const item = {
      id: `f${++contador}`,
      nome: file.name,
      periodo,
      dataIso,
      file,
      registos: null,
      tecnicos: [],
      erro: null,
      aLer: true,
    };
    estado.ficheiros.push(item);
    desenharFicheiros();
    if (estado.fonte === 'pdf') await lerPdf(item);
    else await lerRelatorioExcel(item);
    desenharFicheiros();
  }
}

async function lerPdf(item) {
  try {
    const buffer = await item.file.arrayBuffer();
    const tarefa = pdfjsLib.getDocument({ data: new Uint8Array(buffer) });
    const doc = await tarefa.promise;
    item.registos = await extractRowsFromDocument(doc);
    await tarefa.destroy();

    if (!item.registos.length) {
      item.erro = 'Nenhum registo encontrado (o PDF não parece ser um Document List).';
    } else if (!item.dataIso) {
      item.dataIso = dataPredominante(item.registos);
      item.dataDeduzida = true;
    }
  } catch (err) {
    console.error(err);
    item.erro = `Erro ao ler o PDF: ${err.message}`;
    item.registos = [];
  } finally {
    item.aLer = false;
  }
}

async function lerRelatorioExcel(item) {
  try {
    const buffer = await item.file.arrayBuffer();
    const rel = await lerRelatorio(window.ExcelJS, buffer, item.nome);
    item.registos = rel.registos;
    item.tecnicos = rel.tecnicos;
    item.dataIso = item.dataIso || dataPredominante(rel.registos);
    item.periodo = '';   // cada linha traz o seu período
  } catch (err) {
    console.error(err);
    item.erro = `Erro ao ler o relatório: ${err.message}`;
    item.registos = [];
  } finally {
    item.aLer = false;
  }
}

function dataPredominante(registos) {
  const contagem = {};
  for (const r of registos) if (r.dataReg) contagem[r.dataReg] = (contagem[r.dataReg] || 0) + 1;
  const entradas = Object.entries(contagem).sort((a, b) => b[1] - a[1]);
  return entradas.length ? entradas[0][0] : '';
}

function desenharFicheiros() {
  const lista = $('lista-ficheiros');
  lista.innerHTML = '';
  const temFicheiros = estado.ficheiros.length > 0;
  lista.hidden = !temFicheiros;
  $('acoes-ficheiros').hidden = !temFicheiros;

  const ordenados = [...estado.ficheiros].sort(
    (a, b) => String(a.dataIso).localeCompare(String(b.dataIso)) || a.nome.localeCompare(b.nome)
  );

  for (const f of ordenados) {
    const problemas = f.registos ? f.registos.filter(temProblema) : [];

    const bloco = document.createElement('div');
    bloco.className = 'ficheiro-bloco';

    const div = document.createElement('div');
    div.className = 'ficheiro'
      + (f.erro ? ' com-erro' : problemas.length ? ' com-aviso' : '');

    const info = document.createElement('div');
    info.className = 'ficheiro-info';

    const nome = document.createElement('div');
    nome.className = 'ficheiro-nome';
    nome.textContent = f.nome;
    info.appendChild(nome);

    const meta = document.createElement('div');
    meta.className = 'ficheiro-meta';
    if (f.aLer) {
      meta.textContent = 'A ler…';
    } else if (f.erro) {
      meta.textContent = f.erro;
    } else {
      const taxas = f.registos.reduce((a, r) => a + (r.totalTaxas || 0), 0);
      const marcas = estado.fonte === 'pdf'
        ? [
          `Período: ${f.periodo || '⚠ não identificado'}`,
          `Data: ${f.dataIso ? formatDatePt(f.dataIso) : '⚠ não identificada'}`
            + `${f.dataDeduzida ? ' (do conteúdo)' : ''}`,
          `${f.registos.length} registo(s)`,
          `Taxas: ${taxas.toLocaleString('pt-PT')}`,
        ]
        : [
          `${f.registos.length} registo(s)`,
          `${f.registos.filter((r) => r.tecnico).length} com técnico`,
          `${f.tecnicos.length} técnico(s) no mapa`,
          `Taxas: ${taxas.toLocaleString('pt-PT')}`,
        ];
      if (problemas.length) marcas.push(`⚠ ${problemas.length} por rever`);
      meta.textContent = marcas.join(' · ');
    }
    info.appendChild(meta);

    const acoes = document.createElement('div');
    acoes.className = 'ficheiro-acoes';

    if (problemas.length) {
      const rever = document.createElement('button');
      rever.type = 'button';
      rever.className = 'btn btn-aviso';
      rever.textContent = f.aRever ? 'Fechar' : `Corrigir ${problemas.length}`;
      rever.addEventListener('click', () => { f.aRever = !f.aRever; desenharFicheiros(); });
      acoes.appendChild(rever);
    }

    const remover = document.createElement('button');
    remover.type = 'button';
    remover.className = 'btn-remover';
    remover.title = 'Remover';
    remover.textContent = '✕';
    remover.addEventListener('click', () => {
      estado.ficheiros = estado.ficheiros.filter((x) => x.id !== f.id);
      desenharFicheiros();
    });
    acoes.appendChild(remover);

    div.append(info, acoes);
    bloco.appendChild(div);
    if (f.aRever && problemas.length) bloco.appendChild(painelCorreccao(f));
    lista.appendChild(bloco);
  }

  const temDados = estado.ficheiros.some((f) => f.registos && f.registos.length)
    || (estado.base && estado.base.registos && estado.base.registos.length);
  $('btn-gerar').disabled = !temDados;

  // Acrescenta à lista os técnicos que só aparecem nos dados.
  estado.tecnicosAcrescentados = sincronizarTecnicosDosDados();
  mostrarPrevisao();
}

/* ------------------------------------------------------------------ */
/* Erros de leitura — aviso e correcção manual pela interface          */
/* ------------------------------------------------------------------ */

/** Escapa texto para poder ser interpolado em innerHTML com segurança. */
function escaparHtml(txt) {
  return String(txt).replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

/** Uma linha está por rever quando falta o essencial ou o total é duvidoso. */
function temProblema(r) {
  return !!(
    r.problema
    || r.totalTaxas == null
    || r.totalTaxas <= 0
    || !r.dataReg
    || !String(r.numeroDU || '').trim()
  );
}

/** Reavalia o campo `problema` de um registo depois de uma edição manual. */
function reavaliarProblema(r) {
  if (!String(r.numeroDU || '').trim()) r.problema = 'Nº do DU em falta';
  else if (!r.dataReg) r.problema = 'Data de Reg. em falta';
  else if (r.totalTaxas == null || r.totalTaxas <= 0) r.problema = 'Total das Taxas por confirmar';
  else delete r.problema;
}

/** "01.07.2026", "1/7/2026" ou "2026-07-01" -> ISO; '' se não for data. */
function dataPtParaIso(texto) {
  const s = String(texto || '').trim();
  if (!s) return '';
  let m = /^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$/.exec(s);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  m = /^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})$/.exec(s);
  if (m) return `${m[3]}-${String(m[2]).padStart(2, '0')}-${String(m[1]).padStart(2, '0')}`;
  return '';
}

/**
 * Painel com as linhas por rever de um ficheiro, editáveis à mão. Cada campo
 * actualiza o registo em directo; ao sair do campo, a lista é redesenhada e as
 * linhas que ficaram completas saem do aviso.
 */
function painelCorreccao(f) {
  const painel = document.createElement('div');
  painel.className = 'painel-correccao';

  const nota = document.createElement('p');
  nota.className = 'painel-correccao-nota';
  nota.innerHTML = '⚠ <strong>Leituras por rever.</strong> Confirme os valores com o PDF '
    + 'original. Cada linha sai do aviso assim que ficar completa.';
  painel.appendChild(nota);

  const tabela = document.createElement('table');
  tabela.className = 'tabela-correccao';
  tabela.innerHTML = '<thead><tr>'
    + '<th>Nº do DU</th><th>Data (dd.mm.aaaa)</th><th>Total das Taxas</th>'
    + '<th>Período</th><th>Por rever</th></tr></thead>';
  const corpo = document.createElement('tbody');
  tabela.appendChild(corpo);

  const campo = (valor, aoEditar, extra = {}) => {
    const input = document.createElement('input');
    input.type = 'text';
    input.value = valor;
    if (extra.placeholder) input.placeholder = extra.placeholder;
    if (extra.titulo) input.title = extra.titulo;
    input.addEventListener('input', () => aoEditar(input));
    // Ao sair do campo, redesenha a lista de ficheiros (as linhas já completas
    // desaparecem do aviso e as contagens actualizam-se).
    input.addEventListener('change', () => desenharFicheiros());
    return input;
  };

  const desenhar = () => {
    corpo.innerHTML = '';
    const porRever = f.registos.filter(temProblema);
    for (const r of porRever) {
      const tr = document.createElement('tr');
      const tdProb = document.createElement('td');
      tdProb.className = 'coluna-problema';
      const marcarProblema = () => { tdProb.textContent = r.problema || ''; };

      const tdDU = document.createElement('td');
      tdDU.appendChild(campo(r.numeroDU || '', (input) => {
        r.numeroDU = input.value.trim();
        reavaliarProblema(r);
        marcarProblema();
      }, { placeholder: 'Nº do DU' }));

      const tdData = document.createElement('td');
      const inData = campo(r.dataReg ? formatDatePt(r.dataReg) : '', (input) => {
        const iso = dataPtParaIso(input.value);
        r.dataReg = iso || '';
        input.classList.toggle('invalido', !!input.value.trim() && !iso);
        reavaliarProblema(r);
        marcarProblema();
      }, { placeholder: 'dd.mm.aaaa' });
      tdData.appendChild(inData);

      const tdTotal = document.createElement('td');
      const inTotal = campo(r.totalTaxas != null ? String(r.totalTaxas) : '', (input) => {
        const n = parseAmount(input.value);
        r.totalTaxas = n;
        input.classList.toggle('invalido', !!input.value.trim() && n == null);
        reavaliarProblema(r);
        marcarProblema();
      }, {
        placeholder: '0',
        titulo: r.totalTaxasTexto ? `Lido do PDF: "${r.totalTaxasTexto}"` : '',
      });
      tdTotal.appendChild(inTotal);

      const tdPer = document.createElement('td');
      tdPer.textContent = r.periodo || f.periodo || '—';

      marcarProblema();
      tr.append(tdDU, tdData, tdTotal, tdPer, tdProb);
      corpo.appendChild(tr);
    }
  };

  desenhar();
  painel.appendChild(tabela);
  return painel;
}

/* ------------------------------------------------------------------ */
/* Técnicos                                                            */
/* ------------------------------------------------------------------ */

function adicionarLinhaTecnico(nome = '', tipo = '') {
  const tr = document.createElement('tr');

  const tdNome = document.createElement('td');
  const inNome = document.createElement('input');
  inNome.type = 'text';
  inNome.placeholder = 'Nome do técnico';
  inNome.value = nome;
  tdNome.appendChild(inNome);

  const tdTipo = document.createElement('td');
  const inTipo = document.createElement('input');
  inTipo.type = 'text';
  inTipo.placeholder = 'Ex.: A';
  inTipo.value = tipo;
  tdTipo.appendChild(inTipo);

  const tdAcao = document.createElement('td');
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'btn-remover';
  btn.title = 'Remover';
  btn.textContent = '✕';
  btn.addEventListener('click', () => { tr.remove(); guardarTecnicos(); actualizarContagem(); });
  tdAcao.appendChild(btn);

  for (const campo of [inNome, inTipo]) {
    campo.addEventListener('input', () => { guardarTecnicos(); actualizarContagem(); });
  }

  tr.append(tdNome, tdTipo, tdAcao);
  $('corpo-tecnicos').appendChild(tr);
  return tr;
}

function lerTecnicosDoFormulario() {
  return [...$('corpo-tecnicos').querySelectorAll('tr')]
    .map((tr) => {
      const [nome, tipo] = [...tr.querySelectorAll('input')].map((i) => i.value.trim());
      return { nome, tipo };
    })
    .filter((t) => t.nome);
}

function mostrarTecnicos(lista) {
  $('corpo-tecnicos').innerHTML = '';
  for (const t of lista) adicionarLinhaTecnico(t.nome, t.tipo || '');
  if (!lista.length) adicionarLinhaTecnico();
  actualizarContagem();
}

function actualizarContagem() {
  const lista = lerTecnicosDoFormulario();
  const comTipo = lista.filter((t) => t.tipo).length;
  $('contagem-tecnicos').textContent = lista.length
    ? `${lista.length} técnico(s) · ${comTipo} com tipo definido`
    : 'nenhum técnico definido';
}

/** Guarda a lista no navegador, para não ser preciso reescrevê-la. */
function guardarTecnicos() {
  try {
    localStorage.setItem(CHAVE_GUARDADA, JSON.stringify(lerTecnicosDoFormulario()));
  } catch { /* sem armazenamento disponível — segue sem guardar */ }
}

function tecnicosGuardados() {
  try {
    const bruto = localStorage.getItem(CHAVE_GUARDADA);
    const lista = bruto ? JSON.parse(bruto) : null;
    if (Array.isArray(lista) && lista.length) {
      return lista.filter((t) => t && typeof t.nome === 'string' && t.nome.trim());
    }
  } catch { /* valor inválido — usa a lista predefinida */ }
  return null;
}

$('btn-add-tecnico').addEventListener('click', () => {
  adicionarLinhaTecnico().querySelector('input').focus();
  actualizarContagem();
});
$('btn-repor-tecnicos').addEventListener('click', () => {
  mostrarTecnicos(TECNICOS_PREDEFINIDOS);
  guardarTecnicos();
});
$('btn-limpar-tecnicos').addEventListener('click', () => {
  mostrarTecnicos([]);
  guardarTecnicos();
});

mostrarTecnicos(tecnicosGuardados() || TECNICOS_PREDEFINIDOS);

/** Todos os ficheiros com dados carregados (consolidado base incluído). */
function ficheirosComDados() {
  return [estado.base, ...estado.ficheiros].filter((f) => f && f.registos?.length);
}

/**
 * Técnicos que aparecem na coluna Técnico dos dados mas ainda não estão na
 * lista. Devolve { nome, tipo } — o tipo vem do registo, se lá estiver.
 *
 * @param {Set<string>} jaConhecidos nomes já na lista (em minúsculas)
 */
function tecnicosNovosDosDados(jaConhecidos) {
  const novos = new Map();
  for (const f of ficheirosComDados()) {
    for (const r of f.registos) {
      const nome = (r.tecnico || '').trim();
      if (!nome) continue;
      const chave = nome.toLowerCase();
      if (jaConhecidos.has(chave) || novos.has(chave)) continue;
      novos.set(chave, { nome, tipo: (r.tipo || '').trim() });
    }
  }
  return [...novos.values()];
}

/**
 * Acrescenta à lista visível os técnicos que aparecem nos dados e ainda lá não
 * estão. Devolve os nomes acrescentados (para a interface poder avisar).
 */
function sincronizarTecnicosDosDados() {
  const atuais = new Set(lerTecnicosDoFormulario().map((t) => t.nome.toLowerCase()));
  const novos = tecnicosNovosDosDados(atuais);
  if (!novos.length) return [];
  for (const t of novos) adicionarLinhaTecnico(t.nome, t.tipo || '');
  guardarTecnicos();
  actualizarContagem();
  return novos.map((t) => t.nome);
}

/**
 * Técnicos a escrever nas tabelas de cada Excel: os do formulário, mais os dos
 * relatórios carregados, mais os que só aparecem nos dados (acrescentados
 * automaticamente, para nenhum processo ficar com um técnico fora da lista).
 */
function tecnicosParaOsFicheiros() {
  const juntos = [...lerTecnicosDoFormulario()];
  const vistos = new Set(juntos.map((t) => t.nome.toLowerCase()));
  const acrescentar = (lista) => {
    for (const t of lista) {
      const chave = t.nome.toLowerCase();
      if (vistos.has(chave)) continue;
      vistos.add(chave);
      juntos.push(t);
    }
  };
  const comMapa = ficheirosComDados().filter((f) => f.tecnicos?.length);
  acrescentar(juntarTecnicos(comMapa));
  acrescentar(tecnicosNovosDosDados(vistos));
  return juntos;
}

/* ------------------------------------------------------------------ */
/* Grupos                                                              */
/* ------------------------------------------------------------------ */

function ficheirosValidos() {
  return estado.ficheiros.filter((f) => f.registos && f.registos.length);
}

/** Grupos por dias consecutivos; vazio quando só se quer o consolidado. */
function calcularGrupos() {
  const modo = seleccionado('modo');
  if (modo === 'consolidado') return [];
  const dias = Number(modo) || 3;
  return agruparPorDiasConsecutivos(ficheirosValidos(), dias).map((g) => ({
    ...g,
    rotulo: rotuloDoGrupo(g),
    nomeFicheiro: nomeDoGrupo(g),
  }));
}

function querConsolidado() {
  return juntaRelatorios() || seleccionado('modo') === 'consolidado' || $('gerar-consolidado').checked;
}

/** Todos os relatórios em jogo, com o consolidado base à cabeça. */
function fontesParaJuntar() {
  const lista = [];
  if (estado.base && estado.base.registos?.length) {
    lista.push({ origem: estado.base.nome, registos: estado.base.registos, ehBase: true });
  }
  for (const f of ficheirosValidos()) lista.push({ origem: f.nome, registos: f.registos });
  return lista;
}

/** Junta tudo e resolve as repetições segundo a política escolhida. */
function juntarTudo() {
  return juntarRegistos(fontesParaJuntar(), seleccionado('politica'));
}

function nomeDoConsolidado() {
  return estado.fonte === 'acrescentar' ? 'Relatorio_Consolidado_Actualizado' : 'Relatorio_Consolidado';
}

function mostrarPrevisao() {
  const alvo = $('mensagens');
  alvo.innerHTML = '';
  const validos = ficheirosValidos();
  const temBase = !!(estado.base && estado.base.registos?.length);
  if (!validos.length && !temBase) return;

  const caixas = [];

  if (juntaRelatorios()) {
    const fontes = fontesParaJuntar();
    const { registos, duplicados } = juntarTudo();
    const somaBruta = fontes.reduce((a, f) => a + f.registos.length, 0);
    const diferentes = duplicados.filter((d) => !d.iguais).length;

    const caixa = document.createElement('div');
    caixa.className = 'aviso aviso-info';
    caixa.innerHTML = `<p>Será gerado <strong>${nomeDoConsolidado()}.xlsx</strong> com `
      + `<strong>${registos.length}</strong> processo(s), a partir de ${fontes.length} `
      + `relatório(s) (${somaBruta} linhas ao todo).</p>`
      + (duplicados.length
        ? `<ul><li><strong>${duplicados.length}</strong> repetição(ões) encontrada(s) — `
          + `${duplicados.length - diferentes} com dados iguais, `
          + `<strong>${diferentes}</strong> com dados diferentes.</li>`
          + `<li>Em cada uma fica ${POLITICAS[seleccionado('politica')]}; `
          + 'as outras versões vão para a folha <em>Redundâncias</em>.</li></ul>'
        : '<ul><li>Nenhuma repetição encontrada.</li></ul>');
    caixas.push(caixa);
  } else {
    const grupos = calcularGrupos();
    const linhas = grupos.map((g) => {
      const n = g.ficheiros.reduce((a, f) => a + f.registos.length, 0);
      return `<li><strong>${g.nomeFicheiro}.xlsx</strong> — ${g.rotulo} · `
        + `${g.ficheiros.length} ficheiro(s) · ${n} registo(s)</li>`;
    });
    if (querConsolidado()) {
      const n = validos.reduce((a, f) => a + f.registos.length, 0);
      linhas.push(`<li><strong>Relatorio_Consolidado.xlsx</strong> — tudo · ${n} registo(s)</li>`);
    }
    if (!linhas.length) return;
    const caixa = document.createElement('div');
    caixa.className = 'aviso aviso-info';
    caixa.innerHTML = `<p>Serão gerados <strong>${linhas.length}</strong> ficheiro(s) Excel:</p>`
      + `<ul>${linhas.join('')}</ul>`;
    caixas.push(caixa);
  }

  // Ficheiros com leituras por rever — a corrigir antes de gerar.
  const comProblemas = [...validos, ...(temBase ? [estado.base] : [])]
    .map((f) => ({ f, n: (f.registos || []).filter(temProblema).length }))
    .filter((x) => x.n);
  if (comProblemas.length) {
    const total = comProblemas.reduce((a, x) => a + x.n, 0);
    const caixa = document.createElement('div');
    caixa.className = 'aviso aviso-atencao';
    caixa.innerHTML = `<p>⚠ <strong>${total}</strong> leitura(s) por rever em `
      + `<strong>${comProblemas.length}</strong> ficheiro(s). Use o botão `
      + '<em>Corrigir</em> em cada ficheiro assinalado para acertar os valores antes de gerar.</p>';
    caixas.unshift(caixa);
  }

  // Técnicos que apareceram nos dados e foram acrescentados à lista.
  const acrescentados = estado.tecnicosAcrescentados || [];
  if (acrescentados.length) {
    const caixa = document.createElement('div');
    caixa.className = 'aviso aviso-ok';
    caixa.innerHTML = `<p>➕ <strong>${acrescentados.length}</strong> técnico(s) `
      + 'dos dados foram acrescentados à lista automaticamente: '
      + `${acrescentados.map(escaparHtml).join(', ')}.</p>`;
    caixas.unshift(caixa);
  }

  for (const c of caixas) alvo.appendChild(c);
}

for (const radio of document.querySelectorAll('input[name="modo"]')) {
  radio.addEventListener('change', () => {
    const so = seleccionado('modo') === 'consolidado';
    $('gerar-consolidado').checked = so || $('gerar-consolidado').checked;
    $('gerar-consolidado').disabled = so;
    $('campo-consolidado').classList.toggle('desactivada', so);
    mostrarPrevisao();
  });
}
$('gerar-consolidado').addEventListener('change', mostrarPrevisao);
for (const radio of document.querySelectorAll('input[name="politica"]')) {
  radio.addEventListener('change', mostrarPrevisao);
}

/* ------------------------------------------------------------------ */
/* Geração                                                             */
/* ------------------------------------------------------------------ */

$('btn-gerar').addEventListener('click', gerar);

function progresso(feito, total, texto) {
  $('progresso').hidden = false;
  $('barra-interior').style.width = `${total ? (feito / total) * 100 : 0}%`;
  $('progresso-texto').textContent = texto;
}

async function gerar() {
  const botao = $('btn-gerar');
  botao.disabled = true;
  $('resultados').innerHTML = '';
  $('btn-zip').hidden = true;
  estado.gerados = [];

  const opcoes = {
    tecnicos: tecnicosParaOsFicheiros(),
    estado: $('estado').value.trim() || 'Pago',
    linhasTecnicos: Math.max(2, Number($('linhas-tecnicos').value) || 8),
  };

  try {
    const grupos = calcularGrupos();
    const consolidar = querConsolidado();
    const total = grupos.length + (consolidar ? 1 : 0);
    let feito = 0;

    for (const g of grupos) {
      progresso(feito, total, `A gerar ${g.nomeFicheiro}.xlsx…`);
      await new Promise((r) => setTimeout(r, 0));
      const buffer = await criarWorkbook(window.ExcelJS, g, opcoes);
      estado.gerados.push({ nome: `${g.nomeFicheiro}.xlsx`, blob: paraBlob(buffer), grupo: g });
      feito++;
    }

    if (consolidar) {
      const nome = `${nomeDoConsolidado()}.xlsx`;
      progresso(feito, total, `A gerar ${nome}…`);
      await new Promise((r) => setTimeout(r, 0));

      const usados = juntaRelatorios() && estado.base
        ? [estado.base, ...ficheirosValidos()]
        : ficheirosValidos();
      const todos = { rotulo: 'Consolidado', ficheiros: usados };

      // Ao juntar relatórios, as linhas passam pela verificação de redundâncias.
      const extra = {};
      if (juntaRelatorios()) {
        const politica = seleccionado('politica');
        const { registos, duplicados } = juntarTudo();
        extra.registos = registos;
        extra.duplicados = duplicados;
        extra.politica = POLITICAS[politica];
      }

      const buffer = await criarWorkbook(window.ExcelJS, todos,
        { ...opcoes, ...extra, consolidado: true });
      estado.gerados.push({ nome, blob: paraBlob(buffer), grupo: todos, extra });
      feito++;
    }

    progresso(total, total, `${total} ficheiro(s) gerado(s).`);
    desenharResultados();
  } catch (err) {
    console.error(err);
    const box = document.createElement('div');
    box.className = 'aviso aviso-erro';
    box.textContent = `Erro ao gerar os ficheiros: ${err.message}`;
    $('resultados').appendChild(box);
  } finally {
    botao.disabled = false;
  }
}

function paraBlob(buffer) {
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

function desenharResultados() {
  const alvo = $('resultados');
  alvo.innerHTML = '';

  for (const g of estado.gerados) {
    const linha = document.createElement('div');
    linha.className = 'resultado';

    const info = document.createElement('div');
    const nRegistos = g.extra?.registos
      ? g.extra.registos.length
      : g.grupo.ficheiros.reduce((a, f) => a + f.registos.length, 0);
    const repetidos = g.extra?.duplicados?.length
      ? ` · ${g.extra.duplicados.length} repetição(ões) resolvida(s)` : '';
    info.innerHTML = `<div class="resultado-nome">📗 ${g.nome}</div>`
      + `<div class="ficheiro-meta">${g.grupo.ficheiros.length} ficheiro(s) · `
      + `${nRegistos} registo(s)${repetidos}</div>`;

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn btn-principal';
    btn.textContent = '⬇ Descarregar';
    btn.addEventListener('click', () => descarregar(g.blob, g.nome));

    linha.append(info, btn);
    alvo.appendChild(linha);
  }

  $('btn-zip').hidden = estado.gerados.length < 2;
}

function descarregar(blob, nome) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = nome;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

$('btn-zip').addEventListener('click', async () => {
  const zip = new window.JSZip();
  for (const g of estado.gerados) zip.file(g.nome, g.blob);
  const blob = await zip.generateAsync({ type: 'blob' });
  descarregar(blob, 'Relatorios.zip');
});

aplicarFonte();
