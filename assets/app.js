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
} from './parser.js';
import { criarWorkbook } from './report.js';
import { lerRelatorio, juntarTecnicos } from './leitor-relatorio.js';

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
  fonte: 'pdf',    // 'pdf' | 'relatorio'
  ficheiros: [],   // { id, nome, periodo, dataIso, file, registos, tecnicos, erro }
  gerados: [],     // { nome, blob, grupo }
};

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

  inputFicheiros.accept = ehPdf ? 'application/pdf,.pdf' : '.xlsx';
  $('zona-nota').innerHTML = ehPdf
    ? 'O nome deve conter a data e o período — ex.: <code>Manhã 21.07.2026.pdf</code>, '
      + '<code>01.07.2026_Tarde.pdf</code>, <code>2026-07-23 Madrugada.pdf</code>'
    : 'Relatórios <code>.xlsx</code> gerados por esta aplicação. Os técnicos e os estados que '
      + 'lá estiverem preenchidos são mantidos.';

  // A partir de relatórios só faz sentido consolidar.
  for (const radio of document.querySelectorAll('input[name="modo"]')) {
    radio.disabled = !ehPdf;
    if (!ehPdf) radio.checked = radio.value === 'consolidado';
  }
  document.querySelectorAll('input[name="modo"]')[0].closest('.opcoes')
    .classList.toggle('desactivada', !ehPdf);

  estado.ficheiros = [];
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
    const div = document.createElement('div');
    div.className = 'ficheiro' + (f.erro ? ' com-erro' : '');

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
      meta.textContent = marcas.join(' · ');
    }
    info.appendChild(meta);

    const remover = document.createElement('button');
    remover.type = 'button';
    remover.className = 'btn-remover';
    remover.title = 'Remover';
    remover.textContent = '✕';
    remover.addEventListener('click', () => {
      estado.ficheiros = estado.ficheiros.filter((x) => x.id !== f.id);
      desenharFicheiros();
    });

    div.appendChild(info);
    div.appendChild(remover);
    lista.appendChild(div);
  }

  $('btn-gerar').disabled = !estado.ficheiros.some((f) => f.registos && f.registos.length);
  mostrarPrevisao();
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

/** Técnicos do formulário mais os que vierem dos relatórios carregados. */
function tecnicosParaOsFicheiros() {
  const doFormulario = lerTecnicosDoFormulario();
  const dosRelatorios = juntarTecnicos(estado.ficheiros.filter((f) => f.tecnicos?.length));
  const juntos = [...doFormulario];
  const vistos = new Set(doFormulario.map((t) => t.nome.toLowerCase()));
  for (const t of dosRelatorios) {
    if (vistos.has(t.nome.toLowerCase())) continue;
    vistos.add(t.nome.toLowerCase());
    juntos.push(t);
  }
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
  return seleccionado('modo') === 'consolidado' || $('gerar-consolidado').checked;
}

function mostrarPrevisao() {
  const alvo = $('mensagens');
  alvo.innerHTML = '';
  const validos = ficheirosValidos();
  if (!validos.length) return;

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

  const box = document.createElement('div');
  box.className = 'aviso aviso-info';
  box.innerHTML = `<p>Serão gerados <strong>${linhas.length}</strong> ficheiro(s) Excel:</p>`
    + `<ul>${linhas.join('')}</ul>`;
  alvo.appendChild(box);
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
    linhasTecnicos: Math.max(2, Number($('linhas-tecnicos').value) || 12),
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
      progresso(feito, total, 'A gerar Relatorio_Consolidado.xlsx…');
      await new Promise((r) => setTimeout(r, 0));
      const todos = { rotulo: 'Consolidado', ficheiros: ficheirosValidos() };
      const buffer = await criarWorkbook(window.ExcelJS, todos, { ...opcoes, consolidado: true });
      estado.gerados.push({ nome: 'Relatorio_Consolidado.xlsx', blob: paraBlob(buffer), grupo: todos });
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
    const nRegistos = g.grupo.ficheiros.reduce((a, f) => a + f.registos.length, 0);
    info.innerHTML = `<div class="resultado-nome">📗 ${g.nome}</div>`
      + `<div class="ficheiro-meta">${g.grupo.ficheiros.length} ficheiro(s) · ${nRegistos} registo(s)</div>`;

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
