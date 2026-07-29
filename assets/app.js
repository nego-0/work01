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

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdf.worker.min.mjs', import.meta.url).href;

const $ = (id) => document.getElementById(id);

const estado = {
  ficheiros: [],   // { id, nome, periodo, dataIso, file, registos, erro }
  gerados: [],     // { nome, blob }
};

/* ------------------------------------------------------------------ */
/* Ficheiros                                                           */
/* ------------------------------------------------------------------ */

const zona = $('zona');
const inputFicheiros = $('input-ficheiros');

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
  const fs = [...(e.dataTransfer?.files || [])].filter((f) => /\.pdf$/i.test(f.name));
  adicionarFicheiros(fs);
});

$('btn-limpar').addEventListener('click', () => {
  estado.ficheiros = [];
  desenharFicheiros();
});

let contador = 0;

async function adicionarFicheiros(files) {
  for (const file of files) {
    if (!/\.pdf$/i.test(file.name)) continue;
    if (estado.ficheiros.some((f) => f.nome === file.name && f.file.size === file.size)) continue;

    const { periodo, dataIso } = parseFileName(file.name);
    const registo = {
      id: `f${++contador}`,
      nome: file.name,
      periodo,
      dataIso,
      file,
      registos: null,
      erro: null,
      aLer: true,
    };
    estado.ficheiros.push(registo);
    desenharFicheiros();
    await lerPdf(registo);
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
      // Sem data no nome: usa a data mais frequente dentro do PDF.
      const contagem = {};
      for (const r of item.registos) contagem[r.dataReg] = (contagem[r.dataReg] || 0) + 1;
      item.dataIso = Object.entries(contagem).sort((a, b) => b[1] - a[1])[0][0];
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
      const marcas = [
        `Período: ${f.periodo || '⚠ não identificado'}`,
        `Data: ${f.dataIso ? formatDatePt(f.dataIso) : '⚠ não identificada'}${f.dataDeduzida ? ' (do conteúdo)' : ''}`,
        `${f.registos.length} registo(s)`,
        `Taxas: ${f.registos.reduce((a, r) => a + (r.totalTaxas || 0), 0).toLocaleString('pt-PT')}`,
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
  mostrarPrevisaoGrupos();
}

/* ------------------------------------------------------------------ */
/* Técnicos                                                            */
/* ------------------------------------------------------------------ */

function adicionarLinhaTecnico(nome = '', tipo = '') {
  const tr = document.createElement('tr');

  const tdNome = document.createElement('td');
  const inNome = document.createElement('input');
  inNome.type = 'text';
  inNome.placeholder = 'Ex.: João';
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
  btn.textContent = '✕';
  btn.addEventListener('click', () => tr.remove());
  tdAcao.appendChild(btn);

  tr.append(tdNome, tdTipo, tdAcao);
  $('corpo-tecnicos').appendChild(tr);
}

$('btn-add-tecnico').addEventListener('click', () => adicionarLinhaTecnico());
adicionarLinhaTecnico();
adicionarLinhaTecnico();

function lerTecnicos() {
  return [...$('corpo-tecnicos').querySelectorAll('tr')]
    .map((tr) => {
      const [nome, tipo] = [...tr.querySelectorAll('input')].map((i) => i.value.trim());
      return { nome, tipo };
    })
    .filter((t) => t.nome);
}

/* ------------------------------------------------------------------ */
/* Grupos                                                              */
/* ------------------------------------------------------------------ */

function calcularGrupos() {
  const validos = estado.ficheiros.filter((f) => f.registos && f.registos.length);
  const dias = Math.max(1, Number($('dias-grupo').value) || 3);
  return agruparPorDiasConsecutivos(validos, dias).map((g) => ({
    ...g,
    rotulo: rotuloDoGrupo(g),
    nomeFicheiro: nomeDoGrupo(g),
  }));
}

function mostrarPrevisaoGrupos() {
  const alvo = $('mensagens');
  alvo.innerHTML = '';
  const grupos = calcularGrupos();
  if (!grupos.length) return;

  const box = document.createElement('div');
  box.className = 'aviso aviso-info';
  const total = grupos.length + ($('gerar-consolidado').checked ? 1 : 0);
  const linhas = grupos.map((g) => {
    const n = g.ficheiros.reduce((a, f) => a + f.registos.length, 0);
    return `<li><strong>${g.nomeFicheiro}.xlsx</strong> — ${g.rotulo} · ${g.ficheiros.length} PDF(s) · ${n} registo(s)</li>`;
  });
  if ($('gerar-consolidado').checked) {
    const n = grupos.reduce((a, g) => a + g.ficheiros.reduce((b, f) => b + f.registos.length, 0), 0);
    linhas.push(`<li><strong>Relatorio_Consolidado.xlsx</strong> — todos os períodos · ${n} registo(s)</li>`);
  }
  box.innerHTML = `<p>Serão gerados <strong>${total}</strong> ficheiro(s) Excel:</p><ul>${linhas.join('')}</ul>`;
  alvo.appendChild(box);
}

for (const id of ['dias-grupo', 'gerar-consolidado']) {
  $(id).addEventListener('change', mostrarPrevisaoGrupos);
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
    tecnicos: lerTecnicos(),
    estado: $('estado').value.trim() || 'Pago',
    linhasTecnicos: Math.max(1, Number($('linhas-tecnicos').value) || 12),
  };

  try {
    const grupos = calcularGrupos();
    const consolidar = $('gerar-consolidado').checked;
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
      const todos = {
        rotulo: 'Consolidado',
        ficheiros: grupos.flatMap((g) => g.ficheiros),
      };
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
      + `<div class="ficheiro-meta">${g.grupo.ficheiros.length} PDF(s) · ${nRegistos} registo(s)</div>`;

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
