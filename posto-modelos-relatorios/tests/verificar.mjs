/*
 * Testes de fidelidade. Geram cada documento a partir de dados conhecidos,
 * reabrem o .docx (é um ZIP), e confirmam que:
 *   - o pacote é um Word válido (tem as peças obrigatórias);
 *   - o texto introduzido aparece no document.xml;
 *   - as tabelas têm o número de linhas esperado;
 *   - a agregação mensal é igual à soma dos turnos.
 * Corre sem browser: node tests/verificar.mjs
 */
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import path from 'path';

const require = createRequire(import.meta.url);
const raiz = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
// No navegador o JSZip vem do ficheiro vendorizado (global); nos testes usamos
// o pacote npm, que é o mesmo JSZip 3.10.1.
const JSZip = (await import('jszip')).default;

// Carrega os módulos (UMD: registam-se em globalThis.PMR e devolvem-se).
require(path.join(raiz, 'assets', 'docx.js'));
require(path.join(raiz, 'assets', 'dados.js'));
require(path.join(raiz, 'assets', 'registos.js'));
require(path.join(raiz, 'assets', 'agregacao.js'));
require(path.join(raiz, 'assets', 'modelos.js'));
const { docx, dados, registos, agregacao, modelos } = globalThis.PMR;

let passou = 0, falhou = 0;
function ok(cond, msg) {
  if (cond) { passou++; } else { falhou++; console.error('  ✗ ' + msg); }
}

async function pecas(xml) {
  const zip = await docx.empacotar(JSZip, xml).generateAsync({ type: 'nodebuffer' });
  const lido = await JSZip.loadAsync(zip);
  return {
    tem: (p) => lido.file(p) != null,
    doc: await lido.file('word/document.xml').async('string')
  };
}

const inst = dados.predefinicao().instituicao;

async function testarTRB() {
  console.log('TRB — Termo de Recepção');
  const doc = modelos.termoRecepcao(inst, {
    chefeAgt: 'Fulano AGT', chefePfa: 'Beltrano PFA', companhia: 'TAAG',
    dataTermo: '2026-07-06',
    bagagens: [{ proveniencia: 'Johannesburg', voo: 'DT-578', dataVoo: '2026-07-05',
      passageiro: 'Kiaku Nelson', etiqueta: '026851' }]
  });
  const z = await pecas(doc.xml);
  ok(z.tem('[Content_Types].xml') && z.tem('word/document.xml') && z.tem('word/styles.xml'),
    'pacote .docx válido');
  ok(z.doc.includes('Kiaku Nelson'), 'nome do passageiro presente');
  ok(z.doc.includes('026851'), 'n.º de etiqueta presente');
  ok(z.doc.includes('DT-578'), 'n.º de voo presente');
  ok(z.doc.includes('06 de Julho de 2026'), 'data por extenso presente');
}

async function testarCKL() {
  console.log('CKL — Check-list');
  const its = dados.predefinicao().itensCkl.map((i) => Object.assign({ estado: 'Operacional', observacao: '' }, i));
  const doc = modelos.checklist(inst, {
    turno: 'B', data: '2026-02-02', itens: its,
    ocorrencia: 'Muitos mosquitos no piquete.',
    chefeEntrega: 'Entrega X', chefeRecebe: 'Recebe Y'
  });
  const z = await pecas(doc.xml);
  ok(z.doc.includes('CHECK-LIST'), 'título presente');
  ok(z.doc.includes('02.02.2026'), 'data curta presente');
  ok(z.doc.includes('Muitos mosquitos no piquete.'), 'ocorrência presente');
  const nLinhasTabela = (z.doc.match(/<w:tr>/g) || []).length;
  ok(nLinhasTabela === its.length + 1, `tabela com ${its.length} itens + cabeçalho (tem ${nLinhasTabela})`);
}

function registosExemplo() {
  // 2 turnos no lado Terra + 1 no lado Ar, no mesmo mês; um deles com apreensões.
  const r = [];
  const a = registos.novoRegisto();
  Object.assign(a, {
    id: 'a', data: '2026-04-13', turno: 'B', lado: 'Terra', chefeTurno: 'Jessé Martins',
    voos: 8, passageiros: 600,
    separados: { comercial: 20, naoComercial: 0, regPosterior: 0 },
    apreensoes: [{ tipo: 'Moeda/Valores', passageiro: 'Manuel Dalamana',
      nacionalidade: 'Angolana', origemDestino: 'Addis Abeba', voo: 'ET850',
      valor: '15.050,00 USD', realce: false }]
  });
  const b = registos.novoRegisto();
  Object.assign(b, {
    id: 'b', data: '2026-04-14', turno: 'B', lado: 'Terra', chefeTurno: 'Jessé Martins',
    voos: 6, passageiros: 495,
    separados: { comercial: 12, naoComercial: 0, regPosterior: 0 }
  });
  const c = registos.novoRegisto();
  Object.assign(c, {
    id: 'c', data: '2026-04-20', turno: 'E', lado: 'Ar', chefeTurno: 'Abel Júlio',
    voos: 81, passageiros: 11542, volumes: 2, visitasCatering: 0, visitasBordo: 0,
    furgoneta: { scaneados: 0, inspecionados: 0, apreensoes: 0 }
  });
  r.push(a, b, c);
  return r;
}

async function testarRGA24() {
  console.log('RGA-24 — Relatório 24 h (agregado)');
  const regs = registosExemplo();
  const ag = agregacao.rga24(regs, '2026-04-13');
  ok(ag.valores.length === 1, 'uma apreensão de valores no dia 13');
  ok(ag.diaSeguinte === '2026-04-14', 'dia seguinte calculado');
  const doc = modelos.relatorio24(inst, ag);
  const z = await pecas(doc.xml);
  ok(z.doc.includes('Manuel Dalamana'), 'passageiro da apreensão presente');
  ok(z.doc.includes('15.050,00 USD'), 'valor presente');
  ok(z.doc.includes('07h00'), 'período 07h00 presente');
  ok(z.doc.includes('13 de Abril'), 'dia por extenso presente');
}

async function testarRMT() {
  console.log('RMT — Mensal Terra (agregado)');
  const regs = registosExemplo();
  const ag = agregacao.rmt(regs, '2026-04');
  ok(ag.voos === 14, `voos somados = 14 (tem ${ag.voos})`);
  ok(ag.passageiros === 1095, `passageiros somados = 1095 (tem ${ag.passageiros})`);
  ok(ag.separados.comercial === 32, `separados comerciais = 32 (tem ${ag.separados.comercial})`);
  ok(ag.apreensoes.moeda === 1, 'uma apreensão de moeda no mês');
  const doc = modelos.relatorioTerra(inst, ag);
  const z = await pecas(doc.xml);
  ok(z.doc.includes('1.095'), 'passageiros com separador de milhares');
  ok(z.doc.includes('ABRIL DE 2026'), 'mês por extenso no título');
  ok(z.doc.includes('TOTAL GERAL'), 'linha de total presente');
}

async function testarRMA() {
  console.log('RMA — Mensal Ar (agregado)');
  const regs = registosExemplo();
  const ag = agregacao.rma(regs, '2026-04');
  ok(ag.voos === 81, `voos processados = 81 (tem ${ag.voos})`);
  ok(ag.passageiros === 11542, `passageiros = 11542 (tem ${ag.passageiros})`);
  const equipa = dados.predefinicao().equipa;
  const doc = modelos.relatorioAr(inst, equipa, ag);
  const z = await pecas(doc.xml);
  ok(z.doc.includes('11.542'), 'passageiros com separador de milhares');
  ok(z.doc.includes('Grupo Scanner'), 'cabeçalho do grupo presente');
  ok(z.doc.includes('Abel Júlio'), 'equipa presente');
  ok(z.doc.includes('DADOS DA FURGONETA'), 'secção da furgoneta presente');
}

async function testarReconciliacao() {
  console.log('Reconciliação diário × mensal');
  const regs = registosExemplo();
  const somaDiaria = ['2026-04-13', '2026-04-14'].reduce((tot, dia) =>
    tot + agregacao.rga24(regs, dia).valores.length, 0);
  const mensal = agregacao.rmt(regs, '2026-04').apreensoes.moeda;
  ok(somaDiaria === mensal, `apreensões de valores: soma dos RGA-24 (${somaDiaria}) = RMT (${mensal})`);
}

const main = async () => {
  for (const t of [testarTRB, testarCKL, testarRGA24, testarRMT, testarRMA, testarReconciliacao]) {
    await t();
  }
  console.log(`\n${passou} verificações passaram, ${falhou} falharam.`);
  process.exit(falhou ? 1 : 0);
};
main();
