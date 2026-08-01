/*
 * Versão de linha de comandos: gera os relatórios a partir de um ficheiro de
 * registos (o mesmo JSON que a aplicação exporta), sem browser.
 *
 *   node tools/gerar-cli.mjs <registos.json> [pasta-de-saida] [YYYY-MM]
 *
 * Sem argumentos, corre com registos de exemplo e escreve numa pasta temporária.
 */
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import fs from 'fs';
import path from 'path';

const require = createRequire(import.meta.url);
const raiz = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const JSZip = (await import('jszip')).default;
require(path.join(raiz, 'assets', 'docx.js'));
require(path.join(raiz, 'assets', 'dados.js'));
require(path.join(raiz, 'assets', 'registos.js'));
require(path.join(raiz, 'assets', 'agregacao.js'));
require(path.join(raiz, 'assets', 'modelos.js'));
const { docx, dados, agregacao, modelos } = globalThis.PMR;

const args = process.argv.slice(2);
const fonte = args[0];
const saida = args[1] || path.join(raiz, 'saida');
const mes = args[2];

let registos;
if (fonte && fs.existsSync(fonte)) {
  registos = JSON.parse(fs.readFileSync(fonte, 'utf8'));
} else {
  registos = [
    { data: '2026-04-13', turno: 'B', lado: 'Terra', chefeTurno: 'Jessé Martins',
      voos: 8, passageiros: 600, separados: { comercial: 20, naoComercial: 0, regPosterior: 0 },
      apreensoes: [{ tipo: 'Moeda/Valores', passageiro: 'Manuel Dalamana', nacionalidade: 'Angolana',
        origemDestino: 'Addis Abeba', voo: 'ET850', valor: '15.050,00 USD' }], bagagens: [] },
    { data: '2026-04-14', turno: 'B', lado: 'Terra', chefeTurno: 'Jessé Martins',
      voos: 6, passageiros: 495, separados: { comercial: 12, naoComercial: 0, regPosterior: 0 },
      apreensoes: [], bagagens: [] }
  ];
  console.log('(sem ficheiro de registos — a usar exemplo)');
}

const inst = dados.carregar().instituicao;
const equipa = dados.carregar().equipa;
const meses = mes ? [mes] : [...new Set(registos.map((r) => r.data.slice(0, 7)))];
const dias = [...new Set(registos.map((r) => r.data))];

fs.mkdirSync(saida, { recursive: true });
async function escrever(doc) {
  const buf = await docx.empacotar(JSZip, doc.xml).generateAsync({ type: 'nodebuffer' });
  const f = path.join(saida, doc.nome + '.docx');
  fs.writeFileSync(f, buf);
  console.log('  ' + path.relative(process.cwd(), f));
}

console.log('A gerar em', saida);
for (const dia of dias) await escrever(modelos.relatorio24(inst, agregacao.rga24(registos, dia)));
for (const m of meses) {
  await escrever(modelos.relatorioTerra(inst, agregacao.rmt(registos, m)));
  await escrever(modelos.relatorioAr(inst, equipa, agregacao.rma(registos, m)));
}
console.log('Concluído.');
