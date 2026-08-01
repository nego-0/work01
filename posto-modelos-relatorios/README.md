# Posto Aduaneiro — Modelos e Relatórios

Aplicação web que **preenche os modelos de documentos do posto** e **produz os
relatórios de trabalho automaticamente**, a partir dos dados introduzidos uma só
vez por turno.

É **um único ficheiro HTML**: estilos, bibliotecas e código vão todos lá dentro.
Abre-se com dois cliques, sem servidor e sem Internet. **Os dados nunca saem do
computador** — não há nuvem, nem envio, nem telemetria. É um requisito, não um
extra: os documentos contêm nomes de passageiros, voos, valores e mercadorias
apreendidas.

Segue a estratégia em
[`docs/estrategia-sistema-modelos-relatorios.md`](../docs/estrategia-sistema-modelos-relatorios.md)
e a mesma arquitectura do gerador de relatórios Excel do posto.

## Os dois objectivos, um só cano

O sistema não trata «cinco formulários» soltos. Gira à volta de **um registo de
turno**, introduzido uma vez, de que cada documento é uma projecção:

- Uma **bagagem retida** → **TRB** (na hora).
- As **apreensões do dia** (soma dos turnos, 07h→07h) → **RGA-24** (diário).
- Os **turnos do mês**, agregados → **RMT** e **RMA** (mensais, calculados sozinhos).
- O **estado do material** no fim do turno → **CKL**.

Assim os números não divergem entre documentos: saem todos da mesma fonte. O total
de apreensões de um dia no RGA-24 é, por construção, uma parcela do total do mês
no RMT.

## Os cinco modelos

| Modelo | Nome | Entregue a | Como se gera |
|---|---|---|---|
| **CKL** | Check-list de materiais e equipamentos | Turno entrante | Separador *Check-list* (inventário já preenchido; só o estado e a ocorrência) |
| **TRB** | Termo de Recepção de Bagagens | Companhia aérea | Botão numa bagagem do registo, ou separador *Termo de Recepção* |
| **RGA-24** | Relatório-síntese das últimas 24 h | Gestão aeroportuária | Separador *Relatórios* → escolher o dia (agregado) |
| **RMT** | Relatório mensal — lado Terra | Gestão | Separador *Relatórios* → escolher o mês (agregado) |
| **RMA** | Relatório mensal — lado Ar (Scanner) | Gestão | Separador *Relatórios* → escolher o mês (agregado) |

## Abrir a aplicação

Guardar o `index.html` (botão direito → *Guardar como*) e abri-lo directamente,
ou servir a partir de qualquer servidor estático:

```bash
npm start          # http://127.0.0.1:8080
```

Os documentos saem em `.docx`, prontos a abrir no Word, imprimir ou entregar. Para
PDF, usar o «Guardar como PDF» do próprio Word.

## Como se usa

1. **Registos de Turno** — uma vez por turno, introduzir data, turno, lado
   (Terra/Ar), chefe, voos, passageiros, e o que houver: apreensões, bagagens
   retidas, separados, volumes, visitas, furgoneta. Guardar.
2. **Relatórios** — escolher o dia (RGA-24) ou o mês (RMT/RMA) e gerar. Sai
   agregado dos registos, sem reescrever números.
3. **Termo de Recepção** e **Check-list** — formulários directos para emissão
   avulsa.
4. **Dados-mestre** — a equipa, os chefes de turno, o inventário da check-list e o
   bloco institucional já vêm preenchidos e ficam guardados no navegador. Aqui
   também se exportam/importam os registos (cópia de segurança e passagem entre
   computadores).

## Construir o ficheiro único

O `index.html` é gerado a partir de `src/pagina.html`, `assets/` e `vendor/`:

```bash
npm run build      # reescreve o index.html
```

Depois de mexer em qualquer ficheiro de `assets/` ou `src/`, voltar a correr o
build antes de publicar.

## Linha de comandos

Gerar os relatórios sem browser, a partir de um ficheiro de registos (o mesmo JSON
que a aplicação exporta):

```bash
npm install
node tools/gerar-cli.mjs <registos.json> [pasta-de-saida] [YYYY-MM]
node tools/gerar-cli.mjs                 # corre com um exemplo
```

## Testes

```bash
npm install
npm test
```

Os testes geram cada documento a partir de dados conhecidos, reabrem o `.docx` (é
um ZIP), e confirmam que o pacote é um Word válido, que o texto introduzido
aparece, que as tabelas têm as linhas certas, e que **a agregação mensal é igual à
soma dos turnos** (o diário reconcilia com o mensal).

## Estrutura

```
index.html                 ficheiro único gerado (é o que se publica)
src/pagina.html            molde do HTML
assets/style.css           estilos
assets/docx.js             motor de geração de .docx (OOXML) sobre o JSZip
assets/dados.js            dados-mestre + persistência local
assets/registos.js         registos de turno (fonte única) + persistência
assets/agregacao.js        agrega os turnos nos relatórios (diário e mensal)
assets/modelos.js          os cinco modelos → estrutura do documento
assets/app.js              interface
vendor/jszip.min.js        JSZip (sem CDN)
tools/construir.mjs        junta tudo no ficheiro único
tools/gerar-cli.mjs        geração por linha de comandos
tests/verificar.mjs        testes de fidelidade e de reconciliação
```

## Notas técnicas

- Um `.docx` é um ZIP de XML. O motor (`assets/docx.js`) constrói parágrafos e
  tabelas em OOXML e empacota-os com o **JSZip já vendorizado** — a mesma
  biblioteca, sem segunda dependência.
- Os módulos são UMD: registam-se em `globalThis.PMR` no navegador e podem ser
  importados no Node, o que permite correr os mesmos geradores nos testes e na
  linha de comandos.
- Nos testes, o JSZip vem do pacote npm (é o mesmo 3.10.1); no navegador vem do
  ficheiro vendorizado. A aplicação nunca depende da rede.
