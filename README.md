# Gerador de Relatórios Excel

Aplicação web que lê relatórios PDF do tipo **Document List** (ASYCUDAWorld) e gera
ficheiros Excel prontos a trabalhar.

É **um único ficheiro HTML**: estilos, bibliotecas e código vão todos lá dentro.
Pode ser usado a partir do endereço abaixo ou guardado no disco e aberto com dois
cliques, sem servidor e sem Internet. Os PDFs nunca saem do computador.

## Abrir a aplicação

Endereço final: **https://nego-0.github.io/work01/**

O site é publicado pelo workflow
[`.github/workflows/static.yml`](.github/workflows/static.yml) a cada `push` para o
ramo de trabalho, e pode ser lançado à mão em *Actions → Deploy static content to
Pages → Run workflow*. Requer *Settings → Pages → Source:* **GitHub Actions**.

Para usar sem Internet: guardar o `index.html` (botão direito → *Guardar como*) e
abri-lo directamente. Também serve a partir de qualquer servidor estático:

```bash
npm start          # http://127.0.0.1:8080
```

### Construir o ficheiro único

O `index.html` é gerado a partir de `src/pagina.html`, `assets/` e `vendor/`:

```bash
npm run build      # reescreve o index.html (~2,9 MB)
```

Depois de mexer em qualquer ficheiro de `assets/` ou `src/`, é preciso voltar a
correr o build antes de publicar.

## O que faz

Para cada PDF, a aplicação lê a tabela de processos e produz as 7 colunas pedidas,
por esta ordem:

| # | Coluna | Origem |
|---|---|---|
| 1 | **Período** | nome do ficheiro (Madrugada, Manhã, Tarde, Noite) |
| 2 | **Data de Reg.** | coluna *Data de Reg.* do PDF |
| 3 | **Total das Taxas** | coluna *Total das taxas* do PDF |
| 4 | **Estado** | preenchido com `Pago` (configurável); lista pendente Pago / Não Pago |
| 5 | **Nº do DU** | coluna *N° do DU* do PDF |
| 6 | **Técnico** | em branco, a preencher no Excel |
| 7 | **Tipo** | fórmula: resulta do técnico escolhido |
| — | **Ficheiro (PDF)** | coluna extra, a seguir às 7: de que PDF veio a linha |

**Uma linha por processo, por ordem ascendente de data.** Todos os processos de
todos os PDFs do bloco vão para a tabela, sem agregações — ordenados pela **Data
de Reg.** e, dentro do mesmo dia, pela ordem do documento original. Cada linha tem
a sua própria célula de **Técnico**, para se indicar quem foi o responsável por
aquele processo.

### Leituras por rever

Se um total das taxas não for reconhecido, ou faltar a data ou o Nº do DU, a linha
não é descartada: o ficheiro fica assinalado a laranja com um botão **Corrigir**,
que abre uma pequena tabela onde se acertam à mão o Nº do DU, a data e o total,
comparando com o PDF. A linha sai do aviso assim que ficar completa, e só depois se
gera o Excel.

### Nomes dos ficheiros PDF

O período e a data são lidos do nome. São aceites, entre outros:

```
Manhã 21.07.2026.pdf      01.07.2026_Tarde.pdf      2026-07-23 Madrugada.pdf
```

Se o nome não tiver data, é usada a data predominante dentro do próprio PDF.

## Técnicos

A aplicação já vem com a equipa preenchida — Abednego Agostinho, Raquel Machado,
Leopoldo Maiato, Ruth Contreiras, Gisela Antonio, Vania Chungo, Jeronimo dos
Santos, Mario Massanga, Altair Pereira, Marcolino da Silva, Carolina Costa,
Constancia Cortez, Felson Jorge, Alfredo Jose, Jesse Martins e Virgilio da
Conceição. Os nomes podem ser mudados, removidos ou acrescentados, e o **Tipo**
de cada um pode ficar por definir.

A lista fica guardada no navegador, por isso as alterações mantêm-se de uma
utilização para a outra; o botão *Repor lista predefinida* traz a original de
volta. Estes nomes vão para as **três tabelas de técnicos** de cada Excel (8 por
tabela por omissão) e alimentam a lista pendente da coluna Técnico. No Excel podem
inserir-se mais linhas em qualquer tabela sem perder a lógica.

Ao carregar dados — sobretudo ao juntar relatórios já preenchidos — os técnicos que
aparecem nos dados e ainda não estão na lista são **acrescentados automaticamente**,
para nenhum processo ficar com um responsável fora da lista. A interface avisa quais
foram acrescentados.

## Origem dos dados

| Origem | Para quê |
|---|---|
| **Ficheiros PDF** | o caso normal: extrair os processos dos *Document List* |
| **Relatórios Excel já preenchidos** | juntar relatórios que já levam os técnicos e os estados indicados, sem voltar aos PDFs e sem perder o que foi preenchido |
| **Acrescentar a um consolidado existente** | pegar num consolidado já gerado e juntar-lhe novos relatórios |

Nas duas últimas, os técnicos, os tipos e os estados são lidos de cada ficheiro e
mantidos; os mapas Técnico → Tipo dos vários relatórios são juntos num só. Na
terceira, o consolidado existente é escolhido à parte — é ele que serve de base, e
os outros relatórios juntam-se-lhe.

### Repetições

Ao juntar relatórios, o mesmo processo pode vir em mais do que um ficheiro. Dois
processos são o mesmo quando têm o **mesmo Nº do DU na mesma Data de Reg.** —
cada um entra uma só vez, na posição em que apareceu primeiro.

Qual das versões fica é escolhido antes de gerar:

| Política | Fica |
|---|---|
| **A versão mais preenchida** (por omissão) | a que tiver técnico, tipo e estado indicados; em caso de empate, a que já estava |
| **A que já estava** | o consolidado existente (ou o primeiro relatório carregado) |
| **A que está a ser acrescentada** | os relatórios novos substituem o que lá estava |

Antes de gerar, a aplicação mostra quantas repetições encontrou e quantas trazem
dados diferentes. O ficheiro sai com uma folha **Redundâncias** que lista, uma a
uma, o Nº do DU, a data, se os dados coincidiam, e as duas versões lado a lado —
a mantida e a posta de lado, cada uma com o relatório de onde veio, o técnico, o
estado e o total.

O ficheiro gerado por esta via chama-se `Relatorio_Consolidado_Actualizado.xlsx`,
para não se confundir com o original.

## O que gerar

| Opção | Resultado |
|---|---|
| **Um ficheiro por cada 3 dias seguidos** | cada bloco de até 3 dias consecutivos dá um Excel |
| **Um ficheiro por dia** | cada dia dá o seu Excel |
| **Apenas o consolidado** | um único Excel com tudo |

Nas duas primeiras, o consolidado pode ser gerado também (é o que está por
omissão). Um bloco de 3 dias fecha quando chega aos 3 dias ou quando há um salto
no calendário; ficheiros do mesmo dia ficam sempre juntos.

Exemplo com os dias 1, 2, 3, 6 e 7 de Julho, em blocos de 3 dias:

```
Relatorio_01-07-2026_a_03-07-2026.xlsx   (3 PDFs)
Relatorio_06-07-2026_a_07-07-2026.xlsx   (2 PDFs — bloco interrompido pelo salto)
Relatorio_Consolidado.xlsx               (todos)
```

## Dentro de cada Excel

A folha **Relatório** tem as estatísticas no topo e os dados por baixo:

1. **Resumo** — cabe em **duas linhas**: uma com o *Nº de Processos* e outra com
   o *Total das Taxas*, e uma coluna por indicador:

   | Total | por período | Com técnico · Por atribuir | Pago · Não Pago | cada Tipo · Outros tipos |
   |---|---|---|---|---|

   É a **única parte congelada** da folha, por isso acompanha sempre a leitura da
   tabela. As colunas de tipo são preenchidas sozinhas com os tipos distintos
   definidos nos técnicos, e *Outros tipos* garante que nada fica por contar. As
   contagens abrangem até à **linha 999**, por isso quem acrescentar processos à
   mão por baixo dos dados vê as estatísticas actualizarem-se sem ter de mexer nas
   fórmulas; essas linhas já trazem a fórmula do **Tipo** e a lista pendente do
   Técnico.
2. **Técnicos** — a área editável (células amarelas) com o nome do técnico e o
   tipo correspondente, mais o nº de processos e o total das taxas de cada um. São
   **três tabelas coladas lado a lado**, 8 linhas cada por omissão (24 técnicos).
   Rola com a folha.

   Pode **inserir linhas** dentro de qualquer das três tabelas sem partir nada: as
   fórmulas referem cada tabela por intervalo, por isso a contagem de processos, o
   total das taxas e o **Tipo** de cada processo continuam certos, e o técnico novo
   entra nas estatísticas do topo. Ao inserir uma linha, copie para ela as fórmulas
   das colunas *Nº de Processos* e *Total das Taxas* (a folha de cálculo costuma
   fazê-lo sozinha ao inserir entre linhas já preenchidas).
3. **Dados** — uma linha por processo, com as 7 colunas mais a coluna de origem
   e filtro automático. Rola com a folha.

Há ainda a folha **Ficheiros de Origem** e, quando a junção encontrou processos
repetidos, a folha **Redundâncias**.

Tudo assenta em fórmulas nativas: basta escolher o técnico numa linha de dados
(há lista pendente) para o **Tipo** ser preenchido automaticamente e todas as
estatísticas se actualizarem. Se o João for do tipo A, todos os processos do João
passam a ser do tipo A.

A segunda folha, **Ficheiros de Origem**, resume cada PDF usado: período, data,
nº de processos e total das taxas.

## Linha de comandos

A mesma lógica sem browser:

```bash
npm install
node tools/gerar-cli.mjs <pasta-com-pdfs> [pasta-de-saida]
MODO=1 node tools/gerar-cli.mjs pdfs saida             # um ficheiro por dia
MODO=consolidado node tools/gerar-cli.mjs pdfs saida   # só o consolidado
```

## Testes

```bash
npm install
node tests/verificar.mjs <pasta-com-pdfs>
```

Verifica, entre outras coisas, que o nº de registos extraídos coincide com o nº de
processos do documento (contado por um marcador independente das colunas), que
nenhum total se perdeu ou veio de uma coluna vizinha, que as datas coincidem com o
nome do ficheiro, que os blocos têm dias consecutivos e que o Excel gerado contém
exactamente os mesmos registos, nas colunas e ordem pedidas.

Confirma ainda, PDF a PDF, que todas as linhas de cada ficheiro chegaram ao Excel
(17 + 52 + 50 + 35 + 54 nos exemplos), que não há linhas repetidas nem agregadas, e
que o resumo cabe em duas linhas com todos os indicadores, que as listas
pendentes estão activas e que só o resumo fica congelado.

Verifica ainda a ida e volta: pega num relatório gerado, preenche os técnicos e os
estados, relê-o com o mesmo leitor que a aplicação usa e confirma que nada se
perde — linhas, ordem, técnicos, tipos, estados e totais.

Cobre também a junção: relatórios repetidos não duplicam linhas, cada política
mantém a versão certa, a ordem das fontes é respeitada e o mesmo Nº do DU noutra
data conta como outro processo.

Confirma ainda a ordem ascendente por data (sem perder nem trocar processos de dia,
mantendo dentro de cada dia a ordem original), que o resumo conta até à linha 999 sem
que as linhas em branco falseiem a atribuição, e que as linhas livres já trazem a
fórmula do Tipo.

Nos 5 PDFs de exemplo (208 processos) passaram as 200 verificações, e o resultado foi
ainda confrontado, linha a linha, com uma extracção feita por um método totalmente
diferente (leitura ao nível do caractere): **208/208 registos iguais**.

## Estrutura

```
index.html                    ficheiro único gerado (é o que se publica)
src/pagina.html               molde do HTML
assets/style.css              estilos
assets/parser.js              leitura dos PDFs e agrupamento por dias
assets/report.js              construção dos ficheiros Excel
assets/leitor-relatorio.js    leitura de relatórios já preenchidos
assets/app.js                 ligação entre a interface e os módulos
vendor/                       pdf.js, ExcelJS e JSZip (sem CDN)
tools/construir.mjs           junta tudo no ficheiro único
tools/gerar-cli.mjs           versão de linha de comandos
tests/verificar.mjs           testes de precisão sobre PDFs reais
```

## Notas técnicas

- O texto destes PDFs está desenhado rodado 90°; as coordenadas são normalizadas
  antes de qualquer leitura.
- As colunas são localizadas pelo cabeçalho de cada página, não por posições
  fixas — as posições variam entre ficheiros.
- Quando o nome do destinatário transborda e fica colado ao total das taxas, o
  bloco de texto é dividido em palavras com posição estimada, para que o total
  não se perca.
- Os números levam o prefixo de formato `[$-416]`, que fixa o **ponto** como
  separador de milhares seja quais forem as definições regionais de quem abre o
  ficheiro. (O identificador de Portugal, `816`, usa espaço.)
- No ficheiro único, o worker do pdf.js é registado em `globalThis.pdfjsWorker`,
  que é onde o pdf.js o procura antes de o ir buscar à rede. Assim a leitura dos
  PDFs corre no mesmo sítio que o resto, sem segundo ficheiro e sem esbarrar nas
  restrições que os navegadores impõem a `file://`.
