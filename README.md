# Gerador de Relatórios Excel

Aplicação web que lê relatórios PDF do tipo **Document List** (ASYCUDAWorld) e gera
ficheiros Excel prontos a trabalhar, agrupados por blocos de **3 dias seguidos**.

Corre inteiramente no navegador — os PDFs nunca saem do computador de quem a usa.

## Abrir a aplicação

Endereço final: **https://nego-0.github.io/work01/**

O site é publicado pelo workflow
[`.github/workflows/static.yml`](.github/workflows/static.yml) a cada `push` para o
ramo de trabalho, e pode ser lançado à mão em *Actions → Deploy static content to
Pages → Run workflow*. Requer *Settings → Pages → Source:* **GitHub Actions**.

Também funciona a partir de qualquer servidor estático — basta servir a pasta do
repositório (os módulos ES não funcionam com `file://`):

```bash
npm start          # http://127.0.0.1:8080
```

## O que faz

Para cada PDF, a aplicação lê a tabela de processos e produz as 7 colunas pedidas,
por esta ordem:

| # | Coluna | Origem |
|---|---|---|
| 1 | **Período** | nome do ficheiro (Madrugada, Manhã, Tarde, Noite) |
| 2 | **Data de Reg.** | coluna *Data de Reg.* do PDF |
| 3 | **Total das Taxas** | coluna *Total das taxas* do PDF |
| 4 | **Estado** | preenchido com `Pago` (configurável) |
| 5 | **Nº do DU** | coluna *N° do DU* do PDF |
| 6 | **Técnico** | em branco, a preencher no Excel |
| 7 | **Tipo** | fórmula: resulta do técnico escolhido |
| — | **Ficheiro (PDF)** | coluna extra, a seguir às 7: de que PDF veio a linha |

**Uma linha por processo.** Todos os processos de todos os PDFs do bloco vão para a
tabela, sem agregações — cada linha tem a sua própria célula de **Técnico**, para se
indicar quem foi o responsável por aquele processo em concreto.

### Nomes dos ficheiros PDF

O período e a data são lidos do nome. São aceites, entre outros:

```
Manhã 21.07.2026.pdf      01.07.2026_Tarde.pdf      2026-07-23 Madrugada.pdf
```

Se o nome não tiver data, é usada a data predominante dentro do próprio PDF.

### Agrupamento

Os ficheiros são ordenados por data e agrupados em blocos de até **3 dias
consecutivos** — cada bloco dá origem a um Excel. Um bloco fecha quando chega aos
3 dias ou quando há um salto no calendário. Ficheiros do mesmo dia ficam sempre
no mesmo bloco. No fim é gerado ainda **um único ficheiro consolidado** com todos
os registos.

Exemplo com os dias 1, 2, 3, 6 e 7 de Julho:

```
Relatorio_01-07-2026_a_03-07-2026.xlsx   (3 PDFs)
Relatorio_06-07-2026_a_07-07-2026.xlsx   (2 PDFs — bloco interrompido pelo salto)
Relatorio_Consolidado.xlsx               (todos)
```

## Dentro de cada Excel

A folha **Relatório** tem as estatísticas no topo e os dados por baixo:

1. **Resumo** — total de processos e de taxas, contagem e soma por período,
   e quantos processos ainda estão por atribuir.
2. **Técnicos e Tipos** — à esquerda, a área editável (células amarelas) onde se
   escreve o nome do técnico e o tipo correspondente, com o nº de processos e o
   total das taxas de cada um; à direita, a quantidade e o total de cada tipo,
   sem repetições.
3. **Dados** — uma linha por processo, com as 7 colunas mais a coluna de origem
   e filtro automático.

As estatísticas ocupam apenas as primeiras 30 linhas e não há painéis
congelados, para a tabela de dados começar à vista e deslizar normalmente.

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
DIAS=5 node tools/gerar-cli.mjs pdfs saida    # blocos de 5 dias
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
(17 + 52 + 50 + 35 + 54 nos exemplos) e que não há linhas repetidas nem agregadas.

Nos 5 PDFs de exemplo (208 processos) passaram as 96 verificações, e o resultado foi
ainda confrontado, linha a linha, com uma extracção feita por um método totalmente
diferente (leitura ao nível do caractere): **208/208 registos iguais**.

## Estrutura

```
index.html            interface
assets/parser.js      leitura dos PDFs e agrupamento por dias
assets/report.js      construção dos ficheiros Excel
assets/app.js         ligação entre a interface e os módulos
vendor/               pdf.js, ExcelJS e JSZip (sem CDN — funciona offline)
tools/gerar-cli.mjs   versão de linha de comandos
tests/verificar.mjs   testes de precisão sobre PDFs reais
```

## Notas técnicas

- O texto destes PDFs está desenhado rodado 90°; as coordenadas são normalizadas
  antes de qualquer leitura.
- As colunas são localizadas pelo cabeçalho de cada página, não por posições
  fixas — as posições variam entre ficheiros.
- Quando o nome do destinatário transborda e fica colado ao total das taxas, o
  bloco de texto é dividido em palavras com posição estimada, para que o total
  não se perca.
