# Fluxogramas — Posto de Piquete de Passageiros

Fluxogramas em formato **draw.io** (`.drawio`) dos procedimentos descritos no documento
*Procedimentos Aplicáveis aos Processos do Posto de Piquete de Passageiros*, expandidos com
passos decorrentes de normas aduaneiras nacionais e internacionais.

O estilo segue o modelo BPMN já usado na instituição: piscina com o nome do processo na
vertical, raias por interveniente, faixa de **fases** no topo, tarefas a azul, evento inicial
a verde, evento final a vermelho e *gateways* em losango (`X` exclusivo, `+` paralelo). Os
prazos, as notas do procedimento e o fundamento normativo estão registados numa nota amarela
por baixo da piscina.

## Convenção de cores

| Cor | Significado |
|---|---|
| **Azul** (`#DAE8FC`) | Passo descrito no documento de procedimentos |
| **Lilás** (`#E1D5E7`) | Passo acrescentado por aplicação de normas aduaneiras e por inferência lógica |
| **Azul com contorno reforçado** | Subprocesso detalhado noutro fluxograma |

## Ficheiros

| Ficheiro | Processo |
|---|---|
| `00-mapa-de-processos-completo.drawio` | **Todos os fluxogramas num só ficheiro**: mapa de processos, fluxograma geral integrado e uma página por processo, com ligações clicáveis |
| `11-fluxograma-geral-integrado.drawio` | **Fluxograma geral**: percurso completo do passageiro, em 8 fases e 8 raias, com todos os processos encadeados |
| `01-analise-documental-declaracoes.drawio` | Análise documental, gestão e controlo da tramitação das declarações aduaneiras |
| `02-avaliacao-do-canal-perfil-de-risco.drawio` | Avaliação do canal com base no perfil de risco — controlo de pessoas e bagagens |
| `03-verificacao-de-bagagens.drawio` | Verificação de bagagens |
| `04-armazenamento-e-liberacao-de-mercadoria.drawio` | Armazenamento e autorização para liberação da mercadoria |
| `05-exportacao.drawio` | Exportação |
| `06-emendas-e-cancelamento-de-declaracoes.drawio` | Emendas e cancelamento das declarações |
| `07-participacao-ao-contencioso.drawio` | Participação ao contencioso |
| `08-emissao-de-notas-de-pagamento.drawio` | Emissão de notas de pagamento |
| `09-mercadorias-restritas-e-perigosas.drawio` | Controlo de entrada e saída de mercadorias restritas e perigosas |
| `10-inspeccao-fisica-scanner-divisas.drawio` | Inspeção física e scanner das mercadorias, divisas e passageiros |

## Fluxograma geral integrado

Encadeia todos os processos num único fluxo, em oito fases:

1. Chegada e informação antecipada
2. Triagem e opção de canal — *Processo 2*
3. Verificação e inspeção — *Processos 10 e 3*
4. Classificação, valoração e contencioso — *Processos 1, 9 e 7*
5. Liquidação e pagamento — *Processos 8 e 6*
6. Armazenamento e liberação — *Processo 4*
7. Exportação — *Processo 5*
8. Encerramento e melhoria contínua — relatório mensal e uniformidade na aplicação das normas

## Fundamento dos passos acrescentados

- **Convenção de Quioto Revista**, Anexo Específico J, Capítulo 1 — viajantes, sistema de
  duplo canal, franquias de bagagem pessoal.
- **Acordo sobre a Facilitação do Comércio da OMC** — gestão de risco, informação prévia,
  direito de reclamação e recurso, notificação das decisões.
- **Quadro Normativo SAFE da OMA** — informação antecipada de passageiros, inspeção não
  intrusiva, gestão de risco e selectividade.
- **Acordo de Valoração Aduaneira da OMC** — determinação do valor aduaneiro pelo valor
  transaccional; conversão cambial à taxa oficial.
- **Código Aduaneiro e Pauta Aduaneira de Angola** — classificação pautal, liquidação,
  regime de mercadoria abandonada, infracções e contencioso.
- **Normas cambiais aplicáveis** — declaração obrigatória de divisas e tratamento do
  excedente não declarado.
- **CITES, Convenção de Basileia e IATA-DGR** — mercadorias restritas, espécies protegidas,
  resíduos e mercadorias perigosas.

## Como abrir

- **Online:** https://app.diagrams.net → *File → Open From → Device*
- **Aplicação de secretária:** draw.io Desktop
- **VS Code:** extensão *Draw.io Integration* (abre os `.drawio` directamente)

No ficheiro `00-mapa-de-processos-completo.drawio`, as caixas da primeira página e os
subprocessos do fluxograma geral têm ligação para a página do respectivo fluxograma: basta
clicar na caixa.

## Notas de leitura

- As atribuições *emissão de relatórios de actividade mensais* e *promoção da uniformidade na
  aplicação das normas aduaneiras e instrutivos conjuntos* constam do documento sem sequência
  de passos descrita; figuram no mapa geral como actividades transversais e na fase 8 do
  fluxograma geral integrado.
- Os intervenientes de cada raia foram inferidos do texto do procedimento (técnico aduaneiro,
  chefe de turno, chefe de posto, despachante oficial, polícia fiscal, órgão de tutela,
  administração, contencioso, armazém e sistema ASYCUDA). Devem ser confirmados com a
  estrutura formal do posto antes da publicação.
- Os passos a lilás são propostas de melhoria alinhadas com as normas citadas: carecem de
  validação institucional antes de serem adoptados como procedimento em vigor.
