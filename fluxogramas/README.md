# Fluxogramas — Posto de Piquete de Passageiros

Fluxogramas em formato **draw.io** (`.drawio`) dos procedimentos descritos no documento
*Procedimentos Aplicáveis aos Processos do Posto de Piquete de Passageiros*.

O estilo segue o modelo BPMN já usado na instituição: piscina com o nome do processo na
vertical, raias por interveniente, tarefas a azul, evento inicial a verde, evento final a
vermelho e *gateways* em losango (`X` exclusivo, `+` paralelo). As notas e prazos de cada
procedimento estão registados numa nota amarela por baixo da piscina.

## Ficheiros

| Ficheiro | Processo |
|---|---|
| `00-mapa-de-processos-completo.drawio` | **Todos os processos num só ficheiro**: página 1 com o mapa geral e ligações clicáveis, seguida de uma página por fluxograma |
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

## Como abrir

- **Online:** https://app.diagrams.net → *File → Open From → Device*
- **Aplicação de secretária:** draw.io Desktop
- **VS Code:** extensão *Draw.io Integration* (abre os `.drawio` directamente)

No ficheiro `00-mapa-de-processos-completo.drawio`, as caixas da primeira página têm
ligação para a página do respectivo fluxograma: basta clicar na caixa.

## Notas de leitura

- As **ligações entre processos** estão representadas no mapa geral: a chegada do
  passageiro percorre os processos 10 → 2 → 3 → 1 → 8 → 4; os desvios encaminham para os
  processos 9 (mercadorias restritas e perigosas), 7 (contencioso) e 6 (emendas e
  cancelamento); a exportação (processo 5) constitui um fluxo próprio de saída.
- As atribuições *emissão de relatórios de actividade mensais* e *promoção da uniformidade
  na aplicação das normas aduaneiras e instrutivos conjuntos* constam do documento sem
  sequência de passos descrita, pelo que figuram no mapa geral como actividades
  transversais, sem fluxograma próprio.
- Os intervenientes de cada raia foram inferidos do texto do procedimento (técnico
  aduaneiro, chefe de turno, chefe de posto, despachante oficial, polícia fiscal, órgão de
  tutela, administração, contencioso, armazém e sistema ASYCUDA). Devem ser confirmados
  com a estrutura formal do posto antes da publicação.
