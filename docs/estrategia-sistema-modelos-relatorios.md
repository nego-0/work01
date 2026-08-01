# Estratégia — Sistema de Preenchimento de Modelos e Produção Automática de Relatórios

Posto Aduaneiro do Terminal de Passageiros do Aeroporto Internacional Dr. António
Agostinho Neto — AGT, 3.ª Região Tributária.

Documento de estratégia. Descreve **o quê**, **porquê** e **por que ordem** se
constrói o sistema, e como ele se encaixa no que já existe neste repositório.
Não é código; é o plano que orienta o código a seguir.

---

## 1. Sumário executivo

Pedem-se duas coisas:

1. **Facilitar o preenchimento** de modelos de documentos específicos (CKL, TRB,
   RGA-24, RMT, RMA).
2. **Produzir relatórios de trabalho automaticamente.**

A tese central desta estratégia é que **estes dois objectivos são o mesmo cano
visto em dois momentos diferentes**. Os dados que se registam para preencher um
documento *diário* (a ocorrência da checklist, a bagagem retida, a apreensão do
turno) são exactamente a matéria-prima que, *somada ao longo do mês*, produz os
relatórios *mensais* — sem ninguém voltar a escrever um único número.

Por isso o sistema não gira à volta de «cinco formulários». Gira à volta de **um
registo de turno**, introduzido uma só vez, do qual **cada documento é uma
projecção**:

- Uma bagagem retida → **TRB** (na hora).
- As apreensões do dia (soma dos turnos) → **RGA-24** (diário).
- Os turnos do mês, agregados → **RMT** e **RMA** (mensais, calculados sozinhos).
- O estado do material no fim do turno → **CKL** (inventário + ocorrência).

Esta é, aliás, a mesma ideia que o **Gerador de Relatórios Excel** já vivo neste
repositório aplica aos PDFs do ASYCUDA: capturar linhas estruturadas uma vez e
depois **projectá-las** em resumos e consolidados. O novo sistema é o irmão
natural desse, do lado dos documentos Word, e deve nascer com a mesma arquitectura
e a mesma disciplina.

**Recomendação de topo:** construir uma segunda ferramenta **no mesmo repositório**,
partilhando infra-estrutura já provada — ficheiro único, funcionamento sem
servidor e sem Internet, bibliotecas *vendorizadas* (o **JSZip** já cá está, e um
`.docx` é um ZIP de XML), persistência local no navegador, passo de *build*,
testes de fidelidade e interface em Português.

---

## 2. Contexto e diagnóstico dos modelos

Cinco modelos, três cadências, três destinatários diferentes:

| Modelo | Nome | Cadência | Entregue a | Conteúdo essencial |
|---|---|---|---|---|
| **CKL** | Check-list de materiais e equipamentos | Fim de cada turno | Turno entrante | Inventário (descrição, qtd, estado, observação) + ocorrência diária + assinaturas *entreguei/recebi* |
| **TRB** | Termo de Recepção de Bagagens | Por evento | Companhia aérea | Bagagem retida: proveniência, voo, data, nome, n.º de etiqueta; assinaturas AGT / PFA / Companhia |
| **RGA-24** | Relatório-síntese das últimas 24 h | Diária (07h→07h) | Gestão aeroportuária | Apreensões: valores; mercadoria sujeita a direitos; mercadoria proibida/restrita |
| **RMT** | Relatório Mensal — lado **Terra** | Mensal | Gestão | Voos e passageiros; apreensões; separados de bagagem; importações/exportações temporárias |
| **RMA** | Relatório Mensal — lado **Ar** (Scanner) | Mensal | Gestão | Voos processados, passageiros, volumes inspeccionados na contentorização, apreensões, visitas ao *catering*/bordo; dados da furgoneta |

Dois lados de operação, que alimentam os dois relatórios mensais:

- **Terra** — atendimento no terminal: apreensões, separados de bagagem,
  importações/exportações temporárias → **RMT**.
- **Ar** — grupo *Scanner* na placa, cabines e contentorização → **RMA**.

### O que salta à vista nos modelos reais

- **Muita coisa é constante**: bloco institucional (AGT, 3.ª Região, nome e morada
  do posto, sítio e e-mail, logótipo, marca-d'água), fórmulas de fecho, linhas de
  assinatura. Isto **nunca deve ser reescrito** — é parte do molde.
- **Muita coisa repete-se entre documentos**: turno, data, chefe do posto
  (Edivaldo Bandeira), chefe de turno, equipa, voo, passageiro, nacionalidade,
  apreensão. Escrever isto uma vez em cada documento é a fonte principal de erro e
  de tempo perdido.
- **Os erros de digitação já lá estão** nos exemplos («Entregu5», «C3351i»,
  totais a zero deixados por copiar do mês anterior). Um formulário com validação
  e valores calculados elimina esta classe inteira de problemas.
- **O RMT chegou em `.doc`** (formato binário antigo); os restantes em `.docx`. O
  molde-mestre deve ser fixado uma vez em `.docx`.

---

## 3. Princípio central — «introduzir uma vez, projectar em todo o lado»

O coração do sistema é um **modelo de dados único** para a actividade do posto. Os
documentos deixam de ser ficheiros que se editam e passam a ser **vistas** desse
modelo. Ninguém escreve o «total de voos do mês»: o sistema conta-o a partir dos
turnos.

```
                 ┌─────────────────────────────────────────┐
                 │   DADOS-MESTRE (mudam raramente)         │
                 │   pessoas · turnos · companhias/voos ·   │
                 │   tipos de apreensão · bloco institucional│
                 └─────────────────────────────────────────┘
                                   │ alimentam
                                   ▼
   ┌──────────────────────────────────────────────────────────────┐
   │   REGISTO DIÁRIO / DE TURNO (o que se introduz todos os dias) │
   │   ocorrências · apreensões · bagagens retidas ·              │
   │   contadores de voos/passageiros/volumes · estado do material │
   └──────────────────────────────────────────────────────────────┘
        │ na hora        │ ao fecho do dia       │ ao fecho do mês
        ▼                ▼                       ▼
      TRB              RGA-24                 RMT + RMA
   (por evento)       (diário)               (mensais, agregados)
        ▲                                        ▲
        └──────────── CKL (fim de turno) ────────┘
```

Consequência directa: os números **não podem divergir** entre documentos, porque
saem todos da mesma fonte. O total de apreensões que o RGA-24 mostra num dia é,
por construção, uma parcela do que o RMT mostra no mês.

---

## 4. Modelo de dados

Dividido em **dados-mestre** (persistentes, editáveis, com valores por omissão já
preenchidos — como a lista de técnicos já faz hoje) e **dados transaccionais** (o
dia-a-dia).

### 4.1 Dados-mestre

- **Posto / instituição** — designações, morada, sítio, e-mail, logótipo,
  marca-d'água. Fixos; vivem no molde.
- **Pessoas** — nome e papel: Chefe do Posto, Chefes de Turno, equipa. (Reutilizar
  a lista de pessoal que o gerador de Excel já mantém no navegador.)
- **Turnos** — A, B, C, D, E; lado (Terra/Ar); equipa associada.
- **Companhias e voos** — companhia, n.º de voo, proveniência/destino habituais.
- **Tipos de apreensão** — Moeda/Valores · Mercadoria diversa (sujeita a direitos)
  · Mercadoria proibida/restrita · Duty Free.
- **Inventário-base da CKL** — a lista de equipamentos e quantidades esperadas
  (fotocopiadora, impressoras, scanners/Raio-X, rádios, trotinetas, balcões,
  computadores, balança, ASYCUDA, selos, TPA…), para o turno só marcar o **estado**
  e o que mudou.

### 4.2 Dados transaccionais (por turno/dia)

- **Turno** — data, turno, período, chefe de turno, lado.
- **Contadores** — n.º de voos, passageiros; volumes inspeccionados; dados da
  furgoneta; separados (comercial / não comercial / regularização a posterior);
  importações/exportações temporárias; termos.
- **Apreensões** — tipo, passageiro, nacionalidade, proveniência/destino, voo,
  valor ou descrição da mercadoria, «maior realce» (sim/não + nota).
- **Bagagens retidas** — proveniência, voo, data, passageiro, n.º de etiqueta,
  companhia. (Cada uma gera um TRB e conta como ocorrência.)
- **Estado do material (CKL)** — por item: quantidade, estado, observação; mais a
  **ocorrência diária** e as assinaturas *entreguei/recebi*.

### 4.3 Mapa modelo → campo → origem

O que torna esta estratégia executável é saber, campo a campo, de onde vem cada
valor. Resumo:

| Documento | Campos preenchidos pelo utilizador | Campos automáticos (calculados/herdados) |
|---|---|---|
| **TRB** | proveniência, voo, data, nome, n.º de etiqueta, companhia | bloco institucional, cidade+data por extenso, linhas de assinatura |
| **CKL** | estado/observação por item, ocorrência diária | inventário-base, turno, data, chefe do posto, assinaturas |
| **RGA-24** | apreensões do dia (por linha) | período «07h dia N → 07h dia N+1» por extenso, totais, «maior realce», chefe de turno, visto |
| **RMT** | contadores do lado Terra do mês | voos/passageiros e apreensões **somados dos turnos**, totais gerais, data por extenso |
| **RMA** | contadores do lado Ar do mês | voos/passageiros/volumes/visitas **somados dos turnos**, dados da furgoneta, equipa do turno, data |

A coluna da direita é o Objectivo 2 a acontecer: nasce dos registos diários já
introduzidos para o Objectivo 1.

---

## 5. Arquitectura — herdar o que já está provado

Este repositório já fez as escolhas difíceis certas. A estratégia é **não as
refazer**, mas estendê-las.

| Decisão já tomada | Mantém-se porquê |
|---|---|
| **Ficheiro único HTML, sem servidor** | Abre com dois cliques, corre de `file://`, publica-se no GitHub Pages. Não obriga a instalar nada num aeroporto. |
| **Funciona sem Internet** | A ligação no terminal é incerta; e os dados **não podem** sair do computador. |
| **Dados nunca saem da máquina** | São dados de apreensões, passageiros e bagagens — confidencialidade é obrigatória, não um extra. Sem telemetria, sem CDN, sem envio. |
| **Bibliotecas *vendorizadas*** | Já há `JSZip`, `ExcelJS`, `pdf.js` no `vendor/`. Um `.docx` é um ZIP de XML: **o JSZip já resolve metade do problema**. |
| **Persistência no navegador** | A lista de pessoal já se guarda localmente; os dados-mestre e os registos de turno seguem o mesmo caminho (localStorage/IndexedDB). |
| **Passo de *build* para ficheiro único** | `tools/construir.mjs` inclui tudo no `index.html`. A nova página entra no mesmo processo. |
| **CLI equivalente** | `tools/gerar-cli.mjs` prova que a lógica corre também fora do browser — útil para gerar os mensais em lote. |
| **Testes de fidelidade** | `tests/verificar.mjs` confere, linha a linha, que nada se perde. Os documentos precisam do mesmo rigor. |
| **Interface em Português, GitHub Pages** | Continuidade para quem já usa. |

### 5.1 Como gerar os `.docx` — preenchimento por molde (não de raiz)

Decisão de arquitectura mais importante do lado técnico:

> **Manter os documentos reais como moldes e só substituir os dados**, em vez de
> reconstruir o documento de raiz.

Porquê: o bloco institucional (logótipo, marca-d'água, tipos de letra, rodapé com
sítio/e-mail, linhas de assinatura) fica **igual ao pixel**, sem o reproduzir à
mão e sem risco de o descaracterizar. Um `.docx` é um ZIP; com o **JSZip já
vendorizado**, o motor:

1. Abre o molde `.docx`.
2. Substitui **marcadores de campo** (ex.: `{{voo}}`, `{{passageiro}}`) pelo valor.
3. **Repete linhas de tabela** para listas (cada apreensão, cada bagagem, cada
   item da checklist).
4. Reescreve o ZIP → novo `.docx` pronto a imprimir ou entregar.

É o modelo «mala-directa»: baixo risco, fidelidade máxima, e reutiliza uma
dependência que já existe. Só é preciso um pequeno motor de campos+repetição de
linhas (vendorizado ou escrito à medida, ~algumas centenas de linhas), a par do
que o `report.js` já faz para Excel.

**Formato de saída:** `.docx` (editável, igual à prática actual). O PDF final, para
entrega, sai do «Guardar como PDF»/imprimir do Word ou do navegador — evita
embutir um motor de PDF pesado e mantém o ficheiro único leve. (Geração de PDF
offline pode entrar como extra numa fase posterior, se se justificar.)

### 5.2 Onde vive no repositório

Segunda ferramenta, mesmo repositório: uma página nova (ex.: `src/documentos.html`
→ `documentos.html` no *build*), a partilhar `vendor/`, o passo de *build* e os
testes. Os dois utilitários — «Gerador de Relatórios Excel» e «Modelos e
Relatórios» — passam a ser as duas faces do mesmo sítio. Máximo reaproveitamento,
zero infra-estrutura nova.

---

## 6. Como o mesmo cano serve os dois objectivos

| | Objectivo 1 — preencher modelos | Objectivo 2 — relatórios automáticos |
|---|---|---|
| **Entrada** | O utilizador preenche um formulário guiado (com validação, listas pendentes, valores por omissão) | Nada de novo — usa os registos diários já introduzidos no Objectivo 1 |
| **O que faz o sistema** | Preenche o molde `.docx` e devolve o documento | **Agrega** os turnos do período e projecta nos moldes mensais/diário |
| **Ganho** | Menos tempo, zero reescrita do que é constante, menos erros | Os mensais deixam de se «copiar do mês anterior»; saem dos dados reais, reconciliados |
| **Cobre** | TRB, CKL, e a introdução das apreensões/bagagens | RGA-24 (diário), RMT, RMA (mensais) |

O Objectivo 1 é a porta de entrada dos dados; o Objectivo 2 é a recompensa de os
ter introduzido de forma estruturada uma só vez.

---

## 7. Faseamento

Ordenado por **valor imediato ÷ risco**. Cada fase entrega algo utilizável.

### Fase 0 — Fundações partilhadas
- Fixar os cinco moldes em `.docx` com marcadores de campo e zonas de repetição
  de linhas (o RMT convertido do `.doc`).
- Motor de preenchimento de `.docx` sobre o JSZip (campos + repetição de linhas).
- Modelo de dados-mestre + persistência local; reaproveitar a lista de pessoal.
- Entrar no *build* de ficheiro único e no arranque dos testes.

### Fase 1 — Objectivo 1, o ganho rápido (por evento / fim de turno)
- **TRB**: formulário curto → termo pronto. É o caso mais simples e mais frequente.
- **CKL**: inventário-base pré-preenchido; o turno só marca estado, observações e
  a ocorrência; assinaturas *entreguei/recebi*.
- Introdução das **apreensões** e **bagagens** do turno (alimenta já as fases
  seguintes).
- *Entregável:* preencher TRB e CKL passa de minutos de cópia a segundos.

### Fase 2 — Objectivo 2, a agregação (o valor grande)
- **RGA-24**: junta as apreensões do dia (soma dos turnos, 07h→07h), calcula
  totais e o período por extenso.
- **RMT / RMA**: agregam o mês por lado (Terra/Ar) — voos, passageiros, volumes,
  separados, visitas, furgoneta, apreensões — e projectam nos moldes mensais.
- *Entregável:* os relatórios diário e mensais saem sozinhos dos registos, sem
  reescrita.

### Fase 3 — Qualidade e reconciliação
- **Validações**: campos obrigatórios, nacionalidade/voo coerentes, sem totais
  esquecidos a zero.
- **Reconciliação**: as apreensões dos RGA-24 do mês têm de bater certo com o RMT;
  os voos/passageiros podem cruzar-se com os dados que o gerador de Excel já
  extrai do ASYCUDA, em vez de se reescreverem.
- **Testes de fidelidade** (à imagem de `tests/verificar.mjs`): confirmar que o
  `.docx` gerado contém exactamente os dados introduzidos, que as tabelas repetem
  o número certo de linhas, que os totais agregados batem com a soma dos turnos, e
  que o bloco institucional se mantém intacto.
- Exportação para PDF (opcional) e histórico/arquivo dos documentos gerados.

---

## 8. Qualidade, testes e reconciliação

O repositório já tem uma cultura de teste forte («179 verificações», «208/208
registos iguais», confronto por um método independente). O novo sistema herda-a:

- **Fidelidade do molde** — o documento gerado abre no Word sem aviso e é
  visualmente idêntico ao original, mudando só os dados.
- **Fidelidade dos dados** — cada campo e cada linha de tabela contém exactamente
  o que foi introduzido; nada se perde nem troca de coluna.
- **Correcção da agregação** — o total mensal é igual à soma dos turnos; o diário
  é uma parcela do mensal. Testar com meses inventados de ponta a ponta.
- **Ida-e-volta** — reabrir um documento gerado e reler os dados confirma que o
  ciclo fecha (como o teste de *round-trip* já existente para o Excel).

---

## 9. Segurança e confidencialidade

Não é um requisito acessório — é o motivo pelo qual a arquitectura offline é a
correcta. Os documentos contêm nomes de passageiros, nacionalidades, voos, valores
e mercadorias apreendidas.

- **Os dados não saem da máquina**: sem servidor, sem CDN, sem telemetria — igual
  ao que já se faz com os PDFs.
- **Persistência local** no navegador do posto; sem contas nem nuvem.
- **Sem dependências externas em tempo de execução**: tudo *vendorizado*.
- Se um dia se quiser partilha entre turnos, que seja por ficheiro exportado à
  mão (o próprio registo do turno), nunca por serviço online por omissão.

---

## 10. Riscos e mitigações

| Risco | Mitigação |
|---|---|
| Descaracterizar o layout institucional ao gerar | Preencher **moldes reais**, não construir de raiz (§5.1) |
| RMT em `.doc` legado | Converter e fixar o molde-mestre em `.docx` uma única vez (Fase 0) |
| Números a divergir entre documentos | **Fonte única de dados**; o diário é uma parcela do mensal por construção |
| Adopção — tem de ser mais fácil que copiar o Word do mês anterior | Valores por omissão, listas pendentes, inventário-base pré-preenchido; «introduzir uma vez» tem de ganhar de forma óbvia |
| Erros de digitação (como os dos exemplos) | Validação e campos calculados eliminam a classe inteira |
| Perder dados locais (navegador limpo, máquina trocada) | Exportar/importar o registo de turno como ficheiro; arquivo dos documentos gerados |
| Âmbito a crescer para «app grande» | Manter ficheiro único, offline, sem servidor — a mesma disciplina que manteve o gerador de Excel simples |

---

## 11. Métricas de sucesso

- **Tempo** para emitir um TRB e uma CKL cai de minutos para segundos.
- **Zero reescrita** de números nos mensais: RMT e RMA saem agregados.
- **Zero divergências** entre o somatório dos RGA-24 e o RMT do mês.
- **Erros de preenchimento** (totais esquecidos, campos trocados) tendem a zero
  por via da validação.
- A ferramenta corre **sem Internet** e os dados **nunca** saem do posto.

---

## 12. Decisões em aberto (a confirmar com o posto)

1. **Mesmo repositório** ou repositório separado? (Recomendação: mesmo repositório,
   como segunda ferramenta.)
2. **PDF** faz falta já, ou basta `.docx` + «Guardar como PDF» do Word numa
   primeira fase?
3. **Partilha entre turnos** — chega exportar/importar um ficheiro de registo, ou
   é preciso uma pasta partilhada da rede interna do posto?
4. **Cruzamento com o ASYCUDA/Excel** — os voos/passageiros e processos devem ser
   importados do que o gerador de Excel já extrai, em vez de reintroduzidos?
5. Confirmar a **lista de itens da CKL** e os **tipos de apreensão** como
   dados-mestre definitivos.

---

## 13. Próximo passo concreto

Fase 0 + Fase 1: transformar os cinco anexos em moldes `.docx` parametrizados e
entregar o formulário do **TRB** e da **CKL** a funcionar no ficheiro único —
o menor troço que já poupa trabalho real ao posto e valida a arquitectura de
ponta a ponta.
