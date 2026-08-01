/*
 * Dados-mestre do posto (mudam raramente) e persistência local.
 * Vêm já preenchidos com valores por omissão retirados dos modelos reais;
 * podem ser alterados e ficam guardados no navegador (localStorage).
 */
(function (global, factory) {
  const mod = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = mod;
  global.PMR = global.PMR || {};
  global.PMR.dados = mod;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const CHAVE = 'pmr.dadosMestre.v1';

  // Bloco institucional constante, presente no topo/rodapé de todos os documentos.
  const INSTITUICAO = {
    regiao: '3.ª REGIÃO TRIBUTÁRIA',
    posto: 'POSTO ADUANEIRO DO TERMINAL DE PASSAGEIROS DO AEROPORTO ' +
           'INTERNACIONAL Dr. ANTÓNIO AGOSTINHO NETO',
    postoCurto: 'Posto Aduaneiro do Terminal de Passageiros do Aeroporto ' +
                'Internacional Dr. António Agostinho Neto',
    chefePosto: 'Edivaldo Mauro da Silva Bandeira',
    chefePostoCurto: 'Edivaldo Bandeira',
    cidade: 'Luanda',
    website: 'www.agt.minfin.gov.ao',
    email: 'correspondencia.agt@minfin.gov.ao',
    morada: 'Av. 21 de Janeiro (AEROPORTO 4 DE FEVEREIRO)'
  };

  const TURNOS = ['A', 'B', 'C', 'D', 'E'];

  const CHEFES_TURNO = [
    'Jessé José Maria Martins',
    'Virgílio Márcio Tomás da Conceição',
    'Abel Júlio'
  ];

  // Equipa-tipo (grupo Scanner, turno E do RMA de exemplo).
  const EQUIPA = [
    'Abel Júlio', 'Joaquim Miala', 'Francisca Nkabi', 'Fernanda Cativa',
    'Vasco Vaz', 'Anacleto Kipungo', 'Lucrécio dos Santos', 'Rui Castanheira',
    'Heriberto António', 'Ana Massiala', 'Gioconda Silvério', 'João Kimuini',
    'Jeorgina dos Santos', 'Lurdes Manuel', 'Francisco Neto', 'Claudete Manuel'
  ];

  // Tipos de apreensão — a nomenclatura partilhada por RGA-24 e RMT.
  const TIPOS_APREENSAO = [
    'Moeda/Valores',
    'Mercadoria Diversa',
    'Mercadoria Proibida/Restrita',
    'Duty Free'
  ];

  // Inventário-base da checklist: o turno só marca o estado e a observação.
  const ITENS_CKL = [
    { descricao: 'Máquina Fotocopiadora/Scanner', qtd: '1' },
    { descricao: 'Máquina Impressora', qtd: '2' },
    { descricao: 'Scanners – Sala 1 (Raio X NUCTECH CX10180T)', qtd: '1' },
    { descricao: 'Scanners – Sala 2 (Raio X NUCTECH CX10180T / Bodyscan)', qtd: '1' },
    { descricao: 'Rádio de comunicação – Sala', qtd: '1' },
    { descricao: 'Trotinetas eléctricas', qtd: '4' },
    { descricao: 'Balcão Inox', qtd: '9' },
    { descricao: 'Computadores', qtd: '8' },
    { descricao: 'Balança', qtd: '1' },
    { descricao: 'Sistema (ASYCUDA)', qtd: '-' },
    { descricao: 'Pauta Aduaneira Dec. Leg. Pres. n.º 10/19', qtd: '-' },
    { descricao: 'Chaves da Sala 1 e 2 e Armazéns', qtd: '10' },
    { descricao: 'Selos não reutilizáveis', qtd: '6' },
    { descricao: 'Máquina para contagem de dinheiro (Cash Tester)', qtd: '1' },
    { descricao: 'Livro de Reclamações', qtd: '1' },
    { descricao: 'TPA', qtd: '1' }
  ];

  const ESTADOS = ['Operacional', 'N/Operacional', 'Entregue', 'Não entregue'];

  const LADOS = ['Terra', 'Ar'];

  const SEPARADOS = ['Comercial', 'Não Comercial', 'Regularização a Posterior'];

  const TEMPORARIAS = [
    'Importação Temporária', 'Exportação Temporária',
    'Termo de Responsabilidade', 'Termo de Compromisso'
  ];

  function predefinicao() {
    return {
      instituicao: Object.assign({}, INSTITUICAO),
      turnos: TURNOS.slice(),
      chefesTurno: CHEFES_TURNO.slice(),
      equipa: EQUIPA.slice(),
      tiposApreensao: TIPOS_APREENSAO.slice(),
      itensCkl: ITENS_CKL.map((i) => Object.assign({}, i)),
      estados: ESTADOS.slice(),
      lados: LADOS.slice(),
      separados: SEPARADOS.slice(),
      temporarias: TEMPORARIAS.slice()
    };
  }

  function temStorage() {
    return typeof localStorage !== 'undefined' && localStorage !== null;
  }

  function carregar() {
    const base = predefinicao();
    if (!temStorage()) return base;
    try {
      const bruto = localStorage.getItem(CHAVE);
      if (!bruto) return base;
      const guardado = JSON.parse(bruto);
      // Mistura sobre a predefinição para não perder campos novos.
      return Object.assign(base, guardado, {
        instituicao: Object.assign(base.instituicao, guardado.instituicao || {})
      });
    } catch (e) {
      return base;
    }
  }

  function guardar(d) {
    if (!temStorage()) return;
    localStorage.setItem(CHAVE, JSON.stringify(d));
  }

  function repor() {
    if (temStorage()) localStorage.removeItem(CHAVE);
    return predefinicao();
  }

  return { predefinicao, carregar, guardar, repor, CHAVE };
});
