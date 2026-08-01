/*
 * Registos de turno — os dados transaccionais introduzidos no dia-a-dia.
 * É a fonte única de que saem, por agregação, os relatórios (RGA-24, RMT, RMA).
 * Guardados no navegador; exportáveis/importáveis como ficheiro JSON.
 */
(function (global, factory) {
  const mod = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = mod;
  global.PMR = global.PMR || {};
  global.PMR.registos = mod;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const CHAVE = 'pmr.registos.v1';

  function temStorage() {
    return typeof localStorage !== 'undefined' && localStorage !== null;
  }

  // Estrutura de um registo de turno (todos os campos numéricos são opcionais):
  // {
  //   id, data:'YYYY-MM-DD', turno:'B', lado:'Terra'|'Ar', chefeTurno,
  //   voos, passageiros, volumes, visitasCatering, visitasBordo,
  //   furgoneta:{scaneados, inspecionados, apreensoes},
  //   separados:{comercial, naoComercial, regPosterior},
  //   temporarias:{impTemp, expTemp, termoResp, termoComp},
  //   apreensoes:[{tipo, passageiro, nacionalidade, origemDestino, voo,
  //                valor, descricao, realce, realceNota}],
  //   bagagens:[{proveniencia, voo, dataVoo, passageiro, etiqueta, companhia}]
  // }
  function novoRegisto() {
    return {
      id: 'r' + Date.now() + Math.random().toString(36).slice(2, 7),
      data: new Date().toISOString().slice(0, 10),
      turno: 'B', lado: 'Terra', chefeTurno: '',
      voos: 0, passageiros: 0, volumes: 0, visitasCatering: 0, visitasBordo: 0,
      furgoneta: { scaneados: 0, inspecionados: 0, apreensoes: 0 },
      separados: { comercial: 0, naoComercial: 0, regPosterior: 0 },
      temporarias: { impTemp: 0, expTemp: 0, termoResp: 0, termoComp: 0 },
      apreensoes: [], bagagens: []
    };
  }

  function listar() {
    if (!temStorage()) return [];
    try {
      const bruto = localStorage.getItem(CHAVE);
      return bruto ? JSON.parse(bruto) : [];
    } catch (e) {
      return [];
    }
  }

  function gravarTodos(lista) {
    if (temStorage()) localStorage.setItem(CHAVE, JSON.stringify(lista));
  }

  function guardar(reg) {
    const lista = listar();
    const i = lista.findIndex((r) => r.id === reg.id);
    if (i >= 0) lista[i] = reg; else lista.push(reg);
    lista.sort((a, b) => (a.data < b.data ? -1 : a.data > b.data ? 1 :
      (a.turno < b.turno ? -1 : 1)));
    gravarTodos(lista);
    return reg;
  }

  function apagar(id) {
    gravarTodos(listar().filter((r) => r.id !== id));
  }

  function obter(id) {
    return listar().find((r) => r.id === id) || null;
  }

  return { CHAVE, novoRegisto, listar, guardar, apagar, obter, gravarTodos };
});
