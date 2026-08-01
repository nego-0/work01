/*
 * Agregação — o Objectivo 2. Não recebe nada de novo: pega nos registos de
 * turno já introduzidos e soma-os no período certo para alimentar os
 * relatórios. O diário é, por construção, uma parcela do mensal.
 */
(function (global, factory) {
  const mod = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = mod;
  global.PMR = global.PMR || {};
  global.PMR.agregacao = mod;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const num = (v) => {
    const n = Number(v);
    return isFinite(n) ? n : 0;
  };

  function proximoDia(iso) {
    const d = new Date(iso + 'T00:00:00');
    d.setDate(d.getDate() + 1);
    return d.toISOString().slice(0, 10);
  }

  // RGA-24: dia aduaneiro das 07h do dia D às 07h do dia D+1.
  // (Os turnos são atribuídos a um dia; juntam-se todos os do dia escolhido.)
  function rga24(registos, dia) {
    const doDia = registos.filter((r) => r.data === dia);
    const apre = [];
    doDia.forEach((r) => (r.apreensoes || []).forEach((a) => apre.push(a)));
    const valores = apre.filter((a) => a.tipo === 'Moeda/Valores');
    const mercadoria = apre.filter((a) => a.tipo === 'Mercadoria Diversa');
    const proibida = apre.filter((a) => a.tipo === 'Mercadoria Proibida/Restrita' || a.tipo === 'Duty Free');
    const realces = apre.filter((a) => a.realce && a.realceNota)
      .map((a) => a.realceNota);
    return {
      dia: dia,
      diaSeguinte: proximoDia(dia),
      chefeTurno: (doDia[0] && doDia[0].chefeTurno) || '',
      valores: valores, mercadoria: mercadoria, proibida: proibida,
      realces: realces,
      nRegistos: doDia.length
    };
  }

  function noMes(registos, mes /* 'YYYY-MM' */, lado) {
    return registos.filter((r) =>
      r.data && r.data.slice(0, 7) === mes && (!lado || r.lado === lado));
  }

  // RMT — relatório mensal do lado Terra.
  function rmt(registos, mes) {
    const regs = noMes(registos, mes, 'Terra');
    const soma = (f) => regs.reduce((a, r) => a + num(f(r)), 0);
    const contaApre = (tipo) => regs.reduce((a, r) =>
      a + (r.apreensoes || []).filter((x) => x.tipo === tipo).length, 0);
    return {
      mes: mes,
      voos: soma((r) => r.voos),
      passageiros: soma((r) => r.passageiros),
      chefeTurno: (regs[0] && regs[0].chefeTurno) || '',
      turno: (regs[0] && regs[0].turno) || '',
      apreensoes: {
        moeda: contaApre('Moeda/Valores'),
        diversa: contaApre('Mercadoria Diversa'),
        proibida: contaApre('Mercadoria Proibida/Restrita'),
        dutyFree: contaApre('Duty Free')
      },
      separados: {
        comercial: soma((r) => r.separados && r.separados.comercial),
        naoComercial: soma((r) => r.separados && r.separados.naoComercial),
        regPosterior: soma((r) => r.separados && r.separados.regPosterior)
      },
      temporarias: {
        impTemp: soma((r) => r.temporarias && r.temporarias.impTemp),
        expTemp: soma((r) => r.temporarias && r.temporarias.expTemp),
        termoResp: soma((r) => r.temporarias && r.temporarias.termoResp),
        termoComp: soma((r) => r.temporarias && r.temporarias.termoComp)
      },
      nRegistos: regs.length
    };
  }

  // RMA — relatório mensal do lado Ar (grupo Scanner).
  function rma(registos, mes) {
    const regs = noMes(registos, mes, 'Ar');
    const soma = (f) => regs.reduce((a, r) => a + num(f(r)), 0);
    const apreensoes = regs.reduce((a, r) => a + (r.apreensoes || []).length, 0);
    return {
      mes: mes,
      voos: soma((r) => r.voos),
      passageiros: soma((r) => r.passageiros),
      volumes: soma((r) => r.volumes),
      apreensoes: apreensoes,
      visitasCatering: soma((r) => r.visitasCatering),
      visitasBordo: soma((r) => r.visitasBordo),
      furgoneta: {
        scaneados: soma((r) => r.furgoneta && r.furgoneta.scaneados),
        inspecionados: soma((r) => r.furgoneta && r.furgoneta.inspecionados),
        apreensoes: soma((r) => r.furgoneta && r.furgoneta.apreensoes)
      },
      turno: (regs[0] && regs[0].turno) || '',
      chefeTurno: (regs[0] && regs[0].chefeTurno) || '',
      nRegistos: regs.length
    };
  }

  return { rga24, rmt, rma, noMes, proximoDia };
});
