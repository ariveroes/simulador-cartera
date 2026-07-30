/* ============================================================
 * REENTAL WEALTH — MOTOR DE CÁLCULO
 *
 * Traducción a JavaScript de la pestaña "Cálculos" de
 * Plantilla_Propuesta_ES_Cartera_v4.xlsx
 *
 * Cada función indica entre corchetes la celda original de la
 * que procede, para poder auditar cualquier cifra contra el sheet.
 *
 * Convenciones:
 *  - Todos los porcentajes son fracciones (0.16 = 16%), igual que
 *    en el sheet.
 *  - "estatus" es una de las claves de ESTATUS.
 *  - fx es el tipo de cambio EUR->USD (1.145405 = 1 € son 1,145405 $).
 * ============================================================ */

'use strict';

const ESTATUS = ['Reentel', 'ReentelPro', 'SuperReentel'];

/** Estados que excluyen un proyecto de la selección. */
const ESTADOS_EXCLUIDOS = ['CERRADO', 'NO LANZADO'];

/** Mínimo de meses restantes para que un proyecto sea proponible. */
const MIN_MESES = 6;

/* ------------------------------------------------------------
 * 1. CÁLCULO POR INMUEBLE
 * ------------------------------------------------------------ */

/**
 * Calcula las magnitudes derivadas de un inmueble para un estatus dado.
 *
 * @param {Object} inm            Datos del inmueble (ver README)
 * @param {number} tokens         Nº de tokens adquiridos
 * @param {number} pxPagado       Precio efectivamente pagado por token
 * @param {string} estatus        'Reentel' | 'ReentelPro' | 'SuperReentel'
 * @param {number} fx             Tipo de cambio EUR->USD
 */
function calcularInmueble(inm, { tokens, pxPagado, estatus, fx }) {
  const px = pxPagado != null ? pxPagado : inm.pxEmision;
  const bruto = tokens * px;

  // [Q121] / [R121] — importe según divisa del proyecto
  const importeEUR = inm.divisa === '€' ? bruto : bruto / fx;
  const importeUSD = inm.divisa === '$' ? bruto : bruto * fx;

  // [AA121] rentabilidad de alquiler ajustada al precio pagado
  const rentAlquilerAj = (inm.rentAlquiler[estatus] || 0) * (inm.pxEmision / px);

  // [AD121] plusvalía ajustada al precio pagado
  const plusvaliaAj =
    (inm.pxEmision * (1 + (inm.plusvalia[estatus] || 0)) - px) / px;

  // [AG121] rentabilidad total del periodo restante
  const rentTotal = (rentAlquilerAj / 12) * inm.mesesRestantes + plusvaliaAj;

  // [AJ121] anualización LINEAL (no compuesta), tal cual el sheet
  const rentAnualizada =
    inm.mesesRestantes > 0 ? (rentTotal * 12) / inm.mesesRestantes : 0;

  // [F96] / [I96] flujos en euros
  const rentaMensualEUR = (importeEUR * rentAlquilerAj) / 12;
  const plusvaliaEUR = importeEUR * plusvaliaAj;

  return {
    ...inm,
    tokens,
    pxPagado: px,
    importeEUR,
    importeUSD,
    rentAlquilerAj,
    plusvaliaAj,
    rentTotal,
    rentAnualizada,
    rentaMensualEUR,
    plusvaliaEUR,
  };
}

/* ------------------------------------------------------------
 * 2. AGREGACIÓN DE CARTERA
 * ------------------------------------------------------------ */

const suma = (arr) => arr.reduce((a, b) => a + b, 0);

/**
 * Agrega una lista de inmuebles ya calculados.
 *
 * @param {Array}  items           salida de calcularInmueble()
 * @param {number} fx              tipo de cambio
 * @param {number} costeEstatusEUR coste total del estatus en €, introducido
 *                                 a mano por el usuario (0 para Reentel)
 */
function calcularCartera(items, { fx, costeEstatusEUR = 0 }) {
  const totalTokens = suma(items.map((i) => i.tokens));

  // [E121] peso de cada inmueble = tokens / tokens totales
  const pesos = items.map((i) => (totalTokens ? i.tokens / totalTokens : 0));

  const invEUR = suma(items.map((i) => i.importeEUR)); // [H47]
  const invUSD = suma(items.map((i) => i.importeUSD)); // [K47]

  // [I47] / [G47] reparto por divisa del proyecto
  const enEUR = suma(items.filter((i) => i.divisa === '€').map((i) => i.importeEUR));
  const pctEUR = invEUR ? enEUR / invEUR : 0;

  const ponderado = (campo) => suma(items.map((i, k) => i[campo] * pesos[k]));

  // Subtotales para las variantes "solo proyectos con..."  [AL20 / AM20]
  const baseAlquiler = suma(
    items.filter((i) => i.rentAlquilerAj !== 0).map((i) => i.importeEUR)
  );
  const baseFinal = suma(
    items.filter((i) => i.plusvaliaAj !== 0).map((i) => i.importeEUR)
  );

  // [D63] / [G63] reponderado sobre el subconjunto que realmente paga
  const soloAlquiler = baseAlquiler
    ? suma(
        items
          .filter((i) => i.rentAlquilerAj !== 0)
          .map((i) => i.rentAlquilerAj * (i.importeEUR / baseAlquiler))
      )
    : 0;
  const soloFinal = baseFinal
    ? suma(
        items
          .filter((i) => i.plusvaliaAj !== 0)
          .map((i) => i.plusvaliaAj * (i.importeEUR / baseFinal))
      )
    : 0;

  const costeEstatusUSD = costeEstatusEUR * fx;

  return {
    items,
    pesos,
    nInmuebles: items.filter((i) => i.tokens > 0).length, // [D47]
    totalTokens, // [E47]
    mediaMeses: suma(items.map((i, k) => i.mesesRestantes * pesos[k])), // [F47]
    invEUR,
    invUSD,
    pctEUR,
    pctUSD: 1 - pctEUR, // [J47]
    rentAlquiler: ponderado('rentAlquilerAj'), // [D62]
    plusvalia: ponderado('plusvaliaAj'), // [G62]
    rentTotal: ponderado('rentTotal'), // [J62]
    rentAnualizada: ponderado('rentAnualizada'), // [M62]
    soloAlquiler,
    soloFinal,
    costeEstatusEUR, // [E52]
    costeEstatusUSD, // [E51]
    totalEUR: invEUR + costeEstatusEUR, // [E57]
    totalUSD: invUSD + costeEstatusUSD, // [E56]
    distribucion: {
      ubicacion: agrupar(items, pesos, 'ubicacion'), // [M47:S47]
      dividendos: agrupar(items, pesos, 'tipologiaDividendos'), // [T47:W47]
      explotacion: agrupar(items, pesos, 'tipologiaExplotacion'),
    },
  };
}

function agrupar(items, pesos, campo) {
  const out = {};
  items.forEach((i, k) => {
    const clave = i[campo] || 'Otros';
    out[clave] = (out[clave] || 0) + pesos[k];
  });
  return out;
}

/* ------------------------------------------------------------
 * 3. REINVERSIÓN E INTERÉS COMPUESTO
 * ------------------------------------------------------------ */

const HORIZONTES = [6, 12, 24, 36, 60]; // [D72:H72]

/**
 * Factor de una renta mensual de n periodos capitalizada a im.
 * Equivale a ((1+im)^n - 1)/im, con el límite im -> 0.
 */
function factorRenta(im, n) {
  if (n <= 0) return 0;
  if (im === 0) return n;
  return (Math.pow(1 + im, n) - 1) / im;
}

/**
 * Replica [D75:H75] — valor final del portfolio con reinversión.
 *
 * Los cuatro sumandos del sheet:
 *   1. capital invertido
 *   2. rentas mensuales capitalizadas hasta min(H, meses restantes)
 *   3. plusvalía cobrada al vencimiento y reinvertida el resto del tiempo
 *   4. principal devuelto al vencimiento y reinvertido el resto del tiempo
 */
function calcularReinversion(
  cartera,
  { tasaReinversionAnual, tasaStakingAnual = 0, valorEstatusEUR = 0, horizontes = HORIZONTES }
) {
  const im = tasaReinversionAnual / 12; // [D68]
  const capital = cartera.invEUR; // [D74]

  return horizontes.map((H) => {
    let valorFinal = capital;

    cartera.items.forEach((i) => {
      const T = i.mesesRestantes;

      // 2 — rentas mensuales mientras el proyecto vive
      valorFinal += i.rentaMensualEUR * factorRenta(im, Math.min(H, T));

      // 3 — plusvalía al vencimiento, reinvertida
      if (H >= T) valorFinal += i.plusvaliaEUR * Math.pow(1 + im, H - T);

      // 4 — principal devuelto al vencimiento, reinvertido
      if (H > T) valorFinal += i.importeEUR * (Math.pow(1 + im, H - T) - 1);
    });

    const ganancia = valorFinal - capital; // [D76]
    const staking = ((tasaStakingAnual * valorEstatusEUR) / 12) * H; // [D84] lineal
    const denom = capital + valorEstatusEUR;

    return {
      horizonte: H,
      capital,
      valorFinal,
      ganancia,
      rentAcumulada: capital ? ganancia / capital : 0, // [D77]
      staking,
      gananciaConStaking: ganancia + staking,
      rentAcumuladaConStaking: denom ? (ganancia + staking) / denom : 0, // [D92]
    };
  });
}

/* ------------------------------------------------------------
 * 4. SELECCIÓN DE PROYECTOS
 *
 * Esta parte NO existe en la plantilla: allí los proyectos se
 * eligen a mano. Es lógica nueva.
 * ------------------------------------------------------------ */

const OBJETIVOS = {
  RECURRENTE: 'ingresos_periodicos',
  FINAL: 'rentabilidad_final',
  MAXIMA: 'maxima_rentabilidad',
};

/**
 * Filtra y ordena el universo de proyectos según las preferencias.
 *
 * Nota de diseño: la clasificación por objetivo NO usa el texto de
 * "tipología de dividendos" sino la presencia real de cada componente
 * de rentabilidad, igual que hacen [AL22] y [AM22] del sheet. Hay
 * proyectos rotulados "mensuales + final" cuya rentabilidad de
 * alquiler es 0.
 *
 * @param {Array}  universo
 * @param {string} objetivo   una clave de OBJETIVOS
 * @param {Array}  mercados   ubicaciones permitidas ([] = todas)
 * @param {string} estatus
 * @param {number} minMeses   mínimo de meses restantes para entrar
 */
function seleccionarProyectos(
  universo,
  { objetivo, mercados = [], estatus, minMeses = MIN_MESES }
) {
  const nominal = (p) => {
    // Para ordenar aún no hay tokens ni descuento: se usa el precio
    // de emisión, así el ranking no depende del importe invertido.
    const alquiler = p.rentAlquiler[estatus] || 0;
    const plusvalia = p.plusvalia[estatus] || 0;
    const total = (alquiler / 12) * p.mesesRestantes + plusvalia;
    return {
      alquiler,
      plusvalia,
      total,
      anualizada: p.mesesRestantes > 0 ? (total * 12) / p.mesesRestantes : 0,
    };
  };

  let candidatos = universo
    .filter((p) => !ESTADOS_EXCLUIDOS.includes(String(p.estado).toUpperCase()))
    .filter((p) => (mercados.length ? mercados.includes(p.ubicacion) : true))
    // meses restantes calculados desde la cascada CA -> BD -> I -> F.
    // Descarta nulos, negativos y los que están por debajo del mínimo.
    .filter((p) => Number.isFinite(p.mesesRestantes) && p.mesesRestantes >= minMeses)
    .map((p) => ({ ...p, _m: nominal(p) }));

  if (objetivo === OBJETIVOS.RECURRENTE) {
    candidatos = candidatos.filter((p) => p._m.alquiler !== 0);
    candidatos.sort((a, b) => b._m.alquiler - a._m.alquiler);
  } else if (objetivo === OBJETIVOS.FINAL) {
    candidatos = candidatos.filter((p) => p._m.alquiler === 0 && p._m.plusvalia !== 0);
    candidatos.sort((a, b) => b._m.plusvalia - a._m.plusvalia);
  } else {
    candidatos.sort((a, b) => b._m.anualizada - a._m.anualizada);
  }

  return candidatos;
}

/**
 * Reparte un importe en € a partes iguales entre los mejores candidatos,
 * redondeando a tokens enteros hacia abajo.
 *
 * Si a un proyecto no le llega ni para un token se descarta y se
 * reparte otra vez entre los restantes.
 *
 * @returns {{seleccion: Array, noAsignadoEUR: number}}
 */
function repartirImporte(candidatos, importeEUR, { fx, maxProyectos = 5 }) {
  let lista = candidatos.slice(0, Math.max(0, maxProyectos));

  while (lista.length > 0) {
    const porProyecto = importeEUR / lista.length;
    const conTokens = lista.map((p) => {
      const costeTokenEUR = p.divisa === '€' ? p.pxEmision : p.pxEmision / fx;
      return { p, costeTokenEUR, tokens: Math.floor(porProyecto / costeTokenEUR) };
    });

    const viables = conTokens.filter((x) => x.tokens > 0);
    if (viables.length === lista.length) {
      const asignado = suma(viables.map((x) => x.tokens * x.costeTokenEUR));
      return {
        seleccion: viables.map((x) => ({ proyecto: x.p, tokens: x.tokens })),
        noAsignadoEUR: importeEUR - asignado,
      };
    }
    // alguno no llega a un token: reducimos la cesta y reintentamos
    lista = viables.map((x) => x.p);
  }

  return { seleccion: [], noAsignadoEUR: importeEUR };
}

/* ------------------------------------------------------------
 * 5. ORQUESTADOR
 * ------------------------------------------------------------ */

/**
 * Simula una cartera de principio a fin.
 *
 * modoImporte:
 *   'inmobiliario' -> importe es solo para inmuebles, el estatus va aparte
 *   'total'        -> importe incluye el estatus, que se descuenta primero
 */
function simular(universo, opciones) {
  const {
    objetivo,
    mercados = [],
    estatus,
    importe,
    modoImporte = 'inmobiliario',
    costeEstatusEUR = 0,
    fx,
    maxProyectos = 5,
    minMeses = MIN_MESES,
    tasaReinversionAnual = 0,
    tasaStakingAnual = 0,
  } = opciones;

  const importeInmobiliario =
    modoImporte === 'total' ? importe - costeEstatusEUR : importe;

  if (importeInmobiliario <= 0) {
    return { error: 'El coste del estatus consume todo el importe disponible.' };
  }

  const candidatos = seleccionarProyectos(universo, {
    objetivo,
    mercados,
    estatus,
    minMeses,
  });

  if (candidatos.length === 0) {
    return { error: 'Ningún proyecto disponible cumple esos criterios.' };
  }

  const { seleccion, noAsignadoEUR } = repartirImporte(candidatos, importeInmobiliario, {
    fx,
    maxProyectos,
  });

  if (seleccion.length === 0) {
    return { error: 'El importe no alcanza para adquirir ningún token.' };
  }

  const items = seleccion.map(({ proyecto, tokens }) =>
    calcularInmueble(proyecto, { tokens, pxPagado: proyecto.pxEmision, estatus, fx })
  );

  const cartera = calcularCartera(items, { fx, costeEstatusEUR });
  const reinversion = calcularReinversion(cartera, {
    tasaReinversionAnual,
    tasaStakingAnual,
    valorEstatusEUR: costeEstatusEUR,
  });

  return { cartera, reinversion, noAsignadoEUR, candidatosDescartados: candidatos.length - seleccion.length };
}

module.exports = {
  ESTATUS,
  ESTADOS_EXCLUIDOS,
  MIN_MESES,
  OBJETIVOS,
  HORIZONTES,
  calcularInmueble,
  calcularCartera,
  calcularReinversion,
  seleccionarProyectos,
  repartirImporte,
  simular,
};
