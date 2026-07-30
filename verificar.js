/* ============================================================
 * VERIFICACIÓN DEL MOTOR
 *
 * Reconstruye la cartera exacta que hay guardada en
 * Plantilla_Propuesta_ES_Cartera_v4.xlsx y compara cada resultado
 * con el valor cacheado en el sheet.
 *
 * Los datos de los 4 inmuebles se han extraído de la pestaña
 * "Cálculos". Los de RET-3 no estaban explícitos y se han
 * deducido de las filas 95-99 (rentabilidad 7% sobre precio de
 * emisión 100 pagando 99, sin plusvalía y sin mejora por estatus).
 * ============================================================ */

'use strict';

const M = require('./motor');

const FX = 1.145405; // [D9] GOOGLEFINANCE("CURRENCY:EURUSD")

const CARTERA_PLANTILLA = [
  {
    id: 'DNB-1',
    nombre: 'Dania Beach 1',
    ubicacion: 'USA',
    tipologiaExplotacion: 'Préstamo promotor',
    tipologiaDividendos: 'rendimientos trimestrales + final',
    estado: 'EN EXPLOTACIÓN',
    divisa: '$',
    pxEmision: 100,
    mesesRestantes: 38,
    rentAlquiler: {
      Reentel: 0.0629688189,
      ReentelPro: 0.0629688189,
      SuperReentel: 0.0629688189,
    },
    plusvalia: { Reentel: 0.12, ReentelPro: 0.24, SuperReentel: 0.32 },
    tokens: 100,
  },
  {
    id: 'MRB-1',
    nombre: 'Marbella 1',
    ubicacion: 'España',
    tipologiaExplotacion: 'Préstamo promotor',
    tipologiaDividendos: 'rendimientos mensuales',
    estado: 'EN EXPLOTACIÓN',
    divisa: '€',
    pxEmision: 100,
    mesesRestantes: 4,
    rentAlquiler: {
      Reentel: 0.1030017683,
      ReentelPro: 0.1081144082,
      SuperReentel: 0.1277579317,
    },
    plusvalia: { Reentel: 0, ReentelPro: 0, SuperReentel: 0 },
    tokens: 100,
  },
  {
    id: 'OPO-2',
    nombre: 'Opportunity 2',
    ubicacion: 'Global',
    tipologiaExplotacion: 'Préstamo promotor',
    tipologiaDividendos: 'rendimientos mensuales + final',
    estado: 'EN EXPLOTACIÓN',
    divisa: '€',
    pxEmision: 100,
    mesesRestantes: 22,
    rentAlquiler: { Reentel: 0, ReentelPro: 0, SuperReentel: 0 },
    plusvalia: { Reentel: 0.04, ReentelPro: 0.08, SuperReentel: 0.14 },
    tokens: 400,
  },
  {
    id: 'RET-3',
    nombre: 'Rentas 3',
    ubicacion: 'Global',
    tipologiaExplotacion: 'Préstamo promotor',
    tipologiaDividendos: 'rendimientos mensuales',
    estado: 'EN EXPLOTACIÓN',
    divisa: '$',
    pxEmision: 100,
    mesesRestantes: 36.0328767,
    rentAlquiler: { Reentel: 0.07, ReentelPro: 0.07, SuperReentel: 0.07 },
    plusvalia: { Reentel: 0, ReentelPro: 0, SuperReentel: 0 },
    tokens: 400,
    pxPagado: 99, // descuento aplicado en la plantilla
  },
];

/* --- valores esperados, leídos del sheet ------------------- */

const ESPERADO = {
  nInmuebles: [4, 'D47'],
  totalTokens: [1000, 'E47'],
  mediaMeses: [27.41315068, 'F47'],
  invEUR: [93303.46035, 'H47'],
  invUSD: [106870.25, 'K47'],
  pctEUR: [0.5358858055, 'G47'],

  'Reentel.rentAlquiler': [0.04487988701, 'D62'],
  'Reentel.plusvalia': [0.03204040404, 'G62'],
  'Reentel.rentTotal': [0.1403398944, 'J62'],
  'Reentel.rentAnualizada': [0.05874220593, 'M62'],
  'Reentel.soloAlquiler': [0.07549827724, 'D63'],
  'Reentel.soloFinal': [0.03597552617, 'G63'],

  'SuperReentel.rentAlquiler': [0.04735550334, 'F62'],
  'SuperReentel.plusvalia': [0.09204040404, 'I62'],
  'SuperReentel.rentTotal': [0.2011650998, 'L62'],
  'SuperReentel.rentAnualizada': [0.08935179356, 'O62'],
  'SuperReentel.soloAlquiler': [0.08014265928, 'F63'],
  'SuperReentel.soloFinal': [0.1049535404, 'I63'],

  'SuperReentel.costeEstatusEUR': [7323.872342, 'E52'],
  'SuperReentel.costeEstatusUSD': [8388.8, 'E51'],
  'SuperReentel.totalEUR': [100627.3327, 'E57'],
  'SuperReentel.totalUSD': [115259.05, 'E56'],
};

const ESPERADO_REINV = {
  Reentel: {
    tasa: 0.11,
    valorEstatus: 0,
    valorFinal: [95367.60583, 97558.82059, 104684.2898, 114896.9013, 141992.8709],
    ganancia: [2064.145483, 4255.360246, 11380.82946, 21593.4409, 48689.41054],
    rentAcumulada: [0.0221229253, 0.04560774305, 0.1219764992, 0.2314323694, 0.5218392796],
  },
  SuperReentel: {
    tasa: 0.16,
    valorEstatus: 7323.872342,
    valorFinal: [95554.30974, 98079.64969, 110598.3609, 125340.4058, 171847.4235],
    ganancia: [2250.849394, 4776.189339, 17294.90058, 32036.94547, 78543.96316],
    rentAcumulada: [0.02412396481, 0.05118984142, 0.1853618346, 0.3433628866, 0.8418118992],
    staking: [168.4490639, 336.8981277, 673.7962555, 1010.694383, 1684.490639],
    rentAcumuladaConStaking: [
      0.02404216025, 0.05081211367, 0.1785667607, 0.3284161368, 0.7972829215,
    ],
  },
};

const TASA_STAKING = 0.046; // [D21]

/* --- comparación ------------------------------------------ */

let fallos = 0;
let pruebas = 0;

function comparar(etiqueta, obtenido, esperado, celda, tol = 1e-7) {
  pruebas++;
  const escala = Math.max(1, Math.abs(esperado));
  const desvio = Math.abs(obtenido - esperado) / escala;
  const ok = desvio < tol;
  if (!ok) fallos++;
  const marca = ok ? '  ok  ' : ' FALLO';
  console.log(
    `${marca} ${etiqueta.padEnd(38)} ${fmt(obtenido).padStart(16)}  esperado ${fmt(
      esperado
    ).padStart(16)}  [${celda}]`
  );
}

function fmt(n) {
  if (Math.abs(n) >= 1000) return n.toFixed(4);
  return n.toFixed(9);
}

function construir(estatus, costeEstatusEUR) {
  const items = CARTERA_PLANTILLA.map((inm) =>
    M.calcularInmueble(inm, {
      tokens: inm.tokens,
      pxPagado: inm.pxPagado != null ? inm.pxPagado : inm.pxEmision,
      estatus,
      fx: FX,
    })
  );
  return M.calcularCartera(items, { fx: FX, costeEstatusEUR });
}

console.log('\n=== CARTERA (independiente del estatus) ===\n');
const cReentel = construir('Reentel', 0);
comparar('nInmuebles', cReentel.nInmuebles, ...ESPERADO.nInmuebles);
comparar('totalTokens', cReentel.totalTokens, ...ESPERADO.totalTokens);
comparar('mediaMeses (ponderada)', cReentel.mediaMeses, ...ESPERADO.mediaMeses);
comparar('inversión €', cReentel.invEUR, ...ESPERADO.invEUR);
comparar('inversión $', cReentel.invUSD, ...ESPERADO.invUSD);
comparar('% cartera en €', cReentel.pctEUR, ...ESPERADO.pctEUR);

console.log('\n=== RENTABILIDADES POR ESTATUS ===\n');
const carteras = {
  Reentel: cReentel,
  SuperReentel: construir('SuperReentel', ESPERADO_REINV.SuperReentel.valorEstatus),
};

for (const est of ['Reentel', 'SuperReentel']) {
  const c = carteras[est];
  for (const campo of [
    'rentAlquiler',
    'plusvalia',
    'rentTotal',
    'rentAnualizada',
    'soloAlquiler',
    'soloFinal',
  ]) {
    const clave = `${est}.${campo}`;
    if (ESPERADO[clave]) comparar(clave, c[campo], ...ESPERADO[clave]);
  }
}

console.log('\n=== COSTE DEL ESTATUS Y TOTALES ===\n');
const cSR = carteras.SuperReentel;
comparar('SR.costeEstatusEUR', cSR.costeEstatusEUR, ...ESPERADO['SuperReentel.costeEstatusEUR']);
comparar('SR.costeEstatusUSD', cSR.costeEstatusUSD, ...ESPERADO['SuperReentel.costeEstatusUSD']);
comparar('SR.totalEUR', cSR.totalEUR, ...ESPERADO['SuperReentel.totalEUR']);
comparar('SR.totalUSD', cSR.totalUSD, ...ESPERADO['SuperReentel.totalUSD']);

console.log('\n=== REINVERSIÓN (6 / 12 / 24 / 36 / 60 meses) ===\n');
for (const est of ['Reentel', 'SuperReentel']) {
  const exp = ESPERADO_REINV[est];
  const r = M.calcularReinversion(carteras[est], {
    tasaReinversionAnual: exp.tasa,
    tasaStakingAnual: TASA_STAKING,
    valorEstatusEUR: exp.valorEstatus,
  });
  r.forEach((f, k) => {
    comparar(`${est} valorFinal H=${f.horizonte}`, f.valorFinal, exp.valorFinal[k], 'D75:H75');
    comparar(`${est} ganancia H=${f.horizonte}`, f.ganancia, exp.ganancia[k], 'D76:H76');
    comparar(`${est} rentAcum H=${f.horizonte}`, f.rentAcumulada, exp.rentAcumulada[k], 'D77:H77');
    if (exp.staking) {
      comparar(`${est} staking H=${f.horizonte}`, f.staking, exp.staking[k], 'D91:H91');
      comparar(
        `${est} rentAcum+stk H=${f.horizonte}`,
        f.rentAcumuladaConStaking,
        exp.rentAcumuladaConStaking[k],
        'D92:H92'
      );
    }
  });
}

console.log('\n=== DISTRIBUCIONES ===\n');
const dist = cReentel.distribucion;
comparar('ubicación: USA', dist.ubicacion['USA'], 0.1, 'N47');
comparar('ubicación: España', dist.ubicacion['España'], 0.1, 'P47');
comparar('ubicación: Global', dist.ubicacion['Global'], 0.8, 'Q47');
comparar('divid.: mensuales + final', dist.dividendos['rendimientos mensuales + final'], 0.4, 'T47');
comparar('divid.: mensuales', dist.dividendos['rendimientos mensuales'], 0.5, 'U47');
comparar('divid.: trimestrales + final', dist.dividendos['rendimientos trimestrales + final'], 0.1, 'W47');

console.log(
  `\n${'-'.repeat(72)}\n${pruebas - fallos} de ${pruebas} pruebas correctas` +
    (fallos ? `  —  ${fallos} FALLOS\n` : '  —  motor validado\n')
);

process.exit(fallos ? 1 : 0);
