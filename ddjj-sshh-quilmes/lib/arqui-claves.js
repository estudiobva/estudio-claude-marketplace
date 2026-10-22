// arqui-claves.js — usuario y clave de ARQui (Quilmes) para los scripts.
//
// En Claves_Organismos.xlsx el acceso a ARQui esta en el grupo MUNICIPAL de
// cada sociedad (Usuario | Clave | Municipio = Quilmes). Casi todas comparten
// el mismo usuario del estudio; unas pocas tienen uno propio. Se busca, en orden:
//   1. ARQUI_USUARIO / ARQUI_CLAVE en el entorno o el .env (override puntual).
//   2. La fila de la sociedad (por CUIT), grupo MUNICIPAL con municipio Quilmes.
//   3. El usuario de Quilmes que mas se repite en la planilla (el general).
// Nunca se imprime la clave.

const { loadEnv } = require('./env');
const claves = require('./claves');

const esQuilmes = (r) => {
  const m = r.organismos.MUNICIPAL;
  return Boolean(m && m.usuario && m.clave && /quilmes/i.test(m.municipio || ''));
};

// Todos los accesos de Quilmes distintos de la planilla, el mas usado primero:
// [{ usuario, clave, cuits: [...] }]
function accesosQuilmes() {
  const { rows } = claves.loadSync();
  const pares = new Map();
  for (const r of rows.filter(esQuilmes)) {
    const { usuario, clave } = r.organismos.MUNICIPAL;
    const k = `${usuario}\u0000${clave}`;
    if (!pares.has(k)) pares.set(k, { usuario, clave, cuits: [] });
    if (r.cuit) pares.get(k).cuits.push(r.cuit);
  }
  return [...pares.values()].sort((a, b) => b.cuits.length - a.cuits.length);
}

function credencialesArqui({ cuit } = {}) {
  loadEnv();
  const envUser = (process.env.ARQUI_USUARIO || '').trim();
  const envPass = process.env.ARQUI_CLAVE || '';
  if (envUser && envPass) return { usuario: envUser, clave: envPass, origen: 'env' };

  if (!claves.filePath()) {
    throw new Error('No encontre Claves_Organismos.xlsx ni ARQUI_USUARIO/ARQUI_CLAVE en el .env.');
  }
  const c = claves.normCuit(cuit);
  if (c) {
    const { rows } = claves.loadSync();
    const propia = rows.find((r) => r.cuit === c && esQuilmes(r));
    if (propia) return { usuario: propia.organismos.MUNICIPAL.usuario, clave: propia.organismos.MUNICIPAL.clave, origen: 'archivo' };
  }
  const [general] = accesosQuilmes();
  if (!general) {
    throw new Error('No hay ningun usuario de ARQui en Claves_Organismos.xlsx (grupo MUNICIPAL con municipio Quilmes).');
  }
  return { usuario: general.usuario, clave: general.clave, origen: 'archivo (general de Quilmes)' };
}

module.exports = { credencialesArqui, accesosQuilmes };
