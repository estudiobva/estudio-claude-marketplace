// zetti-claves.js — credenciales de T&S Web (Zetti) para los scripts.
//
// Zetti usa un solo usuario del estudio para todas las sociedades, asi que no va
// por CUIT como ARCA. Se busca, en orden:
//   1. ZETTI_USUARIO / ZETTI_CLAVE en el entorno o el .env (override puntual).
//   2. Claves_Organismos.xlsx, grupo de columnas "ZETTI" (Usuario | Clave):
//      - la fila de la sociedad, si tiene ZETTI propio;
//      - si no, el unico par usuario/clave cargado en la planilla (la fila
//        "Zetti T&S Web (todas las sociedades)").
// Nunca se imprime la clave.

const { loadEnv } = require('./env');
const claves = require('./claves');

function credencialesZetti({ cuit } = {}) {
  loadEnv();
  const envUser = (process.env.ZETTI_USUARIO || '').trim();
  const envPass = process.env.ZETTI_CLAVE || '';
  if (envUser && envPass) return { usuario: envUser, clave: envPass, origen: 'env' };

  if (!claves.filePath()) {
    throw new Error('No encontre Claves_Organismos.xlsx ni ZETTI_USUARIO/ZETTI_CLAVE en el .env.');
  }
  const { rows, file } = claves.loadSync();
  const conZetti = rows.filter((r) => r.organismos.ZETTI && r.organismos.ZETTI.usuario && r.organismos.ZETTI.clave);

  const c = claves.normCuit(cuit);
  const propia = c && conZetti.find((r) => r.cuit === c);
  if (propia) return { usuario: propia.organismos.ZETTI.usuario, clave: propia.organismos.ZETTI.clave, origen: 'archivo' };

  // La fila general se reconoce por el titular ("Zetti T&S Web (todas las
  // sociedades)"). Otras filas sin CUIT pueden tener algo en estas columnas por
  // ser las ultimas de la planilla (paso el 08/10/2026 con un acceso municipal):
  // no se toman como usuario de Zetti.
  const general = conZetti.filter((r) => !r.cuit && /zetti/i.test(r.titular));
  if (general.length === 1) {
    return { usuario: general[0].organismos.ZETTI.usuario, clave: general[0].organismos.ZETTI.clave, origen: 'archivo' };
  }

  const pares = new Map(conZetti.map((r) => [`${r.organismos.ZETTI.usuario}\u0000${r.organismos.ZETTI.clave}`, r]));
  if (pares.size === 1) {
    const r = [...pares.values()][0];
    return { usuario: r.organismos.ZETTI.usuario, clave: r.organismos.ZETTI.clave, origen: 'archivo' };
  }
  if (pares.size === 0) {
    throw new Error(
      `No hay usuario y clave de Zetti en ${file}: completa las columnas ZETTI (Usuario | Clave) ` +
      'de la fila "Zetti T&S Web (todas las sociedades)".'
    );
  }
  throw new Error(
    `Hay ${pares.size} usuarios de Zetti distintos en ${file} y la sociedad ${c || '(sin CUIT)'} no tiene uno propio. ` +
    'Deja uno solo general o carga el de la sociedad en su fila.'
  );
}

module.exports = { credencialesZetti };
