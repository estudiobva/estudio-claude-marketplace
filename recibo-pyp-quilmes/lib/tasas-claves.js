// tasas-claves.js — numeros de cuenta de Publicidad y Propaganda (PyP) de
// Quilmes, de la planilla "Tasas Municipales CLAVES.xlsx".
//
// La planilla esta en la carpeta del estudio de la unidad compartida
// ("0 - ESTUDIO - 00000000000/Tasas Municipales CLAVES.xlsx"), o donde diga
// TASAS_MUNICIPALES_PATH. Hoja QUILMES, tres bloques lado a lado:
//   SEGURIDAD E HIGIENE | MARQUESINAS | PUBLICIDAD Y PROPAGANDA
// cada uno con Nombre | Cuit | Nº de Cuenta. Solo se usa el de PUBLICIDAD Y
// PROPAGANDA: el de Marquesinas a veces trae otro numero y el de Seguridad e
// Higiene es otra tasa.
//
// En este bloque las cuentas van sin digito verificador (ej. 500123456), igual que
// el desplegable de ARQui; si alguien la carga con guion se le saca.
// Los nombres no son confiables (abreviados, con espacios raros, o distintos
// para el mismo CUIT en cada bloque): se busca por CUIT.

const fs = require('fs');
const path = require('path');
const { readWorkbook } = require('./xlsx-min');
const rutas = require('./rutas-bva');

const NOMBRE = 'Tasas Municipales CLAVES.xlsx';
const norm = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
const normCuit = (s) => String(s || '').replace(/\D/g, '');
const normCuenta = (s) => String(s || '').trim().replace(/-\d$/, '').replace(/\D/g, '');

function archivo() {
  const env = (process.env.TASAS_MUNICIPALES_PATH || '').trim();
  if (env) {
    const p = fs.existsSync(env) && fs.statSync(env).isDirectory() ? path.join(env, NOMBRE) : env;
    if (!fs.existsSync(p)) throw new Error(`TASAS_MUNICIPALES_PATH apunta a ${p}, que no existe.`);
    return p;
  }
  const raiz = rutas.raiz();
  const estudio = fs.readdirSync(raiz).find((d) => /^0 - ESTUDIO/i.test(d));
  const p = estudio && path.join(raiz, estudio, NOMBRE);
  if (!p || !fs.existsSync(p)) {
    throw new Error(`No encontre "${NOMBRE}" en la carpeta del estudio de la unidad. Defini TASAS_MUNICIPALES_PATH en el .env.`);
  }
  return p;
}

// [{ nombre, cuit, cuenta, fila }] del bloque PUBLICIDAD Y PROPAGANDA de QUILMES.
// Las filas con CUIT o cuenta vacios / #N/A se devuelven con `problema`.
function cuentasPyp() {
  const file = archivo();
  const wb = readWorkbook(fs.readFileSync(file));
  const hoja = wb.sheetNames.find((n) => norm(n) === 'QUILMES');
  if (!hoja) throw new Error(`${NOMBRE} no tiene la hoja QUILMES.`);
  const rows = wb.rows(hoja);

  let filaTit = -1, col = -1;
  for (let i = 0; i < rows.length && col < 0; i++) {
    const j = rows[i].findIndex((c) => norm(c) === 'PUBLICIDAD Y PROPAGANDA');
    if (j >= 0) { filaTit = i; col = j; }
  }
  if (col < 0) throw new Error(`No encontre el bloque "PUBLICIDAD Y PROPAGANDA" en la hoja ${hoja} de ${NOMBRE}.`);

  // Encabezados debajo del titulo: Nombre | Cuit | Nº de Cuenta (se ubican por texto).
  const enc = rows[filaTit + 1] || [];
  const buscar = (re) => {
    for (let j = col; j < col + 4; j++) if (re.test(norm(enc[j]))) return j;
    return -1;
  };
  const cNom = buscar(/NOMBRE|FARMACIA/), cCuit = buscar(/CUIT/), cCta = buscar(/CUENTA/);
  if ([cNom, cCuit, cCta].includes(-1)) throw new Error(`El bloque PUBLICIDAD Y PROPAGANDA de ${NOMBRE} no tiene las columnas Nombre / Cuit / Nº de Cuenta.`);

  const out = [];
  for (let i = filaTit + 2; i < rows.length; i++) {
    const r = rows[i];
    const nombre = String(r[cNom] || '').replace(/\s+/g, ' ').trim();
    const cuitRaw = String(r[cCuit] || '').trim();
    const ctaRaw = String(r[cCta] || '').trim();
    if (!nombre && !cuitRaw && !ctaRaw) continue;
    const item = { nombre, cuit: normCuit(cuitRaw), cuenta: normCuenta(ctaRaw), fila: i + 1 };
    if (item.cuit.length !== 11) item.problema = `CUIT "${cuitRaw || 'vacio'}"`;
    else if (item.cuenta.length < 6) item.problema = `cuenta "${ctaRaw || 'vacia'}"`;
    out.push(item);
  }
  return { file, filas: out };
}

module.exports = { cuentasPyp, normCuenta, archivo, NOMBRE };
