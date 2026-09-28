// claves-bancos.js — credenciales de home banking desde Claves_Bancos.xlsx.
//
// El archivo esta CIFRADO con contraseña (igual que Claves_Organismos.xlsx),
// asi que no se puede abrir con un parser comun: hay que descifrarlo antes con
// officecrypto-tool.
//
//   CLAVES_BANCOS_PATH      ruta al .xlsx (default: la de la unidad compartida)
//   CLAVES_BANCOS_PASSWORD  contraseña (si falta, se usa CLAVES_ORGANISMOS_PASSWORD:
//                           hoy es la misma en el estudio)
//
// Estructura real de la hoja BANCOS (relevada el 21/09/2026):
//
//   fila 1      BANCO | FARMACIAS | USUARIO | ACCESO/DNI | CLAVE ID | LINKS
//   filas 2-10  un banco por fila, login UNICO para todas las farmacias:
//               COMAFI, PATAGONIA, PROVINCIA, HIPOTECARIO, NACION, FRANCES,
//               MACRO, SANTANDER, GALICIA
//   fila 11     BANCO GALICIA | Sociedades | USUARIO | CONTRASEÑA   <- 2do encabezado
//   filas 12+   una por sociedad, con el login propio de Galicia
//
// Las columnas del segundo bloque NO estan alineadas con las etiquetas de su
// propio encabezado, asi que aca NUNCA se lee por indice fijo: se mapea por
// etiqueta y, si no se puede, se reporta en vez de devolver la columna
// equivocada como si fuera la contraseña.
//
// Este modulo NO imprime credenciales. Para mirar el archivo esta
// scripts/claves-bancos.js, que enmascara.

const fs   = require('fs');
const path = require('path');

const UNIDAD_DEFECTO = 'I:\\Unidades compartidas\\BVA - Sociedades Farmaceuticas';
const NOMBRE = 'Claves_Bancos.xlsx';
const HOJA = 'BANCOS';

function filePath() {
  const p = process.env.CLAVES_BANCOS_PATH && process.env.CLAVES_BANCOS_PATH.trim();
  if (p) {
    const r = p.replace(/^~/, require('os').homedir());
    return fs.statSync(r).isDirectory() ? path.join(r, NOMBRE) : r;
  }
  let base;
  try { base = require('./rutas-bva').raiz(); } catch { base = UNIDAD_DEFECTO; }
  return path.join(base, '0 - ESTUDIO - 00000000000', NOMBRE);
}

function password() {
  return process.env.CLAVES_BANCOS_PASSWORD || process.env.CLAVES_ORGANISMOS_PASSWORD || '';
}

const txt = (c) => String(c == null ? '' : c).trim();
const norm = (s) => txt(s).toUpperCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/[^A-Z0-9]+/g, ' ').trim();

// Devuelve { indice: {ETIQUETA: col} } a partir de una fila de encabezado.
function mapearEncabezado(fila) {
  const m = {};
  fila.forEach((c, i) => { const k = norm(c); if (k) m[k] = i; });
  return m;
}

// Busca una columna por varias etiquetas posibles.
function columna(mapa, ...etiquetas) {
  for (const e of etiquetas) {
    const k = norm(e);
    if (mapa[k] != null) return mapa[k];
    const parcial = Object.keys(mapa).find(x => x.includes(k));
    if (parcial) return mapa[parcial];
  }
  return null;
}

async function leerFilas() {
  const file = filePath();
  if (!fs.existsSync(file)) {
    throw new Error(`No encuentro ${file}. Defini CLAVES_BANCOS_PATH en el .env.`);
  }
  let oct;
  try { oct = require('officecrypto-tool'); }
  catch { throw new Error('Falta la dependencia officecrypto-tool (npm install).'); }

  const raw = fs.readFileSync(file);
  let buf = raw;
  if (oct.isEncrypted(raw)) {
    const pass = password();
    if (!pass) {
      throw new Error(
        `${NOMBRE} esta cifrado y no hay contraseña. Agrega CLAVES_BANCOS_PASSWORD ` +
        `al .env (o CLAVES_ORGANISMOS_PASSWORD, que hoy es la misma).`);
    }
    try { buf = await oct.decrypt(raw, { password: pass }); }
    catch (e) {
      throw new Error(`No pude descifrar ${NOMBRE}: contraseña incorrecta o formato no soportado.`);
    }
  }
  const { readWorkbook } = require('./xlsx-min');
  const wb = readWorkbook(buf);
  if (!wb.sheetNames.includes(HOJA)) {
    throw new Error(`${NOMBRE} no tiene la hoja "${HOJA}". Hojas: ${wb.sheetNames.join(', ')}.`);
  }
  return wb.rows(HOJA).filter(f => f.some(c => txt(c)));
}

// Parsea los dos bloques. Devuelve credenciales SIN imprimirlas.
async function load() {
  const filas = await leerFilas();

  const encB1 = mapearEncabezado(filas[0]);
  const cBanco = columna(encB1, 'BANCO');
  const cFarm  = columna(encB1, 'FARMACIAS');
  const cUsr1  = columna(encB1, 'USUARIO');
  const cAcc1  = columna(encB1, 'ACCESO DNI', 'ACCESO');
  const cClv1  = columna(encB1, 'CLAVE ID', 'CLAVE');
  const cLink1 = columna(encB1, 'LINKS', 'LINK');
  if (cBanco == null || cUsr1 == null || cClv1 == null) {
    throw new Error('La hoja BANCOS no tiene las columnas BANCO/USUARIO/CLAVE esperadas.');
  }

  // El bloque de Galicia arranca en su propia fila de encabezado.
  const iGal = filas.findIndex((f, i) =>
    i > 0 && /BANCO GALICIA/i.test(txt(f[0])) && f.some(c => /sociedad/i.test(txt(c))));

  const fin1 = iGal > 0 ? iGal : filas.length;
  const bancos = filas.slice(1, fin1).map(f => ({
    banco: txt(f[cBanco]),
    farmacias: cFarm != null ? txt(f[cFarm]) : '',
    usuario: txt(f[cUsr1]),
    acceso: cAcc1 != null ? txt(f[cAcc1]) : '',
    clave: txt(f[cClv1]),
    link: cLink1 != null ? txt(f[cLink1]) : '',
  })).filter(b => b.banco);

  let galicia = [];
  let avisoGalicia = null;
  if (iGal > 0) {
    const encB2 = mapearEncabezado(filas[iGal]);
    const cSoc = columna(encB2, 'SOCIEDADES', 'SOCIEDAD');
    const cUsr = columna(encB2, 'USUARIO');
    const cClv = columna(encB2, 'CONTRASENA', 'CONTRASEÑA', 'CLAVE');

    const datos = filas.slice(iGal + 1);
    // Las etiquetas del 2do encabezado no siempre caen sobre la columna de sus
    // datos (el nombre de la sociedad aparece en la col 0 aunque la etiqueta
    // "Sociedades" este en la 1). Se valida y, si no cierra, se avisa en vez de
    // devolver una columna cualquiera como contraseña.
    const colSoc = (cSoc != null && datos.some(f => txt(f[cSoc]))) ? cSoc : 0;
    if (cUsr == null || cClv == null) {
      avisoGalicia = 'No pude ubicar las columnas USUARIO/CONTRASEÑA del bloque Galicia ' +
                     'por su encabezado. Revisa el archivo: no adivino cual es la clave.';
    } else {
      galicia = datos.map(f => ({
        sociedad: txt(f[colSoc]),
        // CUIT(s) de la fila, si los trae (al 24/09/2026 el bloque Galicia NO trae CUIT:
        // la busqueda cae en nombre exacto o en lib/galicia-alias.json).
        cuits: [...new Set(f.map((c, i) => (i === cUsr || i === cClv) ? '' : txt(c).replace(/\D/g, ''))
                           .filter(d => d.length === 11))],
        usuario: txt(f[cUsr]),
        clave: txt(f[cClv]),
      })).filter(g => g.sociedad);
    }
  }

  return { bancos, galicia, avisoGalicia, archivo: filePath() };
}

// Equivalencias CUIT -> nombre en el bloque Galicia. Se buscan en lib/ (checkout
// local) y en ~/.fisco-ar/ (plugin instalado: sobrevive a las actualizaciones).
// Tienen datos de clientes: el archivo nunca se sube al repo.
function aliasGalicia() {
  for (const f of [path.join(__dirname, 'galicia-alias.json'),
                   path.join(require('os').homedir(), '.fisco-ar', 'galicia-alias.json')]) {
    try { return JSON.parse(fs.readFileSync(f, 'utf-8')); } catch { /* siguiente */ }
  }
  return {};
}

// Credenciales para un banco y (si aplica) una sociedad.
//
// Galicia tiene login por sociedad; el resto usa el login unico del banco.
async function buscar({ banco, sociedad, cuit } = {}) {
  const { bancos, galicia, avisoGalicia } = await load();
  const b = norm(banco);
  if (!b) throw new Error('Falta el banco.');

  if (/GALICIA/.test(b) && (sociedad || cuit)) {
    if (avisoGalicia) throw new Error(avisoGalicia);
    // Primero por CUIT (definitivo). Por nombre, SOLO coincidencia exacta: un
    // parecido (nombre corto contenido en el nombre largo de otra) entraria
    // con el usuario de OTRA sociedad.
    const c = String(cuit || '').replace(/\D/g, '');
    let cand = c ? galicia.filter(g => g.cuits.includes(c)) : [];
    let por = 'CUIT';
    if (!cand.length && sociedad) { cand = galicia.filter(g => norm(g.sociedad) === norm(sociedad)); por = 'nombre'; }
    // Equivalencia explicita CUIT -> nombre en el archivo (lib/galicia-alias.json).
    if (!cand.length && c) {
      const alias = aliasGalicia()[c];
      if (alias && !alias.confirmado) {
        throw new Error(`"${sociedad || c}" figura distinto en el bloque Galicia; hay una equivalencia propuesta ` +
          `("${alias.nombreEnClaves}") SIN CONFIRMAR en lib/galicia-alias.json. Confirmala mirando el archivo.`);
      }
      if (alias) { cand = galicia.filter(g => norm(g.sociedad) === norm(alias.nombreEnClaves)); por = 'equivalencia'; }
    }
    if (cand.length === 1) {
      const g = cand[0];
      if (!g.usuario || !g.clave) {
        throw new Error(`La fila de "${g.sociedad}" en el bloque Galicia no tiene usuario o clave cargados.`);
      }
      return { banco: 'GALICIA', sociedad: g.sociedad, usuario: g.usuario, clave: g.clave,
               origen: 'galicia-por-sociedad', encontradaPor: por };
    }
    if (cand.length > 1) {
      throw new Error(`${c ? `El CUIT ${c}` : `"${sociedad}"`} aparece en ${cand.length} filas del bloque Galicia: ` +
                      cand.map(x => x.sociedad).join(' | ') + '. Corregilo en el archivo.');
    }
    throw new Error(`No encontre ${c ? `el CUIT ${c}` : ''}${c && sociedad ? ' ni ' : ''}${sociedad ? `"${sociedad}" (nombre exacto)` : ''} ` +
                    `en el bloque Galicia de ${NOMBRE}.`);
  }

  const cand = bancos.filter(x => norm(x.banco) === b || norm(x.banco).includes(b));
  if (!cand.length) {
    throw new Error(`No encontre el banco "${banco}". Disponibles: ${bancos.map(x => x.banco).join(', ')}.`);
  }
  const hit = cand.find(x => !sociedad || /TODAS/i.test(x.farmacias) ||
                             norm(x.farmacias).includes(norm(sociedad))) || cand[0];
  if (!hit.usuario || !hit.clave) {
    throw new Error(`La fila de ${hit.banco} no tiene usuario o clave cargados.`);
  }
  return { banco: hit.banco, farmacias: hit.farmacias, usuario: hit.usuario,
           acceso: hit.acceso, clave: hit.clave, link: hit.link, origen: 'banco-login-unico' };
}

module.exports = { filePath, load, buscar, NOMBRE, HOJA };
