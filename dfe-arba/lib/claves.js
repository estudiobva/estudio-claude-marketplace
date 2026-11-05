// ─────────────────────────────────────────────────────────────────────────────
// Vendorizado desde el plugin fisco-ar (MIT) — https://github.com/javiergradiche/fisco-ar-claude-plugin
// Commit de origen: 2ad06df · Copiado el 21/09/2026 al repo del Estudio BVA.
//
// Se copia en vez de depender del plugin para que una actualizacion del
// marketplace no pueda romper un cierre. Si ARCA cambia el login, el arreglo va
// ACA (y conviene mirar upstream por si ya lo resolvieron).
// ─────────────────────────────────────────────────────────────────────────────

// claves.js — el archivo Excel de claves como fuente de credenciales.
//
// Alternativa (y complemento) al .env: en vez de mantener a mano un par de
// variables por CUIT, se lee la planilla que el estudio ya usa —
// `Claves_Organismos.xlsx`, protegida con contraseña— y de ahí salen el CUIT de
// login y la clave fiscal de cada titular.
//
// Configuración (en el .env, o exportadas en la shell):
//   CLAVES_ORGANISMOS_PATH=/ruta/…   dónde está el archivo. Acepta la ruta al
//                                    .xlsx o la carpeta que lo contiene, y
//                                    entiende `~`. Si no está definida se busca
//                                    Claves_Organismos.xlsx en <raíz del
//                                    plugin>/ y en ~/.fisco-ar/.
//   CLAVES_ORGANISMOS_PASSWORD=...   contraseña con la que se abre el Excel
//
// Formato esperado de la planilla (se detecta solo, no está hardcodeado):
//   fila con los organismos agrupados   → ...  | ARCA            | ARBA | AGIP | MUNICIPAL
//   fila de encabezados de columna      → # | Titular | CUIT | … | Usuario | Clave | …
//   filas de datos                      → 10 | Rivadavia SRL | 30712345678 | … | 20123456789 | **** | …
//
// Los encabezados se leen de la planilla, así que agregar un organismo nuevo o
// mover columnas no rompe nada mientras se respete el esquema de dos filas.

const fs   = require('fs');
const path = require('path');
const os   = require('os');
const { execFileSync } = require('child_process');
const { readWorkbook } = require('./xlsx-min');

const PLUGIN_ROOT = path.join(__dirname, '..', '..');
const FILE_NAME   = 'Claves_Organismos.xlsx';

// ── Ubicación del archivo ─────────────────────────────────────────────────────

// `~/x` → `/Users/…/x`. El .env se edita a mano y la tilde es lo natural ahí.
function expandHome(p) {
  const s = String(p || '').trim();
  if (!s) return '';
  if (s === '~') return os.homedir();
  if (s.startsWith('~/')) return path.join(os.homedir(), s.slice(2));
  return s;
}

// CLAVES_ORGANISMOS_PATH puede apuntar al .xlsx o a la carpeta que lo contiene.
function fromEnvVar() {
  const p = expandHome(process.env.CLAVES_ORGANISMOS_PATH);
  if (!p) return null;
  try {
    if (fs.statSync(p).isDirectory()) return path.join(p, FILE_NAME);
  } catch { /* no existe: lo devolvemos igual para poder reportar la ruta */ }
  return path.resolve(p);
}

// Rutas donde se busca el archivo, en orden.
function candidatePaths() {
  const out = [];
  const fromVar = fromEnvVar();
  if (fromVar) out.push(fromVar);
  out.push(path.join(PLUGIN_ROOT, FILE_NAME));
  out.push(path.join(os.homedir(), '.fisco-ar', FILE_NAME));
  return [...new Set(out)];
}

// Ruta del archivo de claves, o null si no hay ninguno.
function filePath() {
  for (const p of candidatePaths()) {
    if (p && fs.existsSync(p)) return p;
  }
  return null;
}

// Estado del archivo, para el diagnóstico (`node scripts/doctor.js`).
// Nunca devuelve claves: sólo si el archivo se pudo abrir y cuántas filas trae.
function status() {
  // Tarde a propósito: env.js requiere este módulo, así que se pide recién
  // cuando ya están los dos cargados.
  try { require('./env').loadEnv(); } catch { /* sin .env se mira el entorno */ }
  const configurado = fromEnvVar();
  const file = filePath();
  const out = {
    variable: process.env.CLAVES_ORGANISMOS_PATH || null,
    file,
    buscadoEn: candidatePaths(),
    passwordSet: Boolean(process.env.CLAVES_ORGANISMOS_PASSWORD),
  };

  if (configurado && !fs.existsSync(configurado)) out.rutaConfiguradaNoExiste = configurado;
  if (!file) return { ...out, ok: false, error: 'no encontré el archivo de claves' };

  try {
    const { rows, hoja } = loadSync();
    const conArca = rows.filter(r => r.organismos.ARCA && r.organismos.ARCA.clave).length;
    return { ...out, ok: true, hoja, titulares: rows.length, conClaveArca: conArca };
  } catch (e) {
    return { ...out, ok: false, error: e.message };
  }
}

// ── Normalizadores ────────────────────────────────────────────────────────────

const stripAccents = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
const norm      = (s) => stripAccents(s).toLowerCase().replace(/\s+/g, ' ').trim();
const normCuit  = (s) => String(s || '').replace(/\D/g, '');
const normOrg   = (s) => stripAccents(s).toUpperCase().replace(/[^A-Z0-9]/g, '');

// ARCA se llamó AFIP hasta 2024 y hay planillas viejas dando vueltas.
const ORG_ALIAS = { AFIP: 'ARCA', DGR: 'ARBA', RENTASCABA: 'AGIP' };
const organismo = (s) => ORG_ALIAS[normOrg(s)] || normOrg(s);

// ── Lectura de la planilla ────────────────────────────────────────────────────

// Descifra el .xlsx si hace falta. `officecrypto-tool` es una dependencia
// opcional: sin ella se pueden leer planillas sin contraseña igual.
async function readFileBuffer(file, password) {
  const raw = fs.readFileSync(file);

  let oct = null;
  try { oct = require('officecrypto-tool'); } catch { /* no instalado */ }

  const encrypted = oct ? oct.isEncrypted(raw) : raw.readUInt32LE(0) === 0xe011cfd0;
  if (!encrypted) return raw;

  if (!oct) {
    throw new Error(
      `${path.basename(file)} está protegido con contraseña y falta la dependencia que lo abre. ` +
      'Corré `npm install` en la raíz del plugin.'
    );
  }
  if (!password) {
    throw new Error(
      `${path.basename(file)} está protegido con contraseña. ` +
      'Agregá CLAVES_ORGANISMOS_PASSWORD al .env (o exportala en la shell).'
    );
  }
  try {
    return await oct.decrypt(raw, { password });
  } catch (e) {
    throw new Error(
      `no pude abrir ${path.basename(file)}: contraseña incorrecta o archivo dañado (${e.message}).`
    );
  }
}

// Ubica la fila de encabezados (la que tiene la columna CUIT) y arma, para cada
// columna, a qué organismo y a qué campo corresponde.
function mapColumns(rows) {
  const headerAt = rows.findIndex(r => r.some(c => norm(c) === 'cuit'));
  if (headerAt < 0) {
    throw new Error('no encontré la fila de encabezados (ninguna columna se llama "CUIT").');
  }
  const header = rows[headerAt];
  const groups = rows[headerAt - 1] || [];

  // La fila de organismos viene con celdas combinadas: sólo la primera columna
  // del grupo trae el nombre, así que se arrastra hacia la derecha.
  const columns = [];
  let current = '';
  for (let i = 0; i < header.length; i++) {
    if (norm(groups[i])) current = groups[i];
    const field = norm(header[i]);
    if (!field) { columns.push(null); continue; }
    columns.push({ org: organismo(current), field });
  }
  return { headerAt, columns };
}

// Mapea el encabezado de una columna a una propiedad del titular, o null.
function titularField(field) {
  if (field === '#' || field.startsWith('nro') || field.startsWith('n°')) return 'pos';
  if (field.includes('titular') || field.includes('razon') || field.includes('nombre')) return 'titular';
  if (field.includes('cuit')) return 'cuit';
  if (field.includes('tipo')) return 'tipo';
  if (field.includes('cierre')) return 'cierre';
  return null;
}

// Mapea el encabezado de una columna de organismo a usuario / clave / municipio.
function organismoField(field) {
  if (field.includes('usuario')) return 'usuario';
  if (field.includes('clave') || field.includes('contrasena') || field.includes('password')) return 'clave';
  if (field.includes('municipio') || field.includes('jurisdiccion')) return 'municipio';
  return null;
}

// Convierte la matriz cruda en filas con estructura.
function parseRows(rows) {
  const { headerAt, columns } = mapColumns(rows);
  const out = [];

  for (const raw of rows.slice(headerAt + 1)) {
    const row = { pos: '', titular: '', cuit: '', tipo: '', cierre: '', organismos: {} };
    let hasData = false;

    for (let i = 0; i < columns.length; i++) {
      const col = columns[i];
      const value = String(raw[i] === undefined || raw[i] === null ? '' : raw[i]).trim();
      if (!col || !value) continue;

      // El primer grupo (DATOS DEL TITULAR) no es un organismo: son los datos
      // de la fila. Se reconoce porque sus encabezados son #/Titular/CUIT/…
      const asTitular = titularField(col.field);
      if (asTitular && !['ARCA', 'ARBA', 'AGIP'].includes(col.org)) {
        row[asTitular] = asTitular === 'cuit' ? normCuit(value) : value;
        hasData = true;
        continue;
      }

      const asOrg = organismoField(col.field);
      if (!asOrg || !col.org) continue;
      (row.organismos[col.org] ||= {})[asOrg] = asOrg === 'usuario' && /^\D*\d[\d\s.-]{9,}$/.test(value)
        ? normCuit(value)   // los usuarios que son CUIT se guardan sin guiones
        : value;
      hasData = true;
    }

    // Una fila sirve si identifica a alguien: sin CUIT ni titular es relleno.
    if (hasData && (row.cuit || row.titular)) out.push(row);
  }
  return out;
}

// ── Carga (async e, para los scripts que resuelven credenciales de forma
//    sincrónica al arrancar, una versión sync que delega en un subproceso) ─────

let cache = null;

async function load() {
  if (cache) return cache;
  const file = filePath();
  if (!file) {
    throw new Error(
      `no encontré ${FILE_NAME}. Buscado en: ${candidatePaths().filter(Boolean).join(', ')}. ` +
      'Podés apuntar a otra ruta con CLAVES_ORGANISMOS_PATH.'
    );
  }
  const buf  = await readFileBuffer(file, process.env.CLAVES_ORGANISMOS_PASSWORD);
  const wb   = readWorkbook(buf);
  const rows = parseRows(wb.rows(process.env.CLAVES_ORGANISMOS_SHEET || undefined));
  cache = { file, hoja: wb.sheetNames[0], rows };
  return cache;
}

// Igual que load(), pero sincrónico: el descifrado es asincrónico, así que se
// hace en un subproceso corto y se recibe el JSON ya parseado. Lo usan los
// scripts que resuelven las credenciales antes de entrar a su main() async.
function loadSync() {
  if (cache) return cache;
  const helper = path.join(__dirname, 'claves-dump.js');
  let stdout;
  try {
    stdout = execFileSync(process.execPath, [helper], {
      encoding: 'utf-8',
      maxBuffer: 32 * 1024 * 1024,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
  } catch (e) {
    // El helper imprime el error como JSON en stdout antes de salir con 1.
    const payload = safeParse(e.stdout);
    throw new Error(payload && payload.error ? payload.error : (e.stderr || e.message).trim());
  }
  const payload = safeParse(stdout);
  if (!payload || payload.error) throw new Error((payload && payload.error) || 'no pude leer el archivo de claves');
  cache = payload;
  return cache;
}

const safeParse = (s) => { try { return JSON.parse(s); } catch { return null; } };

// ── Búsqueda ──────────────────────────────────────────────────────────────────

// Busca por nombre de titular, CUIT o número de fila (#).
// Devuelve { row } si hay una sola coincidencia, o { candidatos } si hay varias.
// Tira error si no hay ninguna.
function buscar(rows, { nombre, cuit, pos } = {}) {
  const criterio = [];

  if (cuit) {
    const target = normCuit(cuit);
    criterio.push(`CUIT ${target}`);
    const hit = rows.filter(r => r.cuit === target);
    if (hit.length) return one(hit);
  }

  if (pos !== undefined && pos !== null && String(pos).trim() !== '') {
    const target = String(pos).trim();
    criterio.push(`posición ${target}`);
    const hit = rows.filter(r => String(r.pos).trim() === target);
    if (hit.length) return one(hit);
  }

  if (nombre && String(nombre).trim()) {
    const q = norm(nombre);
    criterio.push(`nombre "${nombre}"`);
    // Exacto primero; recién si no hay, coincidencia parcial.
    const exact = rows.filter(r => norm(r.titular) === q);
    if (exact.length) return one(exact);
    const partial = rows.filter(r => norm(r.titular).includes(q));
    if (partial.length) return one(partial);
  }

  if (!criterio.length) throw new Error('hay que buscar por --nombre, --cuit o --pos.');
  throw new Error(`no encontré ningún titular por ${criterio.join(' ni ')} en el archivo de claves.`);
}

const one = (hits) => (hits.length === 1 ? { row: hits[0] } : { candidatos: hits });

// Credenciales de un organismo para una fila ya encontrada.
// Para ARCA, el usuario es el CUIT de login (puede ser el del apoderado).
function credenciales(row, org = 'ARCA') {
  const key = organismo(org);
  const data = row.organismos[key] || {};
  return {
    cuit:      row.cuit,
    titular:   row.titular,
    organismo: key,
    usuario:   data.usuario || '',
    clave:     data.clave || '',
    ...(data.municipio ? { municipio: data.municipio } : {}),
  };
}

// Oculta una clave dejando sólo el largo — para imprimir sin filtrar secretos.
const mask = (s) => (s ? `${'•'.repeat(Math.min(String(s).length, 12))} (${String(s).length} caracteres)` : '');

module.exports = {
  FILE_NAME, candidatePaths, filePath, status, load, loadSync, buscar, credenciales,
  parseRows, mask, normCuit, norm, organismo,
};
