// ─────────────────────────────────────────────────────────────────────────────
// Vendorizado desde el plugin fisco-ar (MIT) — https://github.com/javiergradiche/fisco-ar-claude-plugin
// Commit de origen: 2ad06df · Copiado el 21/09/2026 al repo del Estudio BVA.
//
// Se copia en vez de depender del plugin para que una actualizacion del
// marketplace no pueda romper un cierre. Si ARCA cambia el login, el arreglo va
// ACA (y conviene mirar upstream por si ya lo resolvieron).
// ─────────────────────────────────────────────────────────────────────────────

// env.js — carga del .env y resolución de credenciales por CUIT.
//
// El .env NUNCA se commitea (está en .gitignore). Se busca, en orden:
//   1. $FISCO_ENV_FILE               (override explícito)
//   2. <raíz del plugin>/.env        (checkout del repo / modo dev)
//   3. ~/.fisco-ar/.env              (plugin instalado desde el marketplace)
//
// Formato de credenciales — un bloque por CUIT consultado:
//   ARCA_20123456789_LOGIN=20123456789     ← CUIT con el que se loguea
//   ARCA_20123456789_PASSWORD=miClave      ← clave fiscal de ESE CUIT login
//
// Cuando hay representación, el LOGIN es el CUIT del representante:
//   ARCA_30712345678_LOGIN=20123456789
//   ARCA_30712345678_PASSWORD=miClave
//
// Las variables que ya existen en el entorno (exportadas en la shell) tienen
// prioridad sobre el archivo: el .env sólo rellena lo que falta.
//
// Además del .env, las credenciales pueden salir del Excel de claves del
// estudio (Claves_Organismos.xlsx — ver lib/claves.js), que permite buscar por
// nombre del titular en vez de por CUIT. El .env siempre tiene prioridad.

const fs   = require('fs');
const path = require('path');
const os   = require('os');
const claves = require('./claves');

// Ajuste al vendorizar: aca lib/ cuelga directo de la raiz del proyecto,
// no de scripts/lib/ como en el plugin. Sin esto el .env propio no se encuentra.
const PLUGIN_ROOT = path.join(__dirname, '..');

function candidatePaths() {
  const out = [];
  if (process.env.FISCO_ENV_FILE) out.push(process.env.FISCO_ENV_FILE);
  out.push(path.join(PLUGIN_ROOT, '.env'));
  out.push(path.join(os.homedir(), '.fisco-ar', '.env'));
  return out;
}

// Parser mínimo de .env (sin dependencias): KEY=VALUE, # comentarios,
// comillas opcionales, `export ` opcional al principio de la línea.
function parseEnv(text) {
  const out = {};
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith('#')) continue;
    const m = line.replace(/^export\s+/, '').match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
    if (!m) continue;
    let value = m[2].trim();
    if ((value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    } else {
      value = value.replace(/\s+#.*$/, '').trim(); // comentario al final de línea
    }
    out[m[1]] = value;
  }
  return out;
}

let loadedFrom = null;

// Carga el primer .env que exista y lo mergea en process.env (sin pisar nada).
// Devuelve el path usado, o null si no encontró ninguno.
function loadEnv() {
  if (loadedFrom !== null) return loadedFrom || null;
  loadedFrom = '';
  for (const p of candidatePaths()) {
    if (!p || !fs.existsSync(p)) continue;
    const vars = parseEnv(fs.readFileSync(p, 'utf-8'));
    for (const [k, v] of Object.entries(vars)) {
      if (process.env[k] === undefined) process.env[k] = v;
    }
    loadedFrom = p;
    break;
  }
  return loadedFrom || null;
}

const normCuit = (c) => String(c || '').replace(/\D/g, '');

// CUITs con credenciales cargadas (los que tienen ARCA_<cuit>_PASSWORD).
function listCuits() {
  loadEnv();
  return Object.entries(process.env)
    .filter(([, v]) => v)  // una clave vacía no cuenta como configurada
    .map(([k]) => (k.match(/^ARCA_(\d{11})_PASSWORD$/) || [])[1])
    .filter(Boolean)
    .sort();
}

// Resuelve { cuit, cuitLogin, password } combinando CLI, .env y el archivo de
// claves (Claves_Organismos.xlsx — ver lib/claves.js).
//
// Precedencia: argumentos de CLI > .env del CUIT pedido > archivo de claves.
// A quién buscar sale de --cuit, o de --nombre (el titular tal como figura en
// el archivo de claves). Si no se pasa ninguno y hay exactamente un CUIT en el
// .env, se usa ese.
function resolveCredentials({ cuit, cuitLogin, password, nombre } = {}) {
  loadEnv();
  let target = normCuit(cuit) || normCuit(cuitLogin);
  let desdeArchivo = null;

  // Búsqueda por nombre: el archivo de claves dice a qué CUIT corresponde.
  if (!target && nombre && String(nombre).trim()) {
    desdeArchivo = buscarEnArchivo({ nombre });
    target = desdeArchivo.cuit;
  }

  const configured = listCuits();
  if (!target && configured.length === 1) target = configured[0];

  if (!target) {
    const hint = configured.length
      ? `CUITs disponibles en el .env: ${configured.join(', ')}.`
      : (claves.filePath()
          ? 'Podés pasar --nombre=<titular> para buscarlo en el archivo de claves.'
          : 'No hay ningún CUIT configurado en el .env (ver .env.example).');
    throw new Error(`Falta --cuit o --nombre. ${hint}`);
  }

  const envLogin = process.env[`ARCA_${target}_LOGIN`];
  const envPass  = process.env[`ARCA_${target}_PASSWORD`];

  // El archivo de claves cubre lo que el .env no tenga.
  if (!password && !envPass && !desdeArchivo && claves.filePath()) {
    try { desdeArchivo = buscarEnArchivo({ cuit: target }); } catch { /* se reporta abajo */ }
  }

  const resolvedLogin = normCuit(cuitLogin) || normCuit(envLogin) ||
                        normCuit(desdeArchivo && desdeArchivo.usuario) || target;
  const resolvedPass  = password || envPass || (desdeArchivo && desdeArchivo.clave) || '';

  if (!resolvedPass) {
    const dondeBuscar = claves.filePath()
      ? `Agregá ARCA_${target}_LOGIN y ARCA_${target}_PASSWORD al .env ` +
        `(${loadedFrom || 'ningún .env encontrado'}), completá la fila del CUIT ${target} ` +
        `en ${claves.filePath()}, o pasá --password.`
      : `Agregá ARCA_${target}_LOGIN y ARCA_${target}_PASSWORD al .env ` +
        `(${loadedFrom || 'ningún .env encontrado'}), o pasá --password.`;
    throw new Error(`No encontré la clave fiscal del CUIT ${target}. ${dondeBuscar}`);
  }

  return {
    cuit: target,
    cuitLogin: resolvedLogin,
    password: resolvedPass,
    ...(desdeArchivo && desdeArchivo.titular ? { titular: desdeArchivo.titular } : {}),
    origen: password ? 'cli' : (envPass ? 'env' : 'archivo'),
  };
}

// Busca un titular en el archivo de claves y devuelve sus credenciales de ARCA.
// Traduce las ambigüedades a un error accionable (con la lista de candidatos).
function buscarEnArchivo(criterio) {
  const { rows } = claves.loadSync();
  const { row, candidatos } = claves.buscar(rows, criterio);
  if (candidatos) {
    const MOSTRAR = 10;
    const lista = candidatos.slice(0, MOSTRAR)
      .map(r => `  ${r.pos}. ${r.titular} (CUIT ${r.cuit})`).join('\n');
    const resto = candidatos.length > MOSTRAR ? `\n  …y ${candidatos.length - MOSTRAR} más.` : '';
    throw new Error(
      `"${criterio.nombre}" coincide con ${candidatos.length} titulares del archivo de claves. ` +
      `Afiná el nombre o pasá --cuit:\n${lista}${resto}`
    );
  }
  return claves.credenciales(row, 'ARCA');
}

// CUITs que están en el archivo de claves (con clave de ARCA cargada).
function listCuitsArchivo() {
  if (!claves.filePath()) return [];
  try {
    return claves.loadSync().rows
      .filter(r => r.cuit && r.organismos.ARCA && r.organismos.ARCA.clave)
      .map(r => r.cuit);
  } catch { return []; }
}

// API key de CapSolver (para los captchas de AGIP). CLI > .env.
function capsolverKey(fromCli) {
  loadEnv();
  return fromCli || process.env.CAPSOLVER_API_KEY || '';
}

// Path del .env efectivamente cargado (para diagnósticos).
function envFile() {
  loadEnv();
  return loadedFrom || null;
}

module.exports = {
  loadEnv, listCuits, listCuitsArchivo, resolveCredentials, capsolverKey,
  envFile, candidatePaths, PLUGIN_ROOT,
};
