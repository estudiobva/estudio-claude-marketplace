#!/usr/bin/env node
/**
 * configurar.js — arma (o completa) el .env del estudio, ~/.fisco-ar/.env, que
 * leen todos los plugins, y deja el marketplace con actualizacion automatica.
 *
 *   node scripts/configurar.js                 detecta y completa lo que falte
 *   node scripts/configurar.js --revisar       no escribe nada: informa
 *   node scripts/configurar.js --unidad="G:\Unidades compartidas\BVA - Sociedades Farmaceuticas"
 *   node scripts/configurar.js --python="C:\...\python.exe"
 *   node scripts/configurar.js --json
 *
 * Detecta:
 *   BVA_UNIDAD_PATH         la unidad compartida (I:\, G:\ u otra letra; en Mac,
 *                           ~/Library/CloudStorage/GoogleDrive-*)
 *   CLAVES_ORGANISMOS_PATH  <unidad>\0 - ESTUDIO - 00000000000\Claves_Organismos.xlsx
 *   BVA_PYTHON              un Python 3 real (no el atajo de la Microsoft Store)
 *
 * Las contrasenas NUNCA se pasan por la linea de comandos ni se imprimen: el
 * instalador las pide en su ventana y las pasa en las variables de entorno
 * BVA_SETUP_CLAVE (Claves_Organismos.xlsx) y BVA_SETUP_CLAVE_BANCOS
 * (Claves_Bancos.xlsx, solo si es distinta). Se prueban abriendo la planilla y
 * solo se guardan si abre.
 *
 * Antes de modificar el .env deja una copia .env.bak-AAAAMMDD-HHMMSS al lado.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const p = require('../lib/plugins');

const args = Object.fromEntries(process.argv.slice(2).map(a => {
  const m = a.match(/^--([^=]+)(?:=(.*))?$/);
  return m ? [m[1], m[2] === undefined ? true : m[2]] : [a, true];
}));
const REVISAR = Boolean(args.revisar);
const JSON_OUT = Boolean(args.json);

const CARPETA_ESTUDIO = '0 - ESTUDIO - 00000000000';
const NOMBRE_UNIDAD = 'BVA - Sociedades Farmaceuticas';

// ── Unidad compartida ────────────────────────────────────────────────────────

function candidatasUnidad() {
  const out = [];
  if (p.WIN) {
    // Google Drive para escritorio monta una letra (G: por defecto; en las PCs
    // del estudio, I:). La carpeta se llama "Unidades compartidas" o "Shared drives".
    const letras = ['I', 'G', ...'DEFHJKLMNOPQRSTUVWXYZ'.split('').filter(l => !'IG'.includes(l))];
    for (const l of letras) {
      for (const sub of ['Unidades compartidas', 'Shared drives']) out.push(`${l}:\\${sub}\\${NOMBRE_UNIDAD}`);
    }
  } else {
    const cs = path.join(os.homedir(), 'Library', 'CloudStorage');
    let cuentas = [];
    try { cuentas = fs.readdirSync(cs).filter(d => /^GoogleDrive-/i.test(d)); } catch { /* sin Drive */ }
    cuentas.sort((a, b) => /estudiobva/i.test(b) - /estudiobva/i.test(a));
    for (const c of cuentas) {
      for (const sub of ['Unidades compartidas', 'Shared drives']) out.push(path.join(cs, c, sub, NOMBRE_UNIDAD));
    }
  }
  return out;
}

const esUnidad = (d) => { try { return fs.statSync(path.join(d, CARPETA_ESTUDIO)).isDirectory(); } catch { return false; } };

function detectarUnidad(env) {
  if (args.unidad && args.unidad !== true) return esUnidad(args.unidad) ? args.unidad : null;
  const actual = p.variable('BVA_UNIDAD_PATH', env);
  if (actual && esUnidad(actual)) return actual;
  return candidatasUnidad().find(esUnidad) || null;
}

// ── Planillas de claves ──────────────────────────────────────────────────────

// Motivo del error, para que el instalador sepa si tiene que pedir la contrasena.
function motivo(msg) {
  if (/contrase/i.test(msg)) return 'clave';
  if (/no encontr|no encuentro|no existe|ENOENT/i.test(msg)) return 'archivo';
  if (/dependencia|officecrypto|npm install|Cannot find module/i.test(msg)) return 'dependencia';
  return 'otro';
}

// Un .xlsx cifrado con contrasena es un contenedor OLE (D0 CF 11 E0); uno comun es un zip (PK).
function cifrado(archivo) {
  try {
    const fd = fs.openSync(archivo, 'r'); const b = Buffer.alloc(4);
    fs.readSync(fd, b, 0, 4, 0); fs.closeSync(fd);
    return b.readUInt32LE(0) === 0xe011cfd0;
  } catch { return null; }
}

// Abre Claves_Organismos.xlsx con el lib/claves.js de un plugin instalado (el
// mismo codigo que usan los scripts). Devuelve cuantos titulares trae, nunca claves.
function probarClavesOrganismos(archivo, clave) {
  const pl = p.instalados().find(x => fs.existsSync(path.join(x.dir, 'lib', 'claves.js')) && !p.nodeFaltantes(x.dir).length);
  if (!pl) return { ok: false, motivo: 'dependencia', error: 'ningun plugin con lib/claves.js tiene las dependencias instaladas (corre dependencias.js)' };
  if (!fs.existsSync(archivo)) return { ok: false, motivo: 'archivo', error: `no existe ${archivo}` };
  process.env.CLAVES_ORGANISMOS_PATH = archivo;
  if (clave) process.env.CLAVES_ORGANISMOS_PASSWORD = clave; else delete process.env.CLAVES_ORGANISMOS_PASSWORD;
  try {
    const s = require(path.join(pl.dir, 'lib', 'claves.js')).status();
    if (s.ok) return { ok: true, titulares: s.titulares, conClaveArca: s.conClaveArca };
    return { ok: false, motivo: motivo(s.error || ''), error: s.error };
  } catch (e) {
    return { ok: false, motivo: motivo(e.message), error: e.message.split('\n')[0] };
  }
}

async function probarClavesBancos(unidad, clave, claveOrganismos) {
  const pl = p.instalados().find(x => x.nombre === 'bancos-bva');
  if (!pl) return null; // no esta instalado: no hace falta
  if (p.nodeFaltantes(pl.dir).length) return { ok: false, motivo: 'dependencia', error: 'bancos-bva sin dependencias (corre dependencias.js)' };
  process.env.BVA_UNIDAD_PATH = unidad;
  if (clave) process.env.CLAVES_BANCOS_PASSWORD = clave; else delete process.env.CLAVES_BANCOS_PASSWORD;
  if (claveOrganismos) process.env.CLAVES_ORGANISMOS_PASSWORD = claveOrganismos;
  try {
    const cb = require(path.join(pl.dir, 'lib', 'claves-bancos.js'));
    const { bancos } = await cb.load();
    return { ok: true, bancos: bancos.length };
  } catch (e) {
    return { ok: false, motivo: motivo(e.message), error: e.message.split('\n')[0] };
  }
}

// ── .env ─────────────────────────────────────────────────────────────────────

const comillas = (v) => `"${v}"`;

// Reemplaza la linea KEY= si existe (no las comentadas); si no, la agrega al final.
function fijar(texto, clave, valor) {
  const re = new RegExp(`^[ \\t]*(export[ \\t]+)?${clave}[ \\t]*=.*$`, 'm');
  const linea = `${clave}=${comillas(valor)}`;
  if (re.test(texto)) return texto.replace(re, () => linea); // funcion: un "$" en la clave no es patron
  return texto.replace(/\s*$/, '') + (texto.trim() ? '\n' : '') + linea + '\n';
}

const sello = () => new Date().toISOString().replace(/[-:]/g, '').replace('T', '-').slice(0, 15);

// ── Actualizacion automatica del marketplace ─────────────────────────────────
// extraKnownMarketplaces.<nombre>.autoUpdate en ~/.claude/settings.json: Claude
// refresca el marketplace y actualiza los plugins al arrancar.

function autoUpdate() {
  const file = path.join(p.claudeDir(), 'settings.json');
  let s = {};
  if (fs.existsSync(file)) {
    s = p.leerJson(file);
    if (!s) return { ok: false, error: `${file} no es JSON valido: no lo toco` };
  }
  const actual = s.extraKnownMarketplaces && s.extraKnownMarketplaces[p.MARKET];
  if (actual && actual.autoUpdate === true) return { ok: true, cambio: false };
  if (REVISAR) return { ok: false, cambio: false, error: 'autoUpdate sin activar' };
  s.extraKnownMarketplaces = s.extraKnownMarketplaces || {};
  s.extraKnownMarketplaces[p.MARKET] = {
    ...(actual || {}),
    source: (actual && actual.source) || { source: 'github', repo: p.REPO },
    autoUpdate: true,
  };
  if (fs.existsSync(file)) fs.copyFileSync(file, `${file}.bak-${sello()}`);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(s, null, 2) + '\n');
  return { ok: true, cambio: true };
}

// ── Principal ────────────────────────────────────────────────────────────────

(async () => {
  const file = p.envFile();
  const existe = fs.existsSync(file);
  const texto0 = existe ? fs.readFileSync(file, 'utf-8') : '';
  const env = p.parseEnv(texto0);
  const nuevos = {};
  const res = { ok: true, envFile: file, envExistia: existe, cambios: [], avisos: [] };

  // Unidad
  const unidad = detectarUnidad(env);
  res.unidad = unidad;
  if (!unidad) {
    res.ok = false;
    res.avisos.push('No encuentro la unidad compartida "BVA - Sociedades Farmaceuticas". ' +
      'Abri Google Drive para escritorio con la cuenta @estudiobva.com y espera a que aparezca, o pasa --unidad=<ruta>.');
  } else if (env.BVA_UNIDAD_PATH !== unidad) {
    nuevos.BVA_UNIDAD_PATH = unidad;
  }

  // Planilla de claves de organismos
  let archivoClaves = p.variable('CLAVES_ORGANISMOS_PATH', env);
  if ((!archivoClaves || !fs.existsSync(archivoClaves)) && unidad) {
    archivoClaves = path.join(unidad, CARPETA_ESTUDIO, 'Claves_Organismos.xlsx');
    nuevos.CLAVES_ORGANISMOS_PATH = archivoClaves;
  }
  res.clavesArchivo = archivoClaves || null;

  // Python
  let py = args.python && args.python !== true ? args.python : null;
  const pyActual = p.variable('BVA_PYTHON', env);
  if (!py) {
    const chequeo = pyActual ? p.correr(pyActual, ['-c', 'import sys;print(sys.executable)'], { timeout: 20000 }) : null;
    if (chequeo && chequeo.ok && !/WindowsApps/i.test(chequeo.out)) py = pyActual;
    else { const d = p.detectarPython(); py = d && d.exe; }
  }
  res.python = py || null;
  if (!py) { res.ok = false; res.avisos.push('No encuentro Python 3.9 o superior.'); }
  else if (env.BVA_PYTHON !== py) nuevos.BVA_PYTHON = py;

  // Contrasena de Claves_Organismos: la del instalador (si vino) o la del .env.
  const claveNueva = process.env.BVA_SETUP_CLAVE || '';
  const claveEnv = p.variable('CLAVES_ORGANISMOS_PASSWORD', env);
  if (archivoClaves) {
    const r = probarClavesOrganismos(archivoClaves, claveNueva || claveEnv);
    r.cifrado = cifrado(archivoClaves);
    // Sin cifrar abre con cualquier contrasena: no se guarda ninguna.
    if (r.ok && r.cifrado && claveNueva && claveNueva !== env.CLAVES_ORGANISMOS_PASSWORD) nuevos.CLAVES_ORGANISMOS_PASSWORD = claveNueva;
    if (r.ok && r.cifrado === false) res.avisos.push('Claves_Organismos.xlsx no esta protegido con contrasena: cualquiera con acceso a la unidad puede leer las claves.');
    if (!r.ok && r.motivo === 'clave' && !claveNueva && !claveEnv) r.error = 'falta la contrasena';
    res.claves = r;
  } else {
    res.claves = { ok: false, motivo: 'archivo', error: 'no se sabe donde esta Claves_Organismos.xlsx (falta la unidad)' };
  }
  if (!res.claves.ok) res.ok = false;

  // Claves_Bancos: por defecto la misma contrasena. Solo se guarda una propia si hace falta.
  if (unidad) {
    const claveBancos = process.env.BVA_SETUP_CLAVE_BANCOS || p.variable('CLAVES_BANCOS_PASSWORD', env);
    // Como claves-bancos.js: sin CLAVES_BANCOS_PASSWORD usa la de organismos.
    const claveOrg = res.claves.ok ? (claveNueva || claveEnv) : '';
    const rb = await probarClavesBancos(unidad, claveBancos, claveOrg);
    if (rb) {
      if (rb.ok && process.env.BVA_SETUP_CLAVE_BANCOS && process.env.BVA_SETUP_CLAVE_BANCOS !== env.CLAVES_BANCOS_PASSWORD) {
        nuevos.CLAVES_BANCOS_PASSWORD = process.env.BVA_SETUP_CLAVE_BANCOS;
      }
      res.clavesBancos = rb;
      if (!rb.ok) res.avisos.push('Claves_Bancos.xlsx no abre: el plugin bancos-bva no va a poder entrar a los bancos.');
    }
  }

  // Carpeta de salida por defecto de los scripts
  res.salidas = p.variable('BVA_SALIDAS_PATH', env) || path.join(os.homedir(), 'Documents', 'BVA-salidas');

  // Escribir
  res.cambios = Object.keys(nuevos);
  if (!REVISAR && res.cambios.length) {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    if (existe) { res.backup = `${file}.bak-${sello()}`; fs.copyFileSync(file, res.backup); }
    let texto = texto0 || '# Estudio BVA — configuracion de los plugins (lo arma setup-bva). NUNCA compartir este archivo.\n';
    if (!/setup-bva/.test(texto)) texto = texto.replace(/\s*$/, '') + '\n\n# setup-bva\n';
    for (const [k, v] of Object.entries(nuevos)) texto = fijar(texto, k, v);
    fs.writeFileSync(file, texto, { mode: 0o600 });
  }
  if (!REVISAR) fs.mkdirSync(res.salidas, { recursive: true });

  res.autoUpdate = autoUpdate();
  if (!res.autoUpdate.ok && !REVISAR) res.avisos.push(`No pude activar la actualizacion automatica: ${res.autoUpdate.error}`);

  if (JSON_OUT) {
    console.log(JSON.stringify(res, null, 2));
  } else {
    const linea = (ok, t, leve) => console.log(`  ${ok ? 'OK   ' : (leve ? 'AVISO' : 'ERROR')} ${t}`);
    console.log(`.env: ${file}${res.backup ? `  (copia: ${path.basename(res.backup)})` : ''}`);
    linea(Boolean(unidad), `Unidad compartida: ${unidad || 'no encontrada'}`);
    linea(Boolean(py), `Python: ${py || 'no encontrado'}`);
    linea(res.claves.ok, `Claves_Organismos.xlsx: ${res.claves.ok ? `abre, ${res.claves.titulares} titulares` : res.claves.error}`);
    if (res.clavesBancos) linea(res.clavesBancos.ok, `Claves_Bancos.xlsx: ${res.clavesBancos.ok ? 'abre' : res.clavesBancos.error}`, true);
    linea(res.autoUpdate.ok, `Actualizacion automatica del marketplace: ${res.autoUpdate.ok ? 'activada' : res.autoUpdate.error}`, true);
    if (res.cambios.length) console.log(`  ${REVISAR ? 'Cambiaria' : 'Guardado'}: ${res.cambios.join(', ')}`);
    for (const a of res.avisos) console.log(`  AVISO ${a}`);
  }
  process.exit(res.ok ? 0 : 1);
})();
