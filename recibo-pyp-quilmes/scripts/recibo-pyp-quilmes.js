#!/usr/bin/env node
/**
 * recibo-pyp-quilmes.js — baja de ARQui (Municipio de Quilmes) el recibo actual
 * de Derechos de Publicidad y Propaganda (marquesinas) de una, varias o todas
 * las sociedades, controla que sea de la sociedad y lo archiva en la carpeta
 * TSH del mes, donde el equipo guarda las boletas de marquesina.
 *
 *   node scripts/recibo-pyp-quilmes.js --nombre="<sociedad>"
 *   node scripts/recibo-pyp-quilmes.js --cuit=<CUIT>
 *   node scripts/recibo-pyp-quilmes.js --nombres="<A>;<B>"
 *   node scripts/recibo-pyp-quilmes.js --todas
 *   node scripts/recibo-pyp-quilmes.js --listar
 *
 * Argumentos:
 *   --nombre / --cuit   Una sociedad (titular en Claves_Organismos.xlsx, o CUIT).
 *   --nombres / --cuits Varias, separadas por ";".
 *   --todas             Todas las cuentas de PyP vinculadas a los usuarios de
 *                       ARQui de la planilla; la sociedad sale del CUIT del PDF.
 *                       Agrega la hoja "Control planilla" (ARQui vs Tasas Municipales CLAVES).
 *   --cuenta            Numero de cuenta de PyP (sin digito verificador), para
 *                       una sola sociedad, si no esta en la planilla de tasas.
 *   --listar            Solo lista las cuentas del desplegable con lo que dice
 *                       la planilla de tasas. No baja nada.
 *   --no-archivar       No copia a la unidad: deja los PDF solo en la carpeta de salida.
 *   --reemplazar        Pisa el PDF del mes si ya esta en la carpeta TSH (default: lo deja).
 *   --ver               Navegador visible (solo si se pide).
 *
 * Cuentas: bloque PUBLICIDAD Y PROPAGANDA de la hoja QUILMES de "Tasas
 * Municipales CLAVES.xlsx" (carpeta del estudio en la unidad), por CUIT. Si la
 * sociedad no esta ahi, se buscan sus recibos entre las cuentas del
 * desplegable que la planilla no asigna a nadie (controlando el CUIT del PDF).
 *
 * Login: usuario y clave del grupo MUNICIPAL (municipio Quilmes) de
 * Claves_Organismos.xlsx, o ARQUI_USUARIO / ARQUI_CLAVE en el .env. Un login
 * por usuario para todo el lote.
 *
 * Control antes de archivar: el CUIT del recibo tiene que ser el de la
 * sociedad y la cuenta del recibo la elegida (mas su digito verificador).
 *
 * Archiva [Sociedad]/01-Impuestos Mensuales/AAAAMM/TSH/AAAAMM-boleta marquesina- <Sociedad>.pdf
 * (AAAAMM = mes del recibo). Copia de todo + resumen Excel en
 * ~/Documents/BVA-salidas/recibo-pyp-quilmes/AAAA-MM-DD/ ($BVA_SALIDAS_PATH).
 *
 * Salida: JSON por stdout; progreso por stderr. Solo lee y descarga: no paga,
 * no presenta DDJJ ni vincula cuentas en ARQui.
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const { credencialesArqui, accesosQuilmes } = require('../lib/arqui-claves');
const tasas = require('../lib/tasas-claves');
const arqui = require('../lib/arqui');
const rutas = require('../lib/rutas-bva');

const args = parseArgs();
const VER = 'ver' in args;
const ARCHIVAR = !('no-archivar' in args);
const REEMPLAZAR = 'reemplazar' in args;
const LISTAR = 'listar' in args;
const TODAS = 'todas' in args;
const CUENTA = args['cuenta'] ? tasas.normCuenta(args['cuenta']) : null;
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
// ARQui bloquea (403) si recibe muchos pedidos seguidos: pausa entre recibos y,
// ante un bloqueo, una espera antes del unico reintento.
const PAUSA_MS = Number(process.env.ARQUI_PAUSA_MS || 2500);
const ESPERA_BLOQUEO_MS = Number(process.env.ARQUI_ESPERA_BLOQUEO_MS || 180000);

if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';
loadEnv();

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length && !TODAS && !LISTAR) fail('Falta --nombre, --cuit, --nombres, --cuits, --todas o --listar');
if (CUENTA && (PEDIDAS.length !== 1 || TODAS)) fail('--cuenta es para una sola sociedad (con --nombre o --cuit)');
if (TODAS && PEDIDAS.length) fail('--todas no se combina con --nombre/--cuit');

const log = (m) => process.stderr.write(`[recibo-pyp-quilmes] ${m}\n`);

const hoy = new Date();
const FECHA = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}-${String(hoy.getDate()).padStart(2, '0')}`;
const SALIDA = path.join(process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'recibo-pyp-quilmes', FECHA);
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'recibo-pyp-'));

// --nombre se busca en el archivo de claves (mismos nombres que el resto de los
// scripts); --cuit va directo.
function cuitDe(pedida) {
  if (pedida.cuit) return claves.normCuit(pedida.cuit);
  const { rows } = claves.loadSync();
  const { row, candidatos } = claves.buscar(rows, { nombre: pedida.nombre });
  if (candidatos) {
    throw new Error(`"${pedida.nombre}" coincide con ${candidatos.length} titulares: ` +
      candidatos.slice(0, 8).map((r) => `${r.titular} (${r.cuit})`).join(', ') + '. Afina el nombre o pasa --cuit.');
  }
  if (!row.cuit) throw new Error(`"${row.titular}" no tiene CUIT en el archivo de claves.`);
  return row.cuit;
}

function nombreDe(cuit, fallback) {
  try { return rutas.nombreSociedad(cuit); } catch { return fallback || cuit; }
}

function leerPdf(pdf) {
  const out = execFileSync(PYTHON, [path.join(__dirname, 'recibo_pyp.py'), 'leer', pdf], { encoding: 'utf-8' });
  return JSON.parse(out.trim().split(/\r?\n/).pop());
}

const dormir = (ms) => new Promise((r) => setTimeout(r, ms));
const fechaDe = (dmy) => { const [d, m, a] = dmy.split('/').map(Number); return new Date(a, m - 1, d); };

// ── Sesion de ARQui (un usuario a la vez) ─────────────────────────────────────
//
// ARQui bloquea (403) si, con la sesion de un usuario abierta despues de muchas
// descargas, entra otro. Por eso hay una sola sesion abierta: al cambiar de
// usuario se cierra el contexto anterior (cookies limpias) y los lotes van
// primero con los usuarios de pocas cuentas y al final con el general.

let browser = null;
const sesiones = new Map(); // usuario -> { cred, page, ctx, logueado }

async function sesion(cred) {
  if (!browser) browser = await chromium.launch(launchOptions());
  let s = sesiones.get(cred.usuario);
  if (!s) {
    for (const [u, otra] of sesiones) {
      await otra.ctx.close().catch(() => {});
      sesiones.delete(u);
    }
    const ctx = await browser.newContext({ acceptDownloads: true });
    s = { cred, ctx, page: await ctx.newPage(), logueado: false };
    s.page.on('dialog', (d) => d.dismiss().catch(() => {}));
    await arqui.prepararPagina(s.page);
    sesiones.set(cred.usuario, s);
  }
  if (!s.logueado) {
    log(`login en ARQui (${cred.origen})...`);
    await arqui.login(s.page, cred);
    s.logueado = true;
  }
  return s;
}

// Corre fn(page). Si la sesion vencio, reloguea y reintenta una vez. Si ARQui
// bloqueo (403), espera y reintenta una vez; si sigue, corta el resto del lote
// (insistir alarga el bloqueo).
let bloqueado = null;
async function conSesion(cred, fn) {
  if (bloqueado) throw bloqueado;
  const intento = async () => fn((await sesion(cred)).page);
  try {
    return await intento();
  } catch (e) {
    if (e.code === 'sesion') {
      log('la sesion vencio: vuelvo a loguear');
      sesiones.get(cred.usuario).logueado = false;
      return intento();
    }
    if (e.code !== 'bloqueo') throw e;
    log(`ARQui bloqueo los pedidos (403): espero ${Math.round(ESPERA_BLOQUEO_MS / 1000)} s y reintento una vez...`);
    await dormir(ESPERA_BLOQUEO_MS);
    for (const s of sesiones.values()) s.logueado = false;
    try {
      return await intento();
    } catch (e2) {
      if (e2.code === 'bloqueo') {
        bloqueado = Object.assign(new Error(`${e2.message} No se siguio con el resto: volver a correr en un rato.`), { code: 'bloqueo' });
        throw bloqueado;
      }
      throw e2;
    }
  }
}

// Baja y lee el recibo de una cuenta (cacheado: --todas y la busqueda por CUIT
// pueden pedir la misma cuenta dos veces).
const recibos = new Map();
let descargas = 0;
async function recibo(cred, cuenta) {
  if (recibos.has(cuenta)) return recibos.get(cuenta);
  if (descargas++ > 0) await dormir(PAUSA_MS);
  const pdf = path.join(TMP, `${cuenta}.pdf`);
  await conSesion(cred, (page) => arqui.descargarRecibo(page, cuenta, pdf));
  const datos = leerPdf(pdf);
  if (datos.error) throw Object.assign(new Error(datos.error), { code: 'pdf' });
  datos.pdf = pdf;
  recibos.set(cuenta, datos);
  return datos;
}

// ── Archivo ───────────────────────────────────────────────────────────────────

// Control + copia a la salida y (si corresponde) a la carpeta TSH del mes.
function guardar(fila, datos, { cuitEsperado, varias }) {
  const cuentaPdf = String(datos.cuenta || '').replace(/-\d$/, '');
  if (cuentaPdf !== fila.cuenta) {
    throw Object.assign(new Error(`El recibo es de la cuenta ${datos.cuenta}, no de la ${fila.cuenta}.`), { code: 'control' });
  }
  if (cuitEsperado && datos.cuit !== cuitEsperado) {
    throw Object.assign(new Error(`El recibo de la cuenta ${datos.cuenta} es del CUIT ${datos.cuit} (${datos.destinatario}), no del ${cuitEsperado}.`), { code: 'control' });
  }
  Object.assign(fila, {
    cuit: datos.cuit,
    cuenta: datos.cuenta,
    destinatario: datos.destinatario,
    clase: (datos.carteles || []).map((c) => c.clase).join(' + '),
    area: (datos.carteles || []).reduce((a, c) => a + c.area, 0) || null,
    periodo: datos.periodo,
    vencimiento: datos.vencimiento,
    vencido: fechaDe(datos.vencimiento) < new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()),
    importe: datos.importe,
    recibo: datos.recibo,
  });
  if (fila.vencido) fila.observacion = `Vencio el ${datos.vencimiento}.`;

  const nombre = `${datos.periodo}-boleta marquesina- ${fila.sociedad}${varias ? ` - ${fila.cuenta}` : ''}.pdf`;
  fs.mkdirSync(SALIDA, { recursive: true });
  fs.copyFileSync(datos.pdf, path.join(SALIDA, nombre));
  fila.archivo = path.join(SALIDA, nombre);

  if (!ARCHIVAR) { fila.estado = 'OK (sin archivar)'; return; }
  let dir;
  try {
    dir = rutas.carpetaTsh(datos.cuit, datos.periodo);
  } catch (e) {
    fila.estado = 'sin carpeta';
    fila.observacion = [fila.observacion, `No archivado: ${e.message}`].filter(Boolean).join(' ');
    return;
  }
  const destino = path.join(dir, nombre);
  if (fs.existsSync(destino) && !REEMPLAZAR) {
    fila.estado = 'ya estaba';
    fila.observacion = [fila.observacion, 'Ya habia un PDF con ese nombre en TSH: no se piso (--reemplazar para pisarlo).'].filter(Boolean).join(' ');
  } else {
    fs.copyFileSync(datos.pdf, destino);
    fila.estado = 'OK';
  }
  fila.archivo = destino;
}

// Una cuenta sin recibo (ARQui: "La cuenta ingresado no existe.") no es una
// falla del script: queda como 'sin recibo' para no mezclarla con los errores.
const errorFila = (fila, e) => Object.assign(fila, {
  estado: e.code === 'sin_recibo' ? 'sin recibo' : 'error', codigo: e.code || 'error', observacion: e.message,
});

// ── Modos ─────────────────────────────────────────────────────────────────────

async function porSociedad(planilla) {
  const out = [];
  const asignadas = new Set(planilla.filas.filter((f) => !f.problema).map((f) => f.cuenta));

  // Primero se resuelve CUIT y usuario de cada una, para ir usuario por
  // usuario dejando al general (el de mas sociedades) para el final.
  const resueltas = [];
  for (const pedida of PEDIDAS) {
    const base = { sociedad: pedida.nombre || pedida.cuit, estado: 'error' };
    try {
      const cuit = cuitDe(pedida);
      base.cuit = cuit;
      base.sociedad = nombreDe(cuit, pedida.nombre);
      resueltas.push({ base, cuit, cred: credencialesArqui({ cuit }) });
    } catch (e) { out.push(errorFila(base, e)); }
  }
  const usos = {};
  for (const r of resueltas) usos[r.cred.usuario] = (usos[r.cred.usuario] || 0) + 1;
  let general = null;
  try { general = (accesosQuilmes()[0] || {}).usuario; } catch { /* solo .env: un usuario */ }
  const peso = (r) => (r.cred.usuario === general ? 1e6 : 0) + usos[r.cred.usuario];
  resueltas.sort((a, b) => peso(a) - peso(b) || (a.cred.usuario < b.cred.usuario ? -1 : a.cred.usuario > b.cred.usuario ? 1 : 0));

  for (const { base, cuit, cred } of resueltas) {

    let cuentas = CUENTA ? [CUENTA] : planilla.filas.filter((f) => f.cuit === cuit && !f.problema).map((f) => f.cuenta);
    let observacion = '';
    if (!cuentas.length) {
      const conProblema = planilla.filas.find((f) => f.cuit === cuit && f.problema);
      observacion = conProblema
        ? `En la planilla de tasas la fila ${conProblema.fila} tiene ${conProblema.problema}.`
        : 'No esta en el bloque PUBLICIDAD Y PROPAGANDA de la planilla de tasas.';
      log(`${base.sociedad}: ${observacion} Busco su recibo entre las cuentas de ARQui que la planilla no asigna.`);
      try {
        const libres = (await conSesion(cred, (page) => arqui.cuentasPyp(page))).filter((c) => !asignadas.has(c));
        for (const c of libres) {
          try { if ((await recibo(cred, c)).cuit === cuit) cuentas.push(c); } catch (e) { log(`  cuenta ${c}: ${e.message}`); }
        }
      } catch (e) { out.push(errorFila(base, e)); continue; }
      if (!cuentas.length) {
        out.push(Object.assign(base, { estado: 'sin cuenta', observacion: `${observacion} Tampoco aparece entre las cuentas de ARQui de este usuario.` }));
        continue;
      }
      observacion += ` Encontrada en ARQui: ${cuentas.join(', ')} (cargarla en la planilla).`;
    }

    for (const cuenta of cuentas) {
      const fila = { ...base, cuenta };
      try {
        log(`${fila.sociedad}: cuenta ${cuenta}...`);
        guardar(fila, await recibo(cred, cuenta), { cuitEsperado: cuit, varias: cuentas.length > 1 });
        if (observacion) fila.observacion = [observacion.trim(), fila.observacion].filter(Boolean).join(' ');
        log(`  ${fila.estado}: $ ${fila.importe} vence ${fila.vencimiento}`);
      } catch (e) {
        errorFila(fila, e);
        log(`  error: ${e.message}`);
      }
      out.push(fila);
    }
  }
  return { recibos: out };
}

function accesos() {
  if ((process.env.ARQUI_USUARIO || '').trim() && process.env.ARQUI_CLAVE) return [credencialesArqui()];
  // El general (el de mas sociedades) va ultimo: ver "Sesion de ARQui".
  const a = accesosQuilmes().reverse();
  if (!a.length) throw new Error('No hay ningun usuario de ARQui en Claves_Organismos.xlsx (grupo MUNICIPAL con municipio Quilmes).');
  return a.map((x, i) => ({ ...x, origen: `usuario ${i + 1} de ${a.length} de la planilla` + (i === a.length - 1 ? ' (general)' : '') }));
}

async function todas(planilla, { soloListar }) {
  const porCuenta = new Map(planilla.filas.filter((f) => !f.problema).map((f) => [f.cuenta, f]));
  // Usuario por usuario: listar sus cuentas y (si no es --listar) bajarlas
  // antes de loguear al siguiente. ARQui suele bloquear (403) el login de un
  // segundo usuario poco despues del primero; asi, si pasa, lo del usuario
  // general (casi todas las cuentas) ya quedo hecho. Se baja todo antes de
  // archivar: para nombrar el archivo hay que saber si la sociedad tiene mas de
  // una cuenta.
  const vistas = new Map(); // cuenta -> cred
  const bajados = new Map(); // cuenta -> datos | Error
  for (const cred of accesos()) {
    let cuentas;
    try {
      cuentas = await conSesion(cred, (page) => arqui.cuentasPyp(page));
      log(`${cred.origen}: ${cuentas.length} cuentas de PyP`);
    } catch (e) {
      log(`${cred.origen}: ${e.message}`);
      continue;
    }
    for (const c of cuentas) {
      if (vistas.has(c)) continue;
      vistas.set(c, cred);
      if (soloListar) continue;
      log(`cuenta ${c}...`);
      try { bajados.set(c, await recibo(cred, c)); } catch (e) { bajados.set(c, e); log(`  error: ${e.message}`); }
    }
  }

  if (soloListar) {
    return {
      cuentas: [...vistas.keys()].map((c) => {
        const p = porCuenta.get(c);
        return { cuenta: c, planilla: p ? `${p.nombre} (${p.cuit})` : null };
      }),
      planilla_sin_arqui: planilla.filas.filter((f) => !f.problema && !vistas.has(f.cuenta)),
      planilla_con_problemas: planilla.filas.filter((f) => f.problema),
    };
  }

  const cuentasPorCuit = {};
  for (const d of bajados.values()) if (!(d instanceof Error)) cuentasPorCuit[d.cuit] = (cuentasPorCuit[d.cuit] || 0) + 1;

  const out = [], control = [];
  for (const [cuenta, d] of bajados) {
    const fila = { sociedad: '', cuenta };
    const datos = d instanceof Error ? null : d;
    try {
      if (!datos) throw d;
      fila.sociedad = nombreDe(datos.cuit, datos.destinatario);
      guardar(fila, datos, { varias: cuentasPorCuit[datos.cuit] > 1 });
      log(`${cuenta} ${fila.sociedad}: ${fila.estado}, $ ${fila.importe} vence ${fila.vencimiento}`);
    } catch (e) {
      errorFila(fila, e);
    }
    out.push(fila);

    const p = porCuenta.get(cuenta);
    let resultado;
    if (!datos) resultado = d.code === 'sin_recibo' ? `Sin recibo en ARQui (${d.message.split(': ').pop()})` : `No se pudo bajar el recibo: ${d.message}`;
    else if (!p) resultado = 'Falta en la planilla (bloque PUBLICIDAD Y PROPAGANDA)';
    else if (p.cuit !== datos.cuit) resultado = `CUIT distinto: la planilla dice ${p.cuit}, el recibo ${datos.cuit}`;
    else resultado = 'OK';
    control.push({
      cuenta, cuit_arqui: datos && datos.cuit, destinatario: datos && datos.destinatario,
      nombre_planilla: p && p.nombre, cuit_planilla: p && p.cuit, resultado,
    });
  }
  for (const f of planilla.filas) {
    if (f.problema) {
      control.push({ cuenta: f.cuenta, nombre_planilla: f.nombre, cuit_planilla: f.cuit, resultado: `Fila ${f.fila} de la planilla con ${f.problema}` });
    } else if (!vistas.has(f.cuenta)) {
      control.push({ cuenta: f.cuenta, nombre_planilla: f.nombre, cuit_planilla: f.cuit, resultado: 'En la planilla pero no vinculada a ningun usuario de ARQui' });
    }
  }
  return { recibos: out, control };
}

// ── Main ──────────────────────────────────────────────────────────────────────

(async () => {
  let res;
  try {
    const planilla = tasas.cuentasPyp();
    log(`planilla de tasas: ${planilla.filas.length} filas en PUBLICIDAD Y PROPAGANDA`);
    res = (TODAS || LISTAR) ? await todas(planilla, { soloListar: LISTAR }) : await porSociedad(planilla);
  } catch (e) {
    if (browser) await browser.close().catch(() => {});
    fail(e.message, e.code ? { codigo: e.code } : {});
  }
  if (browser) await browser.close().catch(() => {});

  if (res.recibos) {
    fs.mkdirSync(SALIDA, { recursive: true });
    const json = path.join(TMP, 'resumen.json');
    fs.writeFileSync(json, JSON.stringify(res), 'utf-8');
    const stamp = new Date().toTimeString().slice(0, 8).replace(/:/g, '');
    const xlsx = path.join(SALIDA, `RECIBOS PYP QUILMES ${FECHA} ${stamp}.xlsx`);
    try {
      execFileSync(PYTHON, [path.join(__dirname, 'recibo_pyp.py'), 'excel', json, xlsx], { encoding: 'utf-8' });
      res.resumen = xlsx;
    } catch (e) {
      res.resumen_error = `No pude armar el Excel de resumen: ${e.message}`;
    }
    const n = (est) => res.recibos.filter((r) => r.estado && r.estado.startsWith(est)).length;
    res.totales = {
      recibos: res.recibos.length, ok: n('OK'), ya_estaban: n('ya estaba'), sin_carpeta: n('sin carpeta'),
      sin_cuenta: n('sin cuenta'), sin_recibo: n('sin recibo'), errores: n('error'), bloqueado: Boolean(bloqueado),
      importe: Math.round(res.recibos.reduce((a, r) => a + (r.importe || 0), 0) * 100) / 100,
    };
  }
  fs.rmSync(TMP, { recursive: true, force: true });
  console.log(JSON.stringify(LISTAR ? { ok: true, ...res } : { ok: true, archivado: ARCHIVAR, salida: SALIDA, ...res }, null, 2));
})();
