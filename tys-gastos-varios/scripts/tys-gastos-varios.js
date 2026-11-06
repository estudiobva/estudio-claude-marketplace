#!/usr/bin/env node
/**
 * tys-gastos-varios.js — baja de T&S Web (Zetti) el reporte 4.5.2 Seguimiento
 * de Valores filtrado por el tipo de valor 56 GAVA - GASTOS VARIOS, lo parte por
 * mes (fecha de emision), lo cruza con el EERR de la sociedad y archiva cada mes
 * en la carpeta del mes de la sociedad:
 *
 *   [Sociedad]/01-Impuestos Mensuales/[AAAA/]AAAAMM/AAAAMM - ZETTI - GASTOS VARIOS.xlsx
 *
 *   node scripts/tys-gastos-varios.js --nombre="<sociedad>"                 # ejercicio hasta el mes anterior
 *   node scripts/tys-gastos-varios.js --nombre="<sociedad>" --periodo=09/2026
 *   node scripts/tys-gastos-varios.js --cuit=<CUIT> --desde=01/01/2026 --hasta=30/09/2026
 *   node scripts/tys-gastos-varios.js --todas --periodo=10/2026
 *   node scripts/tys-gastos-varios.js --listar
 *
 * Argumentos:
 *   --nombre / --cuit   Una sociedad (titular en Claves_Organismos.xlsx, o CUIT).
 *   --nombres / --cuits Varias, separadas por ";".
 *   --todas             Todas las sociedades de Zetti con carpeta en la unidad.
 *   --periodo           Un mes: MM/AAAA o AAAAMM.
 *   --desde / --hasta   Rango de emision DD/MM/AAAA (un archivo por mes con valores).
 *                       Sin --periodo ni --desde: el ejercicio del EERR de la
 *                       sociedad hasta el mes anterior (mes vencido); sin EERR,
 *                       el anio calendario hasta el mes anterior.
 *   --si-existe         reemplazar (default) | saltear | abortar, si el mes ya tiene el archivo.
 *   --no-archivar       Deja los Excel en --salida (default $BVA_SALIDAS_PATH/gastos-varios) y no toca la unidad.
 *   --entidad           Nombre exacto de la entidad de T&S Web, si el CUIT no alcanza
 *                       (p.ej. un solo local en vez de la Sociedad).
 *   --sin-cruce         No busca ni lee el EERR.
 *   --ver               Navegador visible (solo si se pide).
 *
 * Login: el mismo de los otros scripts de Zetti (grupo ZETTI de
 * Claves_Organismos.xlsx, fila "Zetti T&S Web (todas las sociedades)", o
 * ZETTI_USUARIO / ZETTI_CLAVE en el .env). Un solo login para todo el lote.
 *
 * Replica la carga a mano de la habilidad: abre el 4.5.2 con el nodo de la
 * Sociedad (la entidad de mas arriba con ese CUIT, abarca todos los locales),
 * Tipos Valor Asociados = 56 GAVA por la lupa, emision desde/hasta, el resto
 * vacio, orden Valor, CSV. Un solo reporte para todo el rango: se controla
 * entero (detalle vs total de cada concepto, total general, fechas) y despues
 * se parte por mes. Los meses sin valores no generan archivo. No escribe nada
 * en Zetti ni en el EERR.
 *
 * Cada archivo: hojas Resumen (conceptos del mes), Detalle y Cruce EERR.
 * JSON por stdout; progreso por stderr.
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
const { credencialesZetti } = require('../lib/zetti-claves');
const zetti = require('../lib/zetti');
const rutas = require('../lib/rutas-bva');

const args = parseArgs();
const VER = 'ver' in args;
const LISTAR = 'listar' in args;
const TODAS = 'todas' in args;
const CRUCE = !('sin-cruce' in args);
const ARCHIVAR = !('no-archivar' in args);
const SI_EXISTE = String(args['si-existe'] || 'reemplazar').toLowerCase();
if (!['abortar', 'saltear', 'reemplazar'].includes(SI_EXISTE)) fail('--si-existe tiene que ser reemplazar, saltear o abortar');
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const PY = path.join(__dirname, 'gastos_varios_a_excel.py');

if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';
loadEnv();

const RE_FECHA = /^\d{2}\/\d{2}\/\d{4}$/;
for (const k of ['desde', 'hasta']) {
  if (args[k] && !RE_FECHA.test(args[k])) fail(`--${k} tiene que ser DD/MM/AAAA`);
}
if (!!args['desde'] !== !!args['hasta']) fail('Pasa --desde y --hasta juntos');
if (args['periodo'] && args['desde']) fail('Usa --periodo o --desde/--hasta, no los dos');

const pad = (n) => String(n).padStart(2, '0');
const ultimoDia = (a, m) => new Date(a, m, 0).getDate();
let RANGO_FIJO = null;
if (args['periodo']) {
  const s = String(args['periodo']).trim();
  const m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/) || s.match(/^(\d{4})(\d{2})$/);
  if (!m) fail('--periodo tiene que ser MM/AAAA o AAAAMM');
  const [anio, mes] = s.includes('/') || s.includes('-') || s.includes('.') ? [Number(m[2]), Number(m[1])] : [Number(m[1]), Number(m[2])];
  RANGO_FIJO = { desde: `01/${pad(mes)}/${anio}`, hasta: `${pad(ultimoDia(anio, mes))}/${pad(mes)}/${anio}` };
} else if (args['desde']) {
  RANGO_FIJO = { desde: args['desde'], hasta: args['hasta'] };
}
// Ultimo dia del mes anterior: tope del ejercicio por defecto (mes vencido).
const hoy = new Date();
const TOPE = new Date(hoy.getFullYear(), hoy.getMonth(), 0);
const aFecha = (s) => { const [d, m, a] = s.split('/').map(Number); return new Date(a, m - 1, d); };
const aTexto = (d) => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;

const SALIDAS = process.env.BVA_SALIDAS_PATH && process.env.BVA_SALIDAS_PATH.trim();
const DIR_SALIDA = args['salida'] || (SALIDAS ? path.join(SALIDAS, 'gastos-varios') : path.join(require('os').homedir(), 'Documents', 'BVA-salidas', 'gastos-varios'));
const nombreArchivo = (aaaamm) => `${aaaamm} - ZETTI - GASTOS VARIOS.xlsx`;

// Carpeta del mes de la sociedad (01-Impuestos Mensuales/AAAAMM, o /AAAA/AAAAMM
// si la sociedad tiene carpeta de anio), creandola si falta.
function carpetaDelMes(cuit, aaaamm) {
  const mensuales = path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales');
  const dir = rutas.carpetaMes(mensuales, aaaamm);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length && !TODAS && !LISTAR) fail('Falta --nombre, --cuit, --nombres, --cuits, --todas o --listar');
if (args['entidad'] && (PEDIDAS.length !== 1 || TODAS)) fail('--entidad es para una sola sociedad');

let prefijo = '';
const log = (m) => process.stderr.write(`[tys-gastos-varios]${prefijo} ${m}\n`);

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

function py(argv) {
  try {
    const out = execFileSync(PYTHON, [PY, ...argv], { encoding: 'utf-8', maxBuffer: 64 * 1024 * 1024 });
    return JSON.parse(out.trim().split(/\r?\n/).pop());
  } catch (e) {
    const salida = String(e.stdout || '').trim().split(/\r?\n/).pop();
    let det = null; try { det = JSON.parse(salida); } catch { /* no era JSON */ }
    throw Object.assign(new Error(det && det.error ? det.error : String(e.stderr || e.message).trim()), { code: 'control', detalle: det || undefined });
  }
}

// EERR de la sociedad (carpeta "0 - ESTUDIO - 00000000000/EERR 2026" de la unidad).
function eerrDe(cuit) {
  let raiz;
  try { raiz = rutas.raiz(); } catch { return { eerr: null }; }
  const estudio = fs.readdirSync(raiz).find((d) => /^0 - ESTUDIO/i.test(d));
  if (!estudio) return { eerr: null };
  const dir = path.join(raiz, estudio, 'EERR 2026');
  if (!fs.existsSync(dir)) return { eerr: null };
  let nombre = '';
  try {
    const base = path.basename(rutas.carpetaSociedad(cuit)).normalize('NFC');
    nombre = base.replace(/^\d+\s*-\s*/, '').replace(/\s*-\s*\d{11}$/, '');
  } catch { /* sin carpeta */ }
  return py(['buscar-eerr', `--cuit=${cuit}`, `--eerr-dir=${dir}`, `--nombre=${nombre}`]);
}

const aaaammdd = (f) => f.split('/').reverse().join('');

async function unaSociedad(ctx, auth, ents, pedida, tmp) {
  const fila = { sociedad: pedida.nombre || pedida.cuit, cuit: null, ok: false };
  try {
    const cuit = cuitDe(pedida);
    fila.cuit = cuit;

    let entidad, locales = [];
    if (args['entidad']) {
      entidad = ents.find((e) => e.nombre === args['entidad'] || e.corto === args['entidad']);
      if (!entidad) throw Object.assign(new Error(`No hay ninguna entidad "${args['entidad']}" en T&S Web.`), { code: 'sin_entidad' });
      if (entidad.cuit !== cuit) throw Object.assign(new Error(`La entidad ${entidad.nombre} es del CUIT ${entidad.cuit}, no del ${cuit}.`), { code: 'sin_entidad' });
    } else {
      ({ entidad, locales } = zetti.entidadDeCuit(ents, cuit));
    }
    fila.entidad = entidad.nombre;
    fila.nodo = entidad.codigo;
    if (!pedida.nombre) fila.sociedad = entidad.razon || entidad.nombre;

    if (ARCHIVAR) {
      try { rutas.carpetaSociedad(cuit); } catch (e) { throw Object.assign(e, { code: 'sin_carpeta' }); }
    }
    const eerr = CRUCE ? eerrDe(cuit) : { eerr: null };
    let { desde, hasta } = RANGO_FIJO || {};
    if (!desde) {
      if (eerr.eerr) {
        ({ desde, hasta } = eerr);
      } else {
        // Sin EERR: anio calendario del mes anterior.
        desde = `01/01/${TOPE.getFullYear()}`; hasta = aTexto(TOPE);
        fila.periodo_sin_eerr = true;
        log(`sin EERR: tomo el anio calendario, ${desde} a ${hasta}`);
      }
      if (aFecha(hasta) > TOPE) hasta = aTexto(TOPE);
      if (aFecha(desde) > aFecha(hasta)) throw Object.assign(new Error(`El ejercicio del EERR empieza el ${desde}: todavia no hay meses vencidos. Pasa --periodo.`), { code: 'sin_periodo' });
    }
    fila.desde = desde; fila.hasta = hasta;
    if (CRUCE && !eerr.eerr) log('sin EERR para cruzar: los archivos van sin la hoja Cruce EERR');

    log(`${entidad.nombre} (${entidad.codigo}${locales.length ? `, ${locales.length} local/es` : ''}): 4.5.2 GAVA emision ${desde} a ${hasta}`);
    const dir = fs.mkdtempSync(path.join(tmp, 'soc-'));
    const csv = path.join(dir, 'reporte.csv');
    await zetti.seguimientoValores(ctx, auth, entidad, { desde, hasta, tipos: ['GAVA'] }, csv);

    const dirMeses = path.join(dir, 'meses');
    let res;
    try {
      res = py(['armar', csv, dirMeses, `--cuit=${cuit}`, `--desde=${desde}`, `--hasta=${hasta}`,
        `--entidad=${entidad.nombre} (${entidad.codigo})`, `--razon=${entidad.razon || ''}`, ...(eerr.eerr ? [`--eerr=${eerr.eerr}`] : [])]);
    } catch (e) {
      // Dejar el CSV de T&S para revisarlo.
      fs.mkdirSync(DIR_SALIDA, { recursive: true });
      const copia = path.join(DIR_SALIDA, `${cuit} - ${aaaammdd(desde)}-${aaaammdd(hasta)} - REPORTE GAVA CON ERROR.csv`);
      fs.copyFileSync(csv, copia);
      e.message += ` (CSV en ${copia})`;
      throw e;
    }

    // Archivar cada mes en su carpeta.
    for (const mes of res.meses) {
      const destino = ARCHIVAR
        ? path.join(carpetaDelMes(cuit, mes.periodo), nombreArchivo(mes.periodo))
        : path.join(DIR_SALIDA, `${cuit} - ${nombreArchivo(mes.periodo)}`);
      if (!ARCHIVAR) fs.mkdirSync(DIR_SALIDA, { recursive: true });
      if (fs.existsSync(destino) && SI_EXISTE !== 'reemplazar') {
        if (SI_EXISTE === 'abortar') throw Object.assign(new Error(`Ya existe ${destino}. No lo piso: --si-existe=reemplazar o --si-existe=saltear.`), { code: 'ya_existe' });
        log(`ya existe ${destino}: lo salteo`);
        mes.salteado = true;
      } else {
        mes.reemplazado = fs.existsSync(destino) || undefined;
        fs.copyFileSync(mes.archivo_tmp, destino);
      }
      delete mes.archivo_tmp;
      mes.archivo = destino;
    }
    fs.rmSync(dir, { recursive: true, force: true });

    const $ = (n) => Number(n || 0).toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    log(`${res.valores} valores en ${res.conceptos} conceptos, total ${$(res.total_sin_ignorados)} sin IGNORADO (${res.ignorados.cantidad}), ` +
        `${res.ingresados.cantidad} INGRESADO -> ${res.meses.length} mes/es guardado/s` +
        (res.meses_sin_valores.length ? ` (sin valores: ${res.meses_sin_valores.join(', ')})` : ''));
    return { ...fila, ok: true, ...res };
  } catch (e) {
    log(`ERROR ${e.message}`);
    return { ...fila, ok: false, code: e.code || 'error', error: e.message, ...(e.detalle ? { detalle: e.detalle } : {}) };
  }
}

async function procesar(ctx, auth, ents, tmp) {
  const pedidas = [...PEDIDAS];
  if (TODAS) {
    const vistos = new Set(pedidas.map((p) => p.cuit).filter(Boolean));
    for (const c of [...new Set(ents.map((e) => e.cuit).filter(Boolean))]) {
      if (vistos.has(c)) continue;
      try { rutas.carpetaSociedad(c); } catch { continue; }   // sin carpeta en la unidad: no es del estudio
      pedidas.push({ cuit: c });
    }
  }
  const resultados = [];
  for (const [i, pedida] of pedidas.entries()) {
    prefijo = pedidas.length > 1 ? ` [${i + 1}/${pedidas.length}]` : '';
    resultados.push(await unaSociedad(ctx, auth, ents, pedida, tmp));
  }
  prefijo = '';
  const ok = resultados.every((r) => r.ok);
  return pedidas.length === 1 ? resultados[0] : {
    ok, procesadas: resultados.length,
    con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, error: r.error })),
    resultados,
  };
}

(async () => {
  let cred;
  try { cred = credencialesZetti(); } catch (e) { fail(e.message, { code: 'sin_credenciales' }); }

  const browser = await chromium.launch(launchOptions());
  const ctx = await browser.newContext({ acceptDownloads: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'bva-gava-'));
  let salida, ok = false;
  try {
    let auth, page;
    try { ({ page, auth } = await zetti.login(ctx, cred, log)); } catch (e) {
      fail(e.code === 'clave_rechazada'
        ? 'T&S Web rechazo el usuario o la clave de Zetti: revisa la fila ZETTI de Claves_Organismos.xlsx.'
        : e.message, { code: e.code || 'login' });
    }
    log('leyendo entidades y CUITs');
    const ents = await zetti.entidades(page);
    if (LISTAR) {
      salida = ents.map((e) => ({ nombre: e.nombre, corto: e.corto, nivel: e.nivel, codigo: e.codigo, cuit: e.cuit, razon: e.razon }));
      ok = true;
    } else {
      salida = await procesar(ctx, auth, ents, tmp);
      ok = salida.ok;
    }
  } finally {
    await browser.close().catch(() => {});
    fs.rmSync(tmp, { recursive: true, force: true });
  }
  console.log(JSON.stringify(salida, null, 2));
  process.exit(ok ? 0 : 1);
})().catch((e) => fail(e.message));
