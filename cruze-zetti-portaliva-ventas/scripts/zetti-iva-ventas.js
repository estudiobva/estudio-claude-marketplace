#!/usr/bin/env node
/**
 * zetti-iva-ventas.js — baja el Subdiario de IVA Ventas de Zetti (T&S Web,
 * reporte 5.6.5) de una o varias sociedades para un mes, lo pasa a Excel limpio
 * y ordenado por tipo de comprobante, y lo archiva en la unidad compartida al
 * lado del Libro IVA Ventas del Portal IVA.
 *
 *   node scripts/zetti-iva-ventas.js --nombre="<sociedad>" --periodo=09/2026
 *   node scripts/zetti-iva-ventas.js --cuit=<CUIT>
 *   node scripts/zetti-iva-ventas.js --nombres="<A>;<B>" --periodo=09/2026
 *   node scripts/zetti-iva-ventas.js --todas --periodo=09/2026 --si-existe=saltear
 *   node scripts/zetti-iva-ventas.js --listar
 *
 * Argumentos:
 *   --nombre / --cuit   Una sociedad (titular en Claves_Organismos.xlsx, o CUIT).
 *   --nombres / --cuits Varias, separadas por ";".
 *   --todas             Todas las sociedades de Zetti que tienen carpeta en la unidad.
 *   --periodo           MM/AAAA o AAAAMM. Default: el mes anterior (mes vencido).
 *   --entidad           Nombre exacto de la entidad de T&S Web, si el CUIT no alcanza.
 *   --si-existe         abortar (default) | saltear | reemplazar, si ya esta el .xlsx.
 *   --no-archivar       Deja el .xlsx en ~/Documents/BVA-salidas/zetti ($BVA_SALIDAS_PATH) y no toca el Drive.
 *   --listar            Solo lista las entidades de T&S Web con su CUIT.
 *   --ver               Navegador visible (solo si se pide).
 *
 * Login: usuario y clave de Zetti del grupo ZETTI de Claves_Organismos.xlsx
 * (fila "Zetti T&S Web (todas las sociedades)"), o ZETTI_USUARIO / ZETTI_CLAVE
 * en el .env. Un solo login para todo el lote.
 *
 * La entidad de T&S Web se elige por CUIT (la API de nodos lo trae): la de mas
 * arriba del arbol, que abarca todos los locales de la sociedad. Antes de
 * archivar se controla que el CUIT del encabezado del reporte sea el pedido,
 * que todas las fechas sean del periodo y que la suma de los comprobantes de el
 * Total del reporte.
 *
 * Archiva [IVA del periodo]/AAAAMM - ZETTI - VENTAS.xlsx (la misma carpeta que
 * "AAAAMM - PORTAL IVA - VENTAS.xlsx"). El CSV se borra.
 *
 * Salida: JSON por stdout; progreso por stderr. Solo lee: no modifica nada en Zetti.
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
const ARCHIVAR = !('no-archivar' in args);
const LISTAR = 'listar' in args;
const TODAS = 'todas' in args;
const SI_EXISTE = String(args['si-existe'] || 'abortar').toLowerCase();
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');

if (!['abortar', 'saltear', 'reemplazar'].includes(SI_EXISTE)) fail('--si-existe tiene que ser abortar, saltear o reemplazar');
if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';
loadEnv();

function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)))          return s;
  if ((m = s.match(/^(\d{4})[\/\-.](\d{1,2})$/))) return m[1] + String(m[2]).padStart(2, '0');
  return null;
}
function mesAnterior(hoy = new Date()) {
  const d = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
}
const PERIODO = args['periodo'] ? normalizarPeriodo(args['periodo']) : mesAnterior();
if (!PERIODO) fail('--periodo tiene que ser MM/AAAA o AAAAMM');
const ANIO = Number(PERIODO.slice(0, 4));
const MES = Number(PERIODO.slice(4));
const ULTIMO = new Date(ANIO, MES, 0).getDate();
const DESDE = `01/${PERIODO.slice(4)}/${ANIO}`;
const HASTA = `${String(ULTIMO).padStart(2, '0')}/${PERIODO.slice(4)}/${ANIO}`;
const NOMBRE_ARCHIVO = `${PERIODO} - ZETTI - VENTAS.xlsx`;

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length && !TODAS && !LISTAR) fail('Falta --nombre, --cuit, --nombres, --cuits, --todas o --listar');
if (args['entidad'] && (PEDIDAS.length !== 1 || TODAS)) fail('--entidad es para una sola sociedad');

let prefijo = '';
const log = (m) => process.stderr.write(`[zetti-iva-ventas]${prefijo} ${m}\n`);

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

function carpetaDestino(cuit) {
  if (!ARCHIVAR) {
    const dir = path.join(process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'zetti');
    fs.mkdirSync(dir, { recursive: true });
    return dir;
  }
  return rutas.carpetaIva(cuit, PERIODO);
}

async function unaSociedad(ctx, auth, ents, pedida, tmp) {
  const fila = { sociedad: pedida.nombre || pedida.cuit, cuit: null, periodo: PERIODO, ok: false };
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
    if (locales.length) fila.locales = locales.map((l) => l.nombre);
    if (!pedida.nombre) fila.sociedad = entidad.razon || entidad.nombre;

    let carpeta;
    try { carpeta = carpetaDestino(cuit); } catch (e) { throw Object.assign(e, { code: 'sin_carpeta' }); }
    const destino = ARCHIVAR
      ? path.join(carpeta, NOMBRE_ARCHIVO)
      : path.join(carpeta, `${cuit} - ${NOMBRE_ARCHIVO}`);
    if (fs.existsSync(destino)) {
      if (SI_EXISTE === 'saltear') {
        log(`ya existe ${destino}: lo salteo`);
        return { ...fila, ok: true, salteado: true, archivo: destino };
      }
      if (SI_EXISTE === 'abortar') throw Object.assign(new Error(`Ya existe ${destino}. No lo piso: --si-existe=reemplazar o --si-existe=saltear.`), { code: 'ya_existe' });
    }

    log(`${entidad.nombre} (${entidad.codigo}${locales.length ? `, ${locales.length} local/es` : ''}): reporte 5.6.5 del ${DESDE} al ${HASTA}`);
    const dir = fs.mkdtempSync(path.join(tmp, 'soc-'));
    const csv = await zetti.subdiarioIvaVentas(ctx, auth, entidad, { desde: DESDE, hasta: HASTA }, path.join(dir, 'reporte.csv'));

    const xlsx = path.join(dir, NOMBRE_ARCHIVO);
    let res;
    try {
      res = JSON.parse(execFileSync(PYTHON, [path.join(__dirname, 'zetti_a_excel.py'), csv, xlsx, `--periodo=${PERIODO}`], { encoding: 'utf-8' }).trim().split(/\r?\n/).pop());
    } catch (e) {
      const salida = String(e.stdout || '').trim().split(/\r?\n/).pop();
      let det = null; try { det = JSON.parse(salida); } catch { /* no era JSON */ }
      throw Object.assign(new Error(det && det.error ? det.error : (e.stderr || e.message).toString().trim()), { code: 'control', detalle: det || undefined });
    }
    if (res.cuit !== cuit) {
      throw Object.assign(new Error(`El reporte salio con el CUIT ${res.cuit} (${res.razon_social}) y se pidio ${cuit}: no lo archivo.`), { code: 'control' });
    }
    if (res.desde !== DESDE || res.hasta !== HASTA) {
      throw Object.assign(new Error(`El reporte salio del ${res.desde} al ${res.hasta} y se pidio del ${DESDE} al ${HASTA}: no lo archivo.`), { code: 'control' });
    }

    fs.copyFileSync(xlsx, destino);
    fs.rmSync(dir, { recursive: true, force: true });
    log(`${res.comprobantes} comprobantes ${JSON.stringify(res.por_tc)}, total ${res.totales.total.toLocaleString('es-AR', { minimumFractionDigits: 2 })} -> ${destino}`);
    return {
      ...fila, ok: true, razon_social: res.razon_social, comprobantes: res.comprobantes, por_tc: res.por_tc,
      totales: res.totales, control_total_reporte: res.control_total_reporte,
      ...(res.diferencias_total_reporte ? { diferencias_total_reporte: res.diferencias_total_reporte } : {}),
      archivo: destino,
    };
  } catch (e) {
    log(`ERROR ${e.message}`);
    return { ...fila, ok: false, code: e.code || 'error', error: e.message, ...(e.detalle ? { detalle: e.detalle } : {}) };
  }
}

(async () => {
  let cred;
  try { cred = credencialesZetti(); } catch (e) { fail(e.message, { code: 'sin_credenciales' }); }

  const browser = await chromium.launch(launchOptions());
  const ctx = await browser.newContext({ acceptDownloads: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'bva-zetti-'));
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
    ok, periodo: PERIODO, procesadas: resultados.length,
    con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, error: r.error })),
    resultados,
  };
}
