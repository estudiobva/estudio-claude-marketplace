#!/usr/bin/env node
/**
 * portal-iva-descarga.js — baja el Libro IVA Compras y Ventas de una o varias
 * sociedades para un periodo, los pasa a Excel y los archiva en la unidad
 * compartida. Es el paso 1 del circuito NCR (portal-iva.js), hecho para los dos
 * libros con un solo login por sociedad.
 *
 *   node scripts/portal-iva-descarga.js --nombre="<sociedad>" --periodo=08/2026
 *   node scripts/portal-iva-descarga.js --cuit=<CUIT> --periodo=08/2026 --libros=ventas
 *   node scripts/portal-iva-descarga.js --nombres="<sociedad A>;<sociedad B>" --periodo=08/2026
 *
 * Argumentos:
 *   --nombre / --cuit   Una sociedad (titular en Claves_Organismos.xlsx, o CUIT).
 *   --nombres           Varias, separadas por ";". Tambien --cuits="A;B".
 *   --periodo           MM/AAAA o AAAAMM. Default: el mes anterior (se trabaja a mes vencido).
 *   --libros            compras,ventas (default los dos, en ese orden).
 *   --si-existe         abortar (default) | saltear | reemplazar, si ya esta el .xlsx.
 *   --aunque-presentada Entra aunque haya acuse de la DDJJ IVA del periodo.
 *   --no-archivar       Deja los .xlsx en ~/Documents/BVA-salidas/portal-iva ($BVA_SALIDAS_PATH) y no toca el Drive.
 *   --ver               Navegador visible.
 *
 * Por libro: importa desde ARCA al borrador, espera a que la importacion
 * termine, baja el CSV, lo convierte con csv_a_excel.py y archiva
 *   [IVA del periodo]/AAAAMM - PORTAL IVA - COMPRAS.xlsx
 *   [IVA del periodo]/AAAAMM - PORTAL IVA - VENTAS.xlsx
 * El CSV se borra: queda solo el Excel. conciliar-ncr.py lee ese .xlsx igual
 * que el CSV.
 *
 * Salida: JSON por stdout; progreso por stderr. NO presenta la DDJJ.
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { resolveCredentials, envFile } = require('../lib/env');
const { login } = require('../lib/arca-login');
const rutas = require('../lib/rutas-bva');
const pi    = require('../lib/portal-iva');

const args = parseArgs();
const VER = 'ver' in args;
const ARCHIVAR = !('no-archivar' in args);
const AUNQUE_PRESENTADA = 'aunque-presentada' in args;
const SI_EXISTE = String(args['si-existe'] || 'abortar').toLowerCase();
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');

if (!['abortar', 'saltear', 'reemplazar'].includes(SI_EXISTE)) fail('--si-existe tiene que ser abortar, saltear o reemplazar');
if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';

const LIBROS = String(args['libros'] || 'compras,ventas').toLowerCase().split(/[,;\s]+/).filter(Boolean);
if (!LIBROS.length || LIBROS.some((l) => !['compras', 'ventas'].includes(l))) fail('--libros acepta compras, ventas o compras,ventas');

// Acepta MM/AAAA y AAAAMM; internamente todo es AAAAMM.
function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)))          return s;
  if ((m = s.match(/^(\d{4})[\/\-.](\d{1,2})$/))) return m[1] + String(m[2]).padStart(2, '0');
  return null;
}
// Se trabaja a mes vencido: sin --periodo, el mes anterior al de hoy (se corre
// los primeros dias del mes siguiente, antes de presentar).
function mesAnterior(hoy = new Date()) {
  const d = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
}
const PERIODO = args['periodo'] ? normalizarPeriodo(args['periodo']) : mesAnterior();
if (!PERIODO) fail('--periodo tiene que ser MM/AAAA o AAAAMM');

// Sociedades a procesar, cada una como { cuit } o { nombre }.
const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length) fail('Falta --nombre, --cuit, --nombres o --cuits');

let prefijo = '';
const log = (m) => process.stderr.write(`[portal-iva-descarga]${prefijo} ${m}\n`);

// La DDJJ de IVA ya se presento si en la carpeta del mes (o una subcarpeta) hay
// un acuse "AAAAMM-IVA-ACUSE-...". Con el periodo presentado, "Nueva declaracion
// jurada" puede abrir una rectificativa: por eso no se entra. (Mismo criterio
// que ncr-lote.js.)
function acuseIva(cuit) {
  let mes;
  try { mes = rutas.carpetaMes(path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales'), PERIODO); } catch { return null; }
  if (!fs.existsSync(mes)) return null;
  const re = new RegExp(`^${PERIODO}.*\\bIVA\\b.*ACUSE`, 'i');
  const dirs = [mes, ...fs.readdirSync(mes, { withFileTypes: true }).filter((e) => e.isDirectory()).map((e) => path.join(mes, e.name))];
  for (const d of dirs) {
    const hit = fs.readdirSync(d).find((f) => re.test(f));
    if (hit) return path.relative(mes, path.join(d, hit));
  }
  return null;
}

// Desde el Libro abierto, vuelve al menu del borrador (las tarjetas Libro
// Compras / Libro Ventas). abrirLibro() no clickea si la URL ya es verCompras o
// verVentas, asi que sin esto el segundo libro se bajaria del primero.
async function volverAlMenu(iva) {
  if (!/ver(Compras|Ventas)\.do/i.test(iva.url())) return;
  const tarjetas = iva.locator('a.panel, a, .panel').filter({ hasText: /Libro\s+(Compras|Ventas)/i });
  await iva.goto(iva.url().replace(/ver(Compras|Ventas)\.do.*$/i, 'menuPresentacion.do'), { waitUntil: 'domcontentloaded' });
  await iva.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  if (await tarjetas.first().isVisible({ timeout: 10000 }).catch(() => false)) return;
  log('menuPresentacion.do no mostro las tarjetas: vuelvo con el historial');
  await iva.goBack({ waitUntil: 'domcontentloaded' }).catch(() => {});
  await tarjetas.first().waitFor({ state: 'visible', timeout: 20000 });
}

// CSV (o ZIP con el CSV) -> filas de datos y fechas fuera del periodo.
function leerCsv(bajado, tmp) {
  let csv = bajado;
  if (/\.zip$/i.test(bajado)) {
    const dir = fs.mkdtempSync(path.join(tmp, 'zip-'));
    if (process.platform === 'win32') {
      execFileSync('powershell', ['-NoProfile', '-Command',
        `Expand-Archive -LiteralPath '${bajado}' -DestinationPath '${dir}' -Force`]);
    } else {
      execFileSync('unzip', ['-o', '-q', bajado, '-d', dir]);
    }
    const dentro = fs.readdirSync(dir).filter((f) => /\.csv$/i.test(f));
    if (dentro.length !== 1) throw new Error(`El ZIP traia ${dentro.length} CSV (se espera uno): ${bajado}`);
    csv = path.join(dir, dentro[0]);
  }
  // El CSV de ARCA viene en ISO-8859-1.
  const lineas = fs.readFileSync(csv, 'latin1').split(/\r?\n/).filter((l) => l.trim());
  const filas = Math.max(0, lineas.length - 1);
  const fuera = lineas.slice(1)
    .map((l) => (l.match(/(\d{4})-(\d{2})-\d{2}/) || []).slice(1, 3).join(''))
    .filter((f) => f && f !== PERIODO).length;
  return { csv, filas, fuera };
}

// Un libro, con el borrador ya abierto.
async function bajarLibro(iva, libro, destinoDir, tmp) {
  const nombre = rutas.nombreArchivo(PERIODO, libro, 'xlsx');
  const destino = ARCHIVAR ? path.join(destinoDir, nombre) : null;
  if (destino && fs.existsSync(destino) && SI_EXISTE !== 'reemplazar') {
    if (SI_EXISTE === 'saltear') return { libro, ok: true, salteado: true, archivo: destino };
    throw new Error(`Ya existe ${destino}. No lo piso: --si-existe=reemplazar o --si-existe=saltear.`);
  }

  await volverAlMenu(iva);
  await pi.abrirLibro(iva, libro);
  const importacion = await pi.importarDesdeArca(iva);

  const dirLibro = fs.mkdtempSync(path.join(tmp, `${libro}-`));
  const bajado = await pi.descargarCsv(iva, dirLibro);
  const { csv, filas, fuera } = leerCsv(bajado, dirLibro);
  log(`${libro}: CSV con ${filas} filas de datos, ${fuera} con fecha fuera del periodo`);
  if (filas === 0) throw new Error(`El CSV de ${libro} bajo vacio: no archivo nada.`);

  const xlsx = path.join(dirLibro, nombre);
  const salida = execFileSync(PYTHON, [path.join(__dirname, 'csv_a_excel.py'), csv, xlsx, `--libro=${libro}`], { encoding: 'utf-8' });
  const excel = JSON.parse(salida.trim().split(/\r?\n/).pop());
  if (excel.filas !== filas) throw new Error(`El Excel de ${libro} quedo con ${excel.filas} filas y el CSV tenia ${filas}.`);
  log(`${libro}: Excel "${excel.hoja}", ${excel.filas} filas, ${excel.formulas} formulas SUM`);

  let archivo;
  if (destino) {
    fs.copyFileSync(xlsx, destino);
    archivo = destino;
  } else {
    const dir = path.join(process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'portal-iva');
    fs.mkdirSync(dir, { recursive: true });
    archivo = path.join(dir, `${rutas.carpetaSociedad(CUIT_ACTUAL).split(path.sep).pop()} - ${nombre}`);
    fs.copyFileSync(xlsx, archivo);
  }
  // El CSV (y el ZIP) eran intermedios: se borran y queda solo el Excel.
  fs.rmSync(dirLibro, { recursive: true, force: true });
  log(`${libro}: archivado en ${archivo}`);

  // Un .csv viejo con el mismo nombre al lado confunde a conciliar-ncr.py
  // (encuentra dos archivos de compras y corta). No se borra solo: se avisa.
  const csvViejo = destino && destino.replace(/\.xlsx$/, '.csv');
  const aviso = csvViejo && fs.existsSync(csvViejo) ? `Quedo tambien ${path.basename(csvViejo)} de una bajada anterior: borralo si el .xlsx es el bueno.` : null;

  return { libro, ok: true, importacion, filas, filasFueraDePeriodo: fuera, archivo, aviso };
}

let CUIT_ACTUAL = null;

async function unaSociedad(pedida) {
  let cred;
  try {
    cred = resolveCredentials({ cuit: pedida.cuit, nombre: pedida.nombre, cuitLogin: args['cuit-login'], password: args['password'] });
  } catch (e) {
    return { sociedad: pedida.nombre || pedida.cuit, ok: false, error: e.message,
      sugerencia: `Revisa el .env (${envFile() || 'no encontrado'}) o el Excel de claves.` };
  }
  const { cuit, cuitLogin, password, titular } = cred;
  CUIT_ACTUAL = cuit;
  const fila = { sociedad: titular || pedida.nombre || cuit, cuit, periodo: PERIODO, ok: false, libros: [] };

  // La carpeta destino se resuelve ANTES de abrir el navegador: si la sociedad
  // no esta en la unidad, mejor enterarse ahora que despues del scraping.
  let destinoDir = null;
  try {
    if (ARCHIVAR) destinoDir = rutas.carpetaIva(cuit, PERIODO);
    else rutas.carpetaSociedad(cuit);
  } catch (e) { return { ...fila, error: e.message, code: 'destino_no_resuelto' }; }

  const acuse = acuseIva(cuit);
  if (acuse && !AUNQUE_PRESENTADA) {
    return { ...fila, error: `La DDJJ de IVA ${PERIODO} ya esta presentada (${acuse}): no entro, podria abrir una rectificativa. Si igual hace falta, --aunque-presentada.`, code: 'ddjj_presentada' };
  }

  // Todo lo que ya existe y se saltea no necesita login.
  const pendientes = LIBROS.filter((l) => !(ARCHIVAR && SI_EXISTE === 'saltear' && fs.existsSync(path.join(destinoDir, rutas.nombreArchivo(PERIODO, l, 'xlsx')))));
  for (const l of LIBROS.filter((x) => !pendientes.includes(x))) {
    fila.libros.push({ libro: l, ok: true, salteado: true, archivo: path.join(destinoDir, rutas.nombreArchivo(PERIODO, l, 'xlsx')) });
  }
  if (!pendientes.length) return { ...fila, ok: true };

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'bva-iva-'));
  const browser = await chromium.launch(launchOptions());
  const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
  const page = await context.newPage();
  try {
    log(`${fila.sociedad} · CUIT ${cuit} · periodo ${PERIODO} · ${pendientes.join(' + ')}`);
    log('entrando a ARCA…');
    await login(page, { cuitLogin, password });
    const iva = await pi.abrirPortalIva(context, page);
    fila.representando = (await pi.asegurarRepresentacion(iva, cuit)).linea;
    await pi.abrirBorrador(iva, PERIODO);
    fila.datosInicialesConfigurados = await pi.datosIniciales(iva);

    for (const libro of pendientes) {
      try {
        fila.libros.push(await bajarLibro(iva, libro, destinoDir, tmp));
      } catch (e) {
        // Un libro que falla no frena al otro, salvo que sea un corte de ARCA.
        if (['captcha', 'credenciales_invalidas', 'otro_borrador_abierto'].includes(e.code)) throw e;
        log(`${libro}: ERROR ${e.message}`);
        fila.libros.push({ libro, ok: false, error: e.message, code: e.code || 'error' });
      }
    }
    fila.ok = fila.libros.every((l) => l.ok);
  } catch (e) {
    fila.code = e.code || 'error';
    fila.error = e.code === 'captcha'
      ? `ARCA pidio un captcha. Corre \`node scripts/login.js --cuit=${cuit} --quedate\`, resolvelo a mano, y reintenta.`
      : e.code === 'credenciales_invalidas'
        ? `ARCA rechazo las credenciales del CUIT ${cuit}. Verifica la fila en Claves_Organismos.xlsx; ojo con reintentar, ARCA bloquea la clave.`
        : e.message;
    log(`ERROR ${fila.error}`);
  } finally {
    await browser.close().catch(() => {});
    fs.rmSync(tmp, { recursive: true, force: true });
  }
  return fila;
}

(async () => {
  const resultados = [];
  for (const [i, pedida] of PEDIDAS.entries()) {
    prefijo = PEDIDAS.length > 1 ? ` [${i + 1}/${PEDIDAS.length}]` : '';
    const r = await unaSociedad(pedida);
    resultados.push(r);
    // Si ARCA bloqueo la clave o pide captcha, seguir solo empeora las cosas.
    if (['captcha', 'credenciales_invalidas'].includes(r.code)) {
      log('corto el lote: ARCA pidio captcha o rechazo la clave');
      break;
    }
  }
  prefijo = '';
  const ok = resultados.every((r) => r.ok);
  console.log(JSON.stringify(PEDIDAS.length === 1 ? resultados[0] : {
    ok, periodo: PERIODO, libros: LIBROS,
    procesadas: resultados.length, pedidas: PEDIDAS.length,
    con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, error: r.error || r.libros.filter((l) => !l.ok).map((l) => `${l.libro}: ${l.error}`).join(' | ') })),
    resultados,
  }, null, 2));
  process.exit(ok ? 0 : 1);
})();
