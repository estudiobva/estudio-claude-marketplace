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
 *   --aunque-presentada Entra al borrador aunque haya acuse de la DDJJ IVA del periodo.
 *   --presentada        Baja los libros de la DDJJ YA PRESENTADA (ultima secuencia),
 *                       en solo lectura: no importa ni toca el borrador. Ver abajo.
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
 * Con --presentada: Declaraciones juradas presentadas -> Libro IVA -> "Ver" de
 * la ultima secuencia del periodo -> Libro Compras / Ventas (sin IMPORTAR) -> CSV.
 * Antes de archivar controla que el neto gravado del CSV (operaciones y notas de
 * credito) coincida con la vista previa de la DDJJ; si no coincide, no archiva.
 * No mira el acuse de la carpeta: el que manda es ARCA (si no hay presentacion
 * del periodo, corta con sin_presentacion). Nunca clickea "Rectificar".
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
const PRESENTADA = 'presentada' in args;
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

// Neto gravado e IVA del CSV, separando operaciones (positivos) de notas de
// credito (negativas), para compararlos con la vista previa de la DDJJ.
// Una linea del CSV de ARCA: separador ";", textos entre comillas (pueden traer ";").
function camposCsv(linea) {
  const out = []; let cur = '', q = false;
  for (let i = 0; i < linea.length; i++) {
    const ch = linea[i];
    if (q) { if (ch === '"' && linea[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === ';') { out.push(cur); cur = ''; }
    else cur += ch;
  }
  out.push(cur);
  return out.map((c) => c.trim());
}

function totalesCsv(csv) {
  const lineas = fs.readFileSync(csv, 'latin1').split(/\r?\n/).filter((l) => l.trim());
  const cab = camposCsv(lineas[0]);
  const col = (n) => { const i = cab.indexOf(n); if (i < 0) throw new Error(`El CSV no tiene la columna "${n}".`); return i; };
  const num = (v) => Number(String(v || '0').trim().replace(/\./g, '').replace(',', '.')) || 0;
  const cNeto = col('Total Neto Gravado'), cIva = col('Total IVA');
  const t = { operaciones: { neto: 0, iva: 0 }, notasCredito: { neto: 0, iva: 0 } };
  for (const l of lineas.slice(1)) {
    const c = camposCsv(l);
    const neto = num(c[cNeto]), iva = num(c[cIva]);
    const k = neto < 0 || (neto === 0 && iva < 0) ? 'notasCredito' : 'operaciones';
    t[k].neto += Math.abs(neto); t[k].iva += Math.abs(iva);
  }
  for (const k of Object.keys(t)) for (const m of ['neto', 'iva']) t[k][m] = Math.round(t[k][m] * 100) / 100;
  return t;
}

// El neto gravado tiene que dar al centavo. El IVA se informa pero no se exige:
// ARCA redondea por alicuota y en 08/2026 difirio en unos pocos pesos.
function controlarContraVistaPrevia(libro, csvTot, vista) {
  const v = vista && vista[libro];
  if (!v || !v.operaciones) throw new Error(`No pude leer los totales de ${libro} en la vista previa de la DDJJ: no archivo sin controlar.`);
  const esperado = { operaciones: v.operaciones.neto, notasCredito: v.notasCredito ? v.notasCredito.neto : 0 };
  const dif = ['operaciones', 'notasCredito'].map((k) => ({ k, csv: csvTot[k].neto, ddjj: esperado[k], d: Math.round((csvTot[k].neto - esperado[k]) * 100) / 100 }));
  const mal = dif.filter((x) => Math.abs(x.d) > 0.05);
  if (mal.length) {
    throw new Error(`El CSV de ${libro} no coincide con la DDJJ presentada: ` +
      mal.map((x) => `${x.k} neto CSV ${x.csv} vs DDJJ ${x.ddjj}`).join('; ') + '. No lo archivo.');
  }
  return {
    neto_operaciones: esperado.operaciones, neto_notas_credito: esperado.notasCredito,
    iva_csv: Math.round((csvTot.operaciones.iva - csvTot.notasCredito.iva) * 100) / 100,
    iva_ddjj: Math.round((v.operaciones.iva - (v.notasCredito ? v.notasCredito.iva : 0)) * 100) / 100,
  };
}

// Un libro, con el borrador (o la DDJJ presentada) ya abierto.
async function bajarLibro(iva, libro, destinoDir, tmp, presentada = null) {
  const nombre = rutas.nombreArchivo(PERIODO, libro, 'xlsx');
  const destino = ARCHIVAR ? path.join(destinoDir, nombre) : null;
  if (destino && fs.existsSync(destino) && SI_EXISTE !== 'reemplazar') {
    if (SI_EXISTE === 'saltear') return { libro, ok: true, salteado: true, archivo: destino };
    throw new Error(`Ya existe ${destino}. No lo piso: --si-existe=reemplazar o --si-existe=saltear.`);
  }

  let importacion = null;
  if (presentada) {
    await pi.irAlMenuPresentacion(iva);
    await pi.abrirLibro(iva, libro);
    await iva.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
    // Red de seguridad: en una DDJJ presentada no tiene que haber IMPORTAR.
    if (await iva.locator('a, button').filter({ hasText: /^\s*IMPORTAR\s*$/i }).first().isVisible().catch(() => false)) {
      throw new Error(`El Libro ${libro} muestra IMPORTAR: no parece la DDJJ presentada. No sigo.`);
    }
  } else {
    await volverAlMenu(iva);
    await pi.abrirLibro(iva, libro);
    importacion = await pi.importarDesdeArca(iva);
  }

  const dirLibro = fs.mkdtempSync(path.join(tmp, `${libro}-`));
  const bajado = await pi.descargarCsv(iva, dirLibro);
  const { csv, filas, fuera } = leerCsv(bajado, dirLibro);
  log(`${libro}: CSV con ${filas} filas de datos, ${fuera} con fecha fuera del periodo`);
  if (filas === 0) throw new Error(`El CSV de ${libro} bajo vacio: no archivo nada.`);
  let control = null;
  if (presentada) {
    control = controlarContraVistaPrevia(libro, totalesCsv(csv), presentada.totales);
    log(`${libro}: coincide con la DDJJ presentada (neto operaciones ${control.neto_operaciones}, NC ${control.neto_notas_credito})`);
  }

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

  return { libro, ok: true, importacion, control_ddjj: control, filas, filasFueraDePeriodo: fuera, archivo, aviso };
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

  const acuse = PRESENTADA ? null : acuseIva(cuit);
  if (acuse && !AUNQUE_PRESENTADA) {
    return { ...fila, error: `La DDJJ de IVA ${PERIODO} ya esta presentada (${acuse}): no entro al borrador, podria abrir una rectificativa. Para bajar los libros presentados: --presentada.`, code: 'ddjj_presentada' };
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
    log(`${fila.sociedad} · CUIT ${cuit} · periodo ${PERIODO} · ${pendientes.join(' + ')}${PRESENTADA ? ' · DDJJ presentada' : ''}`);
    log('entrando a ARCA…');
    await login(page, { cuitLogin, password });
    const iva = await pi.abrirPortalIva(context, page);
    fila.representando = (await pi.asegurarRepresentacion(iva, cuit)).linea;
    let hoja = iva, presentada = null;
    if (PRESENTADA) {
      presentada = await pi.abrirPresentada(context, iva, PERIODO);
      hoja = presentada.vista;
      fila.ddjj = { formulario: presentada.elegida.formulario, secuencia: presentada.elegida.secuencia,
        presentada_el: presentada.elegida.presentada, secuencias_del_periodo: presentada.secuencias.length };
    } else {
      await pi.abrirBorrador(iva, PERIODO);
      fila.datosInicialesConfigurados = await pi.datosIniciales(iva);
    }

    for (const libro of pendientes) {
      try {
        fila.libros.push(await bajarLibro(hoja, libro, destinoDir, tmp, presentada));
      } catch (e) {
        // Un libro que falla no frena al otro, salvo que sea un corte de ARCA.
        if (['captcha', 'credenciales_invalidas', 'otro_borrador_abierto', 'sin_presentacion'].includes(e.code)) throw e;
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
    ok, periodo: PERIODO, libros: LIBROS, modo: PRESENTADA ? 'presentada' : 'borrador',
    procesadas: resultados.length, pedidas: PEDIDAS.length,
    con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, error: r.error || r.libros.filter((l) => !l.ok).map((l) => `${l.libro}: ${l.error}`).join(' | ') })),
    resultados,
  }, null, 2));
  process.exit(ok ? 0 : 1);
})();
