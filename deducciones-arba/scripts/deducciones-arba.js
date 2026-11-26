#!/usr/bin/env node
/**
 * deducciones-arba.js — deducciones de IIBB que ARBA no toma en la DJ de anticipo.
 *
 * Los bancos informan percepciones sin número de certificado (tipo O, sucursal 0, número 0) y el
 * SIRTAC a veces repite el número de liquidación: ARBA toma una por agente y descarta el resto como
 * "duplicadas". Este script, por sociedad:
 *   1. entra a ARBA con la CIT de Claves_Organismos.xlsx (Playwright, como los otros scripts);
 *   2. baja el ZIP de deducciones (Deducciones → Descarga para importar, rol Contribuyente);
 *   3. busca la DJ mensual del período en "Consulta DDJJ pendientes" y baja el Archivo de deducciones
 *      (lo que ARBA tomó);
 *   4. cruza y arma los lotes de MODIFICACIÓN (operación M): al certificado que ARBA tomó de cada
 *      banco / liquidación SIRTAC le pone la suma de todas. Es lo que ARBA acepta (probado 11/2026);
 *      darlas de alta con otro número da "Deducción duplicada";
 *   5. deja todo en [Sociedad]/01-Impuestos Mensuales/AAAA/AAAAMM/ARBA AAAAMM DEDUCCIONES:
 *      el ZIP RP, el Deducciones_DJ, los lotes CUIT-AAAAMMM-P/T-LOTE1_<MD5>.zip y el Excel
 *      "ARBA AAAAMM <SOC> - percepciones no tomadas.xlsx".
 *
 * NO sube los lotes ni presenta nada: la carga por lote en la DJ la hace la persona.
 *
 *   node scripts/deducciones-arba.js --nombre="Soy 60" --periodo=10/2026
 *   node scripts/deducciones-arba.js --cuits="30715395513;30715491253" --periodo=202610
 *   node scripts/deducciones-arba.js --zip=RP-...zip --xls=Deducciones_DJ_...xls --cuit=... --periodo=202610   (sin ARBA)
 *
 * --nombre / --cuit, --nombres / --cuits   sociedad(es). --periodo MM/AAAA o AAAAMM (default: mes anterior).
 * --iniciar      si la DJ del período no está iniciada, la inicia (queda Pendiente, no se presenta).
 *                Sin este flag, la sociedad queda con code "sin_dj" y no se toca nada.
 * --no-archivar  deja todo en ~/Documents/BVA-salidas/deducciones-arba/<CUIT>/ ($BVA_SALIDAS_PATH).
 * --zip / --xls  cruce con archivos ya bajados (una sola sociedad, no entra a ARBA).
 * --ver          navegador visible.
 *
 * Un login por sociedad, sin reintentos: ARBA bloquea la CIT tras varios intentos fallidos.
 * Salida: JSON por stdout; progreso por stderr.
 */
const fs = require('fs'); const os = require('os'); const path = require('path');
const { parseArgs, fail } = require('../lib/args');
if ('ver' in parseArgs()) process.env.PLAYWRIGHT_HEADLESS = 'false';
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const { launchOptions } = require('../lib/launch');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const rutas = require('../lib/rutas-bva');
const { cruzar } = require('../lib/deducciones');

loadEnv();
const args = parseArgs();
function normalizarPeriodo(p) {
  const s = String(p || '').trim(); let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/))) return s;
  return null;
}
function mesAnterior(hoy = new Date()) { const d = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1); return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`; }
const PER = args['periodo'] ? normalizarPeriodo(args['periodo']) : mesAnterior();
if (!PER) fail('--periodo tiene que ser MM/AAAA o AAAAMM');
const ANIO = PER.slice(0, 4), MES = String(Number(PER.slice(4)));
const ARCHIVAR = !('no-archivar' in args);
const INICIAR = 'iniciar' in args;
const SALIDAS = process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas');

const URL_DFE = 'https://www.arba.gov.ar/Aplicaciones/DomicilioElectronico.asp';
const URL_DJ = 'https://app.arba.gov.ar/IBPresentaciones/welcome.do';

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [...lista(args['nombres']).map((nombre) => ({ nombre })), ...lista(args['cuits']).map((cuit) => ({ cuit }))];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length) fail('Falta --nombre, --cuit, --nombres o --cuits');
if ((args['zip'] || args['xls']) && (PEDIDAS.length > 1 || !args['zip'] || !args['xls'])) fail('--zip y --xls van juntos y con una sola sociedad');

let prefijo = '';
const log = (m) => process.stderr.write(`[deducciones-arba]${prefijo} ${m}\n`);
const num = (s) => s == null ? null : Number(String(s).replace(/[$\s.]/g, '').replace(',', '.'));

let FILAS = null;
function filaDe(pedida) {
  if (!FILAS) FILAS = claves.loadSync().rows;
  if (pedida.cuit) { const c = claves.normCuit(pedida.cuit); return FILAS.find((r) => r.cuit === c) || { cuit: c, titular: null }; }
  const { row, candidatos } = claves.buscar(FILAS, { nombre: pedida.nombre });
  if (candidatos) throw new Error(`"${pedida.nombre}" coincide con ${candidatos.length} titulares: ` + candidatos.slice(0, 8).map((r) => `${r.titular} (${r.cuit})`).join(', ') + '. Afiná el nombre o pasá --cuit.');
  return row;
}

// Nombre corto para los archivos: "3 - Soy 60 - 30715395513" -> "SOY 60"
function nombreCorto(cuit, titular) {
  try { return path.basename(rutas.carpetaSociedad(cuit)).replace(/^\s*\d+\s*-\s*/, '').replace(/\s*-\s*\d{11}\s*$/, '').trim().toUpperCase(); }
  catch { return String(titular || cuit).toUpperCase(); }
}
function carpetaDestino(cuit) {
  if (!ARCHIVAR) return path.join(SALIDAS, 'deducciones-arba', cuit, PER);
  const mes = rutas.carpetaMes(path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales'), PER);
  return path.join(mes, `ARBA ${PER} DEDUCCIONES`);
}
// Nunca pisa: si ya existe un archivo con ese nombre lo deja y lo anota.
function guardar(dir, nombre, buf, out) {
  fs.mkdirSync(dir, { recursive: true });
  const f = path.join(dir, nombre);
  if (fs.existsSync(f)) { (out.ya_existian = out.ya_existian || []).push(nombre); return f; }
  fs.writeFileSync(f, buf); (out.archivos = out.archivos || []).push(nombre); return f;
}

// ── ARBA ─────────────────────────────────────────────────────────────────────
async function login(p, usuario, clave, cuit) {
  await p.goto(URL_DFE, { waitUntil: 'domcontentloaded', timeout: 45000 });
  await p.waitForSelector('#CUIT', { timeout: 20000 });
  await p.fill('#CUIT', usuario); await p.fill('#clave_Cuit', clave);
  await Promise.all([p.waitForLoadState('domcontentloaded'), p.locator('button[type=submit]').first().click()]);
  await p.waitForTimeout(2000);
  // Algunas sociedades (las que además son agentes) pasan por "Seleccione un rol": va Contribuyente.
  if (/Seleccione un rol/i.test(await p.innerText('body').catch(() => ''))) {
    const sel = p.locator('select').first();
    const opciones = await sel.locator('option').allInnerTexts();
    const cont = opciones.find((o) => /contribuyente/i.test(o));
    if (!cont) throw Object.assign(new Error('La CIT no tiene rol Contribuyente: ' + opciones.join(' / ')), { code: 'sin_rol_contribuyente' });
    await sel.selectOption({ label: cont });
    await Promise.all([p.waitForLoadState('domcontentloaded'), p.locator('button:has-text("Continuar"), input[value="Continuar"]').first().click()]);
    await p.waitForTimeout(2000);
  }
  if (/sso\.arba/.test(p.url())) {
    const t = (await p.innerText('body')).replace(/\s+/g, ' ');
    throw Object.assign(new Error('ARBA no aceptó el ingreso (no se reintenta, ARBA bloquea la CIT): ' + t.slice(0, 160)), { code: 'arba_login_rechazado' });
  }
  const t = await p.innerText('body');
  const c = (t.match(/Cuit:\s*Usuario:\s*([\d-]+)/) || [])[1];
  if (c && c.replace(/\D/g, '') !== cuit) throw Object.assign(new Error('ARBA entró como ' + c), { code: 'otro_cuit' });
  if (!/Deducciones/.test(t)) throw Object.assign(new Error('No aparece el menú Deducciones (¿rol distinto de Contribuyente?)'), { code: 'sin_menu_deducciones' });
}

async function bajarZip(p) {
  await p.locator('a:has-text("Descarga para Importar")').first().evaluate((a) => a.click());
  await p.waitForLoadState('domcontentloaded'); await p.waitForTimeout(1200);
  const txt = p.locator('input[type=text]:visible');
  await txt.nth(0).fill(ANIO); await txt.nth(1).fill(PER.slice(4));
  await Promise.all([p.waitForLoadState('domcontentloaded'), p.locator('input[value="Consultar"], button:has-text("Consultar")').first().click()]);
  await p.waitForTimeout(2000);
  const t = await p.innerText('body');
  if (!/Total de Deducciones/i.test(t)) throw Object.assign(new Error('ARBA no mostró el total de deducciones de ' + PER), { code: 'sin_deducciones' });
  const totales = {};
  for (const m of (t.split(/Total de Deducciones/i)[1] || '').matchAll(/([A-ZÁÉÍÓÚ /]+?)\s*\$\s*([\d.,-]+)/g)) totales[m[1].trim()] = num(m[2]);
  const [dl] = await Promise.all([p.waitForEvent('download', { timeout: 60000 }), p.locator('input[value="Descargar"], button:has-text("Descargar")').first().click()]);
  return { nombre: dl.suggestedFilename(), buf: fs.readFileSync(await dl.path()), totales };
}

// Pantallas de la app de DJ. El menú queda escondido detrás del ícono de rayitas, así que se va
// directo a la URL. La primera carga después del login a veces queda vacía: se reintenta una vez.
const PANTALLA_DJ = { pendientes: ['preBuscarDDJJPendientes.do', /Buscar declaraciones juradas pendientes/i], inicio: ['preInicioDDJJ.do', /Iniciar declaraci/i] };
async function irADJ(p, cual) {
  const [url, espera] = PANTALLA_DJ[cual];
  await p.goto(URL_DJ, { waitUntil: 'domcontentloaded', timeout: 45000 }); await p.waitForTimeout(1500);
  if (!/Presentaciones de DJ/i.test(await p.innerText('body'))) throw Object.assign(new Error('No abrió Ingresos Brutos - Presentaciones de DJ'), { code: 'sin_app_dj' });
  for (let i = 0; i < 2; i++) {
    await p.goto('https://app.arba.gov.ar/IBPresentaciones/' + url, { waitUntil: 'domcontentloaded', timeout: 45000 }); await p.waitForTimeout(1500);
    if (espera.test(await p.innerText('body'))) return;
  }
  throw Object.assign(new Error('No abrió la pantalla ' + url), { code: 'sin_app_dj' });
}

async function buscarPendiente(p) {
  await irADJ(p, 'pendientes');
  await p.locator('a[href="#tabs-2"]').first().click(); await p.waitForTimeout(500);
  const sel = p.locator('#tabs-2 select');
  // [estado, año desde, mes desde, año hasta, mes hasta]
  await sel.nth(1).selectOption(ANIO); await sel.nth(2).selectOption(MES);
  await sel.nth(3).selectOption(ANIO); await sel.nth(4).selectOption(MES);
  await p.locator('#tabs-2 button:has-text("Buscar DDJJ"), #tabs-2 input[value="Buscar DDJJ"]').first().click(); await p.waitForTimeout(3000);
  // En la grilla, "Ver detalle de la DJ <nro>" es el ícono de la fila (texto oculto para lectores).
  const ver = p.locator('text=/Ver detalle de la DJ/').first();
  return (await ver.count()) ? ver : null;
}

async function iniciarDJ(p) {
  await irADJ(p, 'inicio');
  const sel = p.locator('select:visible');
  await sel.nth(0).selectOption('1'); await sel.nth(1).selectOption(ANIO); await sel.nth(2).selectOption(MES);
  await Promise.all([p.waitForLoadState('domcontentloaded'), p.locator('button:has-text("Iniciar DJ"), input[value="Iniciar DJ"]').first().click()]);
  await p.waitForTimeout(2500);
}

async function bajarDJ(p, verDetalle) {
  // El texto está oculto adentro del ícono: se clickea el elemento clickeable que lo contiene.
  await Promise.all([p.waitForLoadState('domcontentloaded'),
    verDetalle.evaluate((el) => (el.closest('a,button,[onclick]') || el).click())]);
  await p.waitForTimeout(2500);
  if (!/Detalle de DJ/i.test(await p.innerText('body'))) throw Object.assign(new Error('No abrió el detalle de la DJ'), { code: 'sin_detalle' });
  const t = (await p.innerText('body')).replace(/\s+/g, ' ');
  const g = (re) => (t.match(re) || [])[1];
  const dj = { comprobante: g(/Nro comprobante:\s*(\d+)/), periodo: (t.match(/Año - Período:\s*(\d{4})\s*-\s*(\d{1,2})/) || []).slice(1).join('-'),
    estado: g(/Nro comprobante:\s*\d+\s*\(([^)]+)\)/), ded_agentes: num(g(/declaradas por los agentes \$\s*([\d.,-]+)/)), ded_contrib: num(g(/declaradas por el contribuyente \$?\s*([\d.,-]+)/)) };
  if (dj.periodo !== `${ANIO}-${MES}`) throw Object.assign(new Error(`El detalle es de ${dj.periodo}, no de ${PER}`), { code: 'otro_periodo' });
  const ok = p.locator('button:visible:has-text("Aceptar")'); if (await ok.count()) { await ok.first().click(); await p.waitForTimeout(400); }
  await p.locator('a:visible', { hasText: /^\s*Deducciones\s*$/ }).first().click(); await p.waitForTimeout(1500);
  // "Archivo de deducciones" dispara un cartel nativo antes de la descarga: se acepta.
  p.once('dialog', (d) => d.accept().catch(() => {}));
  const [dl] = await Promise.all([p.waitForEvent('download', { timeout: 90000 }), p.locator('a:visible:has-text("Archivo de deducciones")').first().click()]);
  return { dj, nombre: dl.suggestedFilename(), buf: fs.readFileSync(await dl.path()) };
}

// ── Excel de no tomadas ──────────────────────────────────────────────────────
async function excel(r, soc) {
  const wb = new ExcelJS.Workbook();
  const hoja = (nombre, cols, filas) => {
    const ws = wb.addWorksheet(nombre); ws.columns = cols.map(([header, width, fmt]) => ({ header, width, style: fmt ? { numFmt: fmt } : {} }));
    ws.getRow(1).font = { bold: true }; filas.forEach((f) => ws.addRow(f)); return ws;
  };
  const M = '#,##0.00';
  const { cp, ct } = r._detalle;
  const nt = cp.filter((d) => !d.tomada);
  const ws1 = hoja('No tomadas', [['Línea TXT', 12], ['CUIT Agente', 15], ['Fecha', 12], ['Tipo', 6], ['Suc', 6], ['Nro', 10], ['Importe', 14, M], ['Línea original', 80]],
    nt.map((d) => [d.linea, d.cuit, d.fecha, d.tipo, d.suc, d.nro, d.imp, d.raw]));
  ws1.addRow([]); ws1.addRow(['TOTAL', '', '', '', '', '', r.percepciones.importe_no_tomado]).font = { bold: true };
  const banc = cp.filter((d) => d.tipo === 'O' && d.suc === 0 && d.nro === 0)
    .sort((a, b) => a.cuit.localeCompare(b.cuit) || a.fecha.slice(3, 5).localeCompare(b.fecha.slice(3, 5)) || a.fecha.localeCompare(b.fecha));
  hoja('Bancarias sin nro (todas)', [['CUIT Agente', 15], ['Fecha', 12], ['Importe', 14, M], ['Estado ARBA', 16]], banc.map((d) => [d.cuit, d.fecha, d.imp, d.tomada ? 'TOMADA' : 'NO TOMADA']));
  if (r.sirtac.liquidaciones.length) {
    const filas = [];
    for (const l of r.sirtac.liquidaciones) for (const x of l.lineas) filas.push([l.cuit, x.fecha, l.liquidacion, x.imp, x.tomada ? 'TOMADA' : 'NO TOMADA (mismo nro de liquidación)']);
    hoja('SIRTAC no tomada', [['CUIT Agente', 15], ['Fecha', 12], ['Nro liquidación', 24], ['Importe', 14, M], ['Estado ARBA', 40]], filas);
  }
  const rs = hoja('Resumen lotes', [['Lote', 44], ['Línea', 75]], r.lotes.flatMap((l) => l.lineas.map((x) => [l.zipNombre, x])));
  rs.addRow([]); rs.addRow([`${soc} ${PER} · faltante total $ ${r.faltante.toFixed(2)} · subir cada ZIP por Carga por lote de la DJ`]);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ── Una sociedad ─────────────────────────────────────────────────────────────
async function procesar(out, zip, xls) {
  const r = cruzar({ zip: zip.buf, xls: xls.buf, cuit: out.cuit, periodo: PER });
  out.secciones = r.secciones; out.percepciones = r.percepciones; out.sirtac = r.sirtac;
  out.faltante = r.faltante; out.otras_diferencias = r.otras_diferencias;
  out.lotes = r.lotes.map((l) => ({ zip: l.zipNombre, lineas: l.lineas }));
  if (out.dj && out.dj.ded_agentes != null) {
    const totZip = Object.values(r.secciones).reduce((s, x) => s + x.total_txt, 0);
    out.control = { total_zip: Math.round(totZip * 100) / 100, ded_agentes_dj: out.dj.ded_agentes, diferencia: Math.round((totZip - out.dj.ded_agentes) * 100) / 100,
      cierra: Math.abs(totZip - out.dj.ded_agentes - r.faltante) < 0.01 };
  }
  if (!r.faltante && !r.lotes.length) { out.ok = true; out.nada_que_hacer = true; log('ARBA tomó todo, no hace falta lote'); return; }
  const soc = nombreCorto(out.cuit, out.sociedad);
  const dir = carpetaDestino(out.cuit); out.carpeta = dir;
  guardar(dir, zip.nombre, zip.buf, out);
  guardar(dir, xls.nombre, xls.buf, out);
  for (const l of r.lotes) { guardar(dir, l.base + '.txt', l.txt, out); guardar(dir, l.zipNombre, l.zip, out); }
  guardar(dir, `ARBA ${PER} ${soc} - percepciones no tomadas.xlsx`, await excel(r, soc), out);
  out.ok = out.otras_diferencias.length === 0 && (!out.control || out.control.cierra)
    && !r.percepciones.bancos.some((b) => b.error) && !r.sirtac.liquidaciones.some((l) => l.error);
  log(`faltan $ ${r.faltante} · lotes: ${r.lotes.map((l) => l.tipo).join(', ') || '-'} · ${dir}`);
}

async function unaSociedad(pedida) {
  let fila;
  try { fila = args['zip'] && pedida.cuit ? { cuit: claves.normCuit(pedida.cuit), titular: null } : filaDe(pedida); } catch (e) { return { sociedad: pedida.nombre || pedida.cuit, ok: false, code: 'sociedad', error: e.message }; }
  const cuit = fila.cuit;
  const out = { cuit, sociedad: fila.titular || pedida.nombre || cuit, periodo: PER, ok: false };
  try {
    if (args['zip']) {
      const f = (x) => ({ nombre: path.basename(x), buf: fs.readFileSync(x) });
      await procesar(out, f(args['zip']), f(args['xls'])); return out;
    }
    const a = fila.titular ? claves.credenciales(fila, 'ARBA') : null;
    if (!a || !a.clave) return { ...out, code: 'sin_cit', error: 'Sin CIT de ARBA en Claves_Organismos.xlsx.' };
    log(`${out.sociedad} · CUIT ${cuit} · ${PER}`);
    const b = await chromium.launch(launchOptions());
    try {
      const ctx = await b.newContext({ locale: 'es-AR', viewport: { width: 1366, height: 900 }, acceptDownloads: true });
      const p = await ctx.newPage();
      await login(p, claves.normCuit(a.usuario) || cuit, a.clave, cuit);
      const zip = await bajarZip(p); out.totales_arba = zip.totales; log('ZIP ' + zip.nombre);
      let ver = await buscarPendiente(p);
      if (!ver && INICIAR) { log('DJ no iniciada: la inicio'); await iniciarDJ(p); out.dj_iniciada = true; ver = await buscarPendiente(p); }
      if (!ver) {
        const dir = path.join(SALIDAS, 'deducciones-arba', cuit, PER); fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(path.join(dir, zip.nombre), zip.buf);
        return { ...out, code: 'sin_dj', error: `La DJ ${PER} no está iniciada (o ya se presentó). Iniciala y volvé a correr, o usá --iniciar. ZIP guardado en ${dir}` };
      }
      const xls = await bajarDJ(p, ver); out.dj = xls.dj; log(`DJ ${xls.dj.comprobante} (${xls.dj.estado}) · agentes $ ${xls.dj.ded_agentes}`);
      await procesar(out, zip, xls);
    } finally { await b.close().catch(() => {}); }
  } catch (e) { out.code = e.code || out.code || 'error'; out.error = e.message.split('\n')[0]; log('ERROR ' + out.error); }
  return out;
}

(async () => {
  const resultados = [];
  for (const [i, pedida] of PEDIDAS.entries()) { prefijo = PEDIDAS.length > 1 ? ` [${i + 1}/${PEDIDAS.length}]` : ''; resultados.push(await unaSociedad(pedida)); }
  prefijo = '';
  const ok = resultados.every((r) => r.ok);
  console.log(JSON.stringify(PEDIDAS.length === 1 ? resultados[0] : { ok, periodo: PER, procesadas: resultados.length,
    con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, code: r.code, error: r.error })), resultados }, null, 1));
  process.exit(ok ? 0 : 1);
})();
