#!/usr/bin/env node
/**
 * constancias-anticipo-arba.js — baja de ARBA (IBPresentaciones > Reimpresion > Constancia anticipo)
 * las constancias de las DDJJ de anticipo de Ingresos Brutos YA PRESENTADAS de una o varias sociedades
 * para un rango de periodos: el PDF original de resumenPDF.do + los datos del detalle (bases, deducciones,
 * impuesto determinado). Las archiva en la carpeta IIBB del mes como AAAAMM-IIBB-DJ-<TOKEN>.pdf.
 *
 *   node scripts/constancias-anticipo-arba.js --nombre="<sociedad>" --desde=202601 --hasta=202608
 *   node scripts/constancias-anticipo-arba.js --cuit=<CUIT> --periodo=09/2026
 *   node scripts/constancias-anticipo-arba.js --nombres="<A>;<B>" --periodo=09/2026 --no-archivar
 *
 * --nombre / --cuit, --nombres / --cuits   sociedad(es), como en los otros scripts.
 * --desde / --hasta  AAAAMM. --periodo (MM/AAAA o AAAAMM) es un solo mes. Default: el mes anterior.
 * --no-archivar   deja los PDF solo en ~/Documents/BVA-salidas/arba-ddjj/<CUIT>/ ($BVA_SALIDAS_PATH).
 * --token         <TOKEN> del nombre archivado. Default: el que ya usan los archivos AAAAMM-IIBB-DJ-<TOKEN>
 *                 (o AAAAMM-IVA-DDJJ-<TOKEN>) de esa sociedad; si no hay ninguno, hay que pasarlo.
 * --ver           navegador visible (solo si se pide).
 *
 * Un login por sociedad con su CIT (Claves_Organismos.xlsx, columnas ARBA), sin reintentos: ARBA
 * bloquea la CIT tras varios intentos fallidos. Si un periodo tiene rectificativa se toma la DJ
 * presentada mas reciente. Nunca pisa: si en la carpeta del mes ya hay un AAAAMM-IIBB-DJ-* lo deja.
 * Solo lectura en ARBA. Salida: JSON por stdout; progreso por stderr.
 */
const fs = require('fs'); const os = require('os'); const path = require('path');
const { parseArgs, fail } = require('../lib/args');
if ('ver' in parseArgs()) process.env.PLAYWRIGHT_HEADLESS = 'false';
const { chromium } = require('playwright');
const { launchOptions } = require('../lib/launch');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const rutas = require('../lib/rutas-bva');
const { tokenSociedad } = require('../lib/token-sociedad');

loadEnv();
const args = parseArgs();
function normalizarPeriodo(p) {
  const s = String(p || '').trim(); let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/))) return s;
  return null;
}
function mesAnterior(hoy = new Date()) {
  const d = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
}
const UNO = args['periodo'] ? normalizarPeriodo(args['periodo']) : null;
if (args['periodo'] && !UNO) fail('--periodo tiene que ser MM/AAAA o AAAAMM');
const DESDE = UNO || String(args['desde'] || args['hasta'] || mesAnterior());
const HASTA = UNO || String(args['hasta'] || args['desde'] || mesAnterior());
if (!/^\d{6}$/.test(DESDE) || !/^\d{6}$/.test(HASTA) || DESDE > HASTA) fail('--desde y --hasta en AAAAMM, desde <= hasta');
const ARCHIVAR = !('no-archivar' in args);
const SALIDAS = process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas');
const URL_CONSTANCIA = 'https://app.arba.gov.ar/IBPresentaciones/preBuscarDDJJCerradasParaConstancia.do';

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length) fail('Falta --nombre, --cuit, --nombres o --cuits');
if (args['token'] && PEDIDAS.length > 1) fail('--token va con una sola sociedad');

let prefijo = '';
const log = (m) => process.stderr.write(`[arba-ddjj]${prefijo} ${m}\n`);
const num = (s) => s == null ? null : Number(String(s).replace(/[$\s.]/g, '').replace(',', '.'));

function periodos() {
  const out = []; let y = +DESDE.slice(0, 4), m = +DESDE.slice(4);
  while (`${y}${String(m).padStart(2, '0')}` <= HASTA) { out.push(`${y}${String(m).padStart(2, '0')}`); if (++m > 12) { m = 1; y++; } }
  return out;
}

let FILAS = null;
function filaDe(pedida) {
  if (!FILAS) FILAS = claves.loadSync().rows;
  if (pedida.cuit) {
    const c = claves.normCuit(pedida.cuit);
    return FILAS.find((r) => r.cuit === c) || { cuit: c, titular: null };
  }
  const { row, candidatos } = claves.buscar(FILAS, { nombre: pedida.nombre });
  if (candidatos) {
    throw new Error(`"${pedida.nombre}" coincide con ${candidatos.length} titulares: ` +
      candidatos.slice(0, 8).map((r) => `${r.titular} (${r.cuit})`).join(', ') + '. Afina el nombre o pasa --cuit.');
  }
  return row;
}

// Archiva en [mes]/IIBB/AAAAMM-IIBB-DJ-<TOKEN>.pdf si no hay ya una DJ del mes (suelta o en IIBB/).
function archivar(cuit, per, pdf) {
  const mes = rutas.carpetaMes(path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales'), per);
  const iibb = path.join(mes, 'IIBB');
  const re = new RegExp(`^${per}\\s*-\\s*IIBB\\s*-\\s*DJ\\s*-`, 'i');
  const ya = [mes, iibb].filter((x) => fs.existsSync(x)).flatMap((x) => fs.readdirSync(x)).find((f) => re.test(f));
  if (ya) return { archivo: 'ya existía: ' + ya };
  const tok = args['token'] ? String(args['token']) : tokenSociedad(cuit, 'iibb-dj');
  if (!tok) return { archivo_error: 'sin_token: la sociedad no tiene archivos AAAAMM-IIBB-DJ-<TOKEN> para copiar el nombre; pasa --token (preguntale al equipo cual usa).' };
  fs.mkdirSync(iibb, { recursive: true });
  const dest = path.join(iibb, `${per}-IIBB-DJ-${tok}.pdf`);
  fs.copyFileSync(pdf, dest);
  return { archivo: path.relative(rutas.raiz(), dest) };
}

async function unaSociedad(pedida) {
  let fila;
  try { fila = filaDe(pedida); } catch (e) { return { sociedad: pedida.nombre || pedida.cuit, ok: false, code: 'sociedad', error: e.message }; }
  const cuit = fila.cuit;
  const out = { cuit, sociedad: fila.titular || pedida.nombre || cuit, ok: false, ddjj: [] };
  const a = fila.titular ? claves.credenciales(fila, 'ARBA') : null;
  if (!a || !a.clave) return { ...out, code: 'sin_cit', error: 'Sin CIT de ARBA en Claves_Organismos.xlsx (si presenta por Convenio Multilateral o AGIP, esto no aplica).' };
  log(`${out.sociedad} · CUIT ${cuit} · ${DESDE}-${HASTA}`);
  const b = await chromium.launch(launchOptions());
  const ctx = await b.newContext({ locale: 'es-AR', viewport: { width: 1366, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage();
  try {
    await p.goto(URL_CONSTANCIA, { waitUntil: 'domcontentloaded', timeout: 45000 });
    await p.waitForSelector('#CUIT', { timeout: 20000 });
    await p.fill('#CUIT', claves.normCuit(a.usuario) || cuit); await p.fill('#clave_Cuit', a.clave);
    await Promise.all([p.waitForURL((u) => !/sso\.arba/.test(String(u)), { timeout: 30000 }).catch(() => {}), p.locator('button[type=submit]').first().click()]);
    if (/sso\.arba/.test(p.url())) {
      const t = (await p.innerText('body')).replace(/\s+/g, ' ');
      return { ...out, code: 'arba_login_rechazado', error: 'ARBA no aceptó el ingreso (no se reintenta, ARBA bloquea la CIT): ' + t.slice(0, 160) };
    }
    await p.waitForTimeout(1500);
    const usuario = ((await p.innerText('body')).match(/Usuario:\s*([\d-]+)\s+(.+?)\s+\d{2}\/\d{2}\/\d{4}/) || []);
    out.usuario = usuario[1]; out.razon = usuario[2];
    if (usuario[1] && usuario[1].replace(/\D/g, '') !== cuit) return { ...out, code: 'otro_cuit', error: `ARBA entró como ${usuario[1]} ${usuario[2]}` };
    // La primera entrada a la constancia tiene que ser por el menu: directo a la URL queda la pantalla vacia.
    await p.locator('a:has-text("Reimpresión")').first().click(); await p.waitForTimeout(400);
    await Promise.all([p.waitForLoadState('domcontentloaded'), p.locator('a:has-text("Constancia anticipo")').first().click()]); await p.waitForTimeout(2000);
    const set = async (id, v) => p.evaluate(([id, v]) => { const e = document.getElementById(id); e.value = v; e.dispatchEvent(new Event('change', { bubbles: true })); }, [id, v]);
    // Busqueda "por criterios varios"; la tabla se lee entera con la API de DataTables (pagina de a 5).
    const buscar = async (d, h) => {
      await p.goto(URL_CONSTANCIA, { waitUntil: 'domcontentloaded' }); await p.waitForTimeout(800);
      await p.click('#ui-id-2'); await p.waitForTimeout(400);
      await set('anioMensualDesde', d.slice(0, 4)); await set('mesDesde', String(Number(d.slice(4))));
      await set('anioMensualHasta', h.slice(0, 4)); await set('mesHasta', String(Number(h.slice(4))));
      await p.click('#buscar'); await p.waitForTimeout(2500);
      return p.evaluate(() => window.jQuery('#resultados').dataTable().fnGetData().map((r) => r.slice(0, 7).map((x) => String(x).replace(/<[^>]+>/g, '').trim())));
    };
    const per = periodos();
    const todas = await buscar(per[0], per[per.length - 1]);
    out.listado = todas.map((r) => r.join('|'));
    const dir = path.join(SALIDAS, 'arba-ddjj', cuit); fs.mkdirSync(dir, { recursive: true });
    for (const pp of per) {
      const delMes = todas.filter((r) => r[3] === pp.slice(0, 4) && Number(r[4]) === Number(pp.slice(4)));
      if (!delMes.length) { out.ddjj.push({ periodo: pp, falta: true }); log(`${pp}: sin DJ presentada`); continue; }
      const filasMes = await buscar(pp, pp);
      // la vigente: la de presentacion mas reciente (rectificativa si la hay)
      const fecha = (x) => { const m = String(x[6] || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + m[2] + m[1] : ''; };
      let idx = 0; filasMes.forEach((x, i) => { if (fecha(x) > fecha(filasMes[idx]) || (fecha(x) === fecha(filasMes[idx]) && /Rectif/i.test(x[5]))) idx = i; });
      await Promise.all([p.waitForLoadState('domcontentloaded'), p.locator('#resultados tbody tr').nth(idx).locator('span.iconoEditar').click()]); await p.waitForTimeout(1800);
      const t = (await p.innerText('body')).replace(/\s+/g, ' ');
      const g = (re) => (t.match(re) || [])[1];
      const m = t.match(/Año - Período:\s*(\d{4})\s*-\s*(\d{1,2})/); const perDet = m ? m[1] + String(m[2]).padStart(2, '0') : null;
      const d = { periodo: pp, periodo_detalle: perDet, comprobante: g(/Nro comprobante:\s*(\d+)/), tipo: g(/Tipo de DJ:\s*(\S+(?: \d+)?)/), secuencias: filasMes.length,
        gravados: num(g(/Gravados \$\s*([\d.,]+)/)), no_gravados: num(g(/No gravados \$\s*([\d.,]+)/)), tasa0: num(g(/Tasa 0 % \$\s*([\d.,]+)/)),
        ded_agentes: num(g(/declaradas por los agentes \$\s*([\d.,-]+)/)), ded_contrib: num(g(/declaradas por el contribuyente \$\s*([\d.,-]+)/)),
        impuesto: num(g(/Impuesto determinado de período \$\s*([\d.,-]+)/)), monto_imponible: num(g(/Monto imponible declarado total \$\s*([\d.,-]+)/)) };
      if (perDet !== pp) { d.error = `el detalle dice ${perDet}`; out.ddjj.push(d); log(`${pp}: ERROR ${d.error}`); continue; }
      // resumenPDF.do no recibe parametros: devuelve la DJ que quedo abierta en la sesion.
      const pdf = await p.evaluate(async () => { const r = await fetch('/IBPresentaciones/resumenPDF.do', { credentials: 'include' }); return Array.from(new Uint8Array(await r.arrayBuffer())); });
      const buf = Buffer.from(pdf);
      if (buf.slice(0, 5).toString() !== '%PDF-') { d.error = 'resumenPDF no devolvió un PDF'; out.ddjj.push(d); log(`${pp}: ERROR ${d.error}`); continue; }
      d.pdf = path.join(dir, `${pp}.pdf`); fs.writeFileSync(d.pdf, buf); d.bytes = buf.length;
      fs.writeFileSync(path.join(dir, `${pp}.json`), JSON.stringify(d, null, 1));
      if (ARCHIVAR) { try { Object.assign(d, archivar(cuit, pp, d.pdf)); } catch (e) { d.archivo_error = e.message; } }
      out.ddjj.push(d); log(`${pp}: ${d.tipo || ''} impuesto ${d.impuesto} ${d.archivo || d.archivo_error || ''}`);
    }
    out.ok = out.ddjj.every((x) => x.falta || (!x.error && x.impuesto != null));
  } catch (e) { out.code = out.code || 'error'; out.error = e.message.split('\n')[0]; log('ERROR ' + out.error); }
  finally { await b.close().catch(() => {}); }
  return out;
}

(async () => {
  const resultados = [];
  for (const [i, pedida] of PEDIDAS.entries()) {
    prefijo = PEDIDAS.length > 1 ? ` [${i + 1}/${PEDIDAS.length}]` : '';
    resultados.push(await unaSociedad(pedida));
  }
  prefijo = '';
  const ok = resultados.every((r) => r.ok);
  console.log(JSON.stringify(PEDIDAS.length === 1 ? resultados[0] : {
    ok, desde: DESDE, hasta: HASTA, procesadas: resultados.length,
    con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, code: r.code, error: r.error })),
    resultados,
  }, null, 1));
  process.exit(ok ? 0 : 1);
})();
