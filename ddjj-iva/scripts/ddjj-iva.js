#!/usr/bin/env node
/**
 * ddjj-iva.js — lee la DDJJ de IVA (F.2051) YA PRESENTADA de una o varias sociedades para un rango de
 * periodos y guarda la vista previa en PDF + los importes de la determinacion (debito, credito, saldos
 * anteriores, retenciones).
 * Solo lectura: Portal IVA -> Declaraciones juradas presentadas -> "Declaración Jurada de IVA" CONSULTAR ->
 * "Ver declaración jurada" de la ultima secuencia del periodo -> CONTINUAR -> CONTINUAR (vista previa).
 * Nunca clickea Rectificar ni Generar VEP.
 *
 *   node scripts/ddjj-iva.js --nombre="<sociedad>" --desde=202601 --hasta=202608
 *   node scripts/ddjj-iva.js --cuit=<CUIT> --periodo=09/2026 --archivar
 *   node scripts/ddjj-iva.js --nombres="<A>;<B>" --desde=202601 --hasta=202608
 *
 * --nombre / --cuit, --nombres / --cuits   sociedad(es), como en los otros scripts.
 * --desde / --hasta  AAAAMM. --periodo (MM/AAAA o AAAAMM) es un solo mes. Default: el mes anterior.
 * --salida    carpeta local para los PDF y JSON (default ~/Documents/BVA-salidas/ddjj-iva/<CUIT>)
 * --archivar  copia el PDF a la carpeta IVA del mes en la unidad como AAAAMM-IVA-DDJJ-<TOKEN>.pdf,
 *             solo si en la carpeta del mes no hay ya un *IVA*DDJJ* (no pisa lo del equipo)
 * --token     <TOKEN> del nombre archivado. Default: el que ya usan los archivos AAAAMM-IVA-DDJJ-<TOKEN>
 *             (o AAAAMM-IIBB-DJ-<TOKEN>) de esa sociedad; si no hay ninguno, hay que pasarlo. Solo con una sociedad.
 * --ver       navegador visible (solo si se pide).
 * Salida: JSON por stdout ({ ok, cuit, ddjj: [{ periodo, secuencia, df, cf, st_ant, st, sld_ant, ret, sld, pdf, archivo }] };
 * con varias sociedades { ok, procesadas, con_error, resultados }). Progreso por stderr.
 */
const fs = require('fs'); const os = require('os'); const path = require('path');
const { parseArgs, fail } = require('../lib/args');
if ('ver' in parseArgs()) process.env.PLAYWRIGHT_HEADLESS = 'false';
const { chromium } = require('playwright');
const { launchOptions } = require('../lib/launch');
const { resolveCredentials, envFile } = require('../lib/env'); const { login } = require('../lib/arca-login');
const rutas = require('../lib/rutas-bva'); const pi = require('../lib/portal-iva');
const { tokenSociedad } = require('../lib/token-sociedad');

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
const SALIDAS = process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas');
const ARCHIVAR = 'archivar' in args;

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length) fail('Falta --nombre, --cuit, --nombres o --cuits');
if (args['token'] && PEDIDAS.length > 1) fail('--token va con una sola sociedad');
if (args['salida'] && PEDIDAS.length > 1) fail('--salida va con una sola sociedad');

let prefijo = '';
const log = (m) => process.stderr.write(`[ddjj-iva]${prefijo} ${m}\n`);
const num = (s) => s == null ? null : Number(String(s).replace(/[$\s.]/g, '').replace(',', '.'));

function periodos() {
  const out = []; let y = +DESDE.slice(0, 4), m = +DESDE.slice(4);
  while (`${y}${String(m).padStart(2, '0')}` <= HASTA) { out.push(`${y}${String(m).padStart(2, '0')}`); if (++m > 12) { m = 1; y++; } }
  return out;
}

function tokenDe(cuit) {
  if (args['token']) return String(args['token']);
  const t = tokenSociedad(cuit, 'iva-ddjj');
  if (!t) throw Object.assign(new Error('La sociedad no tiene archivos AAAAMM-IVA-DDJJ-<TOKEN> ni AAAAMM-IIBB-DJ-<TOKEN> para copiar el token del nombre: pasa --token (preguntale al equipo cual usa).'), { code: 'sin_token' });
  return t;
}

function leerVista(t) {
  const g = (re) => { const m = t.match(re); return m ? num(m[1]) : null; };
  const n = '\\s*\\$\\s*(-?[\\d.]+,\\d{2})';
  return {
    periodo_vista: (t.match(/Per[ií]odo\s+(\d{6})/) || [])[1] || null,
    secuencia: (t.match(/Secuencia\s+([^\n]+)/) || [])[1] || null,
    transaccion: (t.match(/Nro\. de Transacci[oó]n\s+(\d+)/) || [])[1] || null,
    presentada_el: (t.match(/Fecha de Presentaci[oó]n\s+([\d/]+)/) || [])[1] || null,
    df: g(new RegExp('Total del d[ée]bito fiscal del per[ií]odo' + n)),
    cf: g(new RegExp('Total del cr[ée]dito fiscal del per[ií]odo' + n)),
    st_ant: g(new RegExp('Saldo t[ée]cnico a favor del contribuyente del per[ií]odo anterior' + n)),
    st_arca: g(new RegExp('Saldo t[ée]cnico a favor de ARCA' + n)),
    st: g(new RegExp('Saldo t[ée]cnico a favor del contribuyente(?! del per)' + n)),
    sld_ant: g(new RegExp('libre disponibilidad del per[ií]odo anterior neto de usos' + n)),
    ret: g(new RegExp('Total de retenciones, percepciones y pagos a cuenta[^$]*' + n)),
    sld: g(new RegExp('Saldo de libre disponibilidad a favor del contribuyente del per[ií]odo' + n)),
    saldo_arca: g(new RegExp('Saldo de impuesto a favor de ARCA' + n)),
  };
}

async function abrirListaDdjj(iva, base) {
  // directo a la lista (#/declaraciones.anteriores); si no carga, por las tarjetas
  await iva.goto(base + '#/declaraciones.anteriores', { waitUntil: 'domcontentloaded' }).catch(() => {});
  if (await iva.locator('button[title="Ver declaración jurada"]').first().waitFor({ state: 'visible', timeout: 20000 }).then(() => true).catch(() => false)) {
    await iva.waitForTimeout(1000); return;
  }
  await iva.goto(base + '#/anteriores', { waitUntil: 'domcontentloaded' }).catch(() => {});
  await iva.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
  await iva.waitForSelector('button:has-text("CONSULTAR"), a:has-text("CONSULTAR")', { timeout: 30000 });
  await iva.waitForTimeout(1000);
  // dos tarjetas con descripcion parecida: la de la DDJJ es la que se titula "Declaración Jurada de IVA"
  const card = iva.locator('a.panel, .panel, div').filter({ hasText: /Declaraci[oó]n Jurada de IVA/ }).filter({ has: iva.locator('button:has-text("CONSULTAR"), a:has-text("CONSULTAR")') }).last();
  const btn = card.locator('button:has-text("CONSULTAR"), a:has-text("CONSULTAR")').first();
  if (await btn.count()) await btn.click(); else await iva.locator('button:has-text("CONSULTAR"), a:has-text("CONSULTAR")').nth(1).click();
  await iva.waitForURL(/declaraciones\.anteriores/, { timeout: 30000 });
  await iva.waitForSelector('button[title="Ver declaración jurada"]', { timeout: 30000 });
  await iva.waitForTimeout(1500);
}

function archivar(cuit, per, pdf) {
  const mes = rutas.carpetaMes(path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales'), per);
  const re = new RegExp(`^${per}.*IVA.*DDJJ`, 'i');
  const dirs = fs.existsSync(mes) ? [mes, ...fs.readdirSync(mes, { withFileTypes: true }).filter((e) => e.isDirectory()).map((e) => path.join(mes, e.name))] : [];
  const ya = dirs.flatMap((x) => fs.readdirSync(x).map((f) => path.join(x, f))).find((f) => re.test(path.basename(f)));
  if (ya) return 'ya existía: ' + path.relative(mes, ya);
  const dest = path.join(rutas.carpetaIva(cuit, per), `${per}-IVA-DDJJ-${tokenDe(cuit)}.pdf`);
  fs.copyFileSync(pdf, dest);
  return path.relative(rutas.raiz(), dest);
}

async function unaSociedad(pedida) {
  let cred;
  try {
    cred = resolveCredentials({ cuit: pedida.cuit, nombre: pedida.nombre });
  } catch (e) {
    return { sociedad: pedida.nombre || pedida.cuit, ok: false, error: e.message,
      sugerencia: `Revisa el .env (${envFile() || 'no encontrado'}) o el Excel de claves.` };
  }
  const { cuit, cuitLogin, password, titular } = cred;
  const SALIDA = args['salida'] || path.join(SALIDAS, 'ddjj-iva', cuit);
  fs.mkdirSync(SALIDA, { recursive: true });
  const out = { ok: false, cuit, sociedad: titular || pedida.nombre || cuit, ddjj: [] };
  log(`${out.sociedad} · CUIT ${cuit} · ${DESDE}-${HASTA}`);
  const browser = await chromium.launch(launchOptions());
  const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1366, height: 900 } });
  const page = await context.newPage();
  try {
    await login(page, { cuitLogin, password });
    const iva = await pi.abrirPortalIva(context, page);
    out.representando = (await pi.asegurarRepresentacion(iva, cuit)).linea;
    // URL del Portal IVA (la DDJJ puede abrirse en la misma pestaña y llevarla a siapweb-f2051)
    const base = iva.url().replace(/#.*$/, '');
    for (const per of periodos()) {
      const mmaaaa = `${per.slice(4)}/${per.slice(0, 4)}`;
      let d;
      for (let intento = 1; intento <= 3; intento++) {
        d = { periodo: per };
        try {
          await abrirListaDdjj(iva, base);
          // filas del periodo: la primera es la ultima secuencia (rectificativa mas reciente)
          const filas = iva.locator('tr').filter({ has: iva.locator('td', { hasText: new RegExp(`^\\s*${per.slice(4)}\\/${per.slice(0, 4)}\\s*$`) }) });
          const n = await filas.count();
          if (!n) { d.falta = true; log(`${mmaaaa}: sin DDJJ presentada`); break; }
          d.secuencias = n;
          const nueva = context.waitForEvent('page', { timeout: 15000 }).catch(() => null);
          await filas.first().locator('button[title="Ver declaración jurada"]').click();
          let v = await nueva; const propia = !v; if (!v) v = iva;
          await v.waitForLoadState('domcontentloaded'); await v.waitForURL(/f2051/, { timeout: 60000 });
          for (let i = 0; i < 3 && !/vistaPrevia/.test(v.url()); i++) {
            await v.locator('button:has-text("CONTINUAR"), a:has-text("CONTINUAR")').first().waitFor({ state: 'visible', timeout: 45000 });
            await v.waitForTimeout(800);
            await v.locator('button:has-text("CONTINUAR"), a:has-text("CONTINUAR")').first().click();
            await v.waitForTimeout(2500);
          }
          if (!/vistaPrevia/.test(v.url())) throw new Error('no llegué a la vista previa');
          await v.waitForSelector('text=Determinación del impuesto', { timeout: 45000 });
          await v.waitForTimeout(1500);
          const t = await v.innerText('body');
          Object.assign(d, leerVista(t));
          if (d.periodo_vista && d.periodo_vista !== per) throw new Error(`la vista previa es de ${d.periodo_vista}`);
          d.pdf = path.join(SALIDA, `${per}.pdf`);
          await v.pdf({ path: d.pdf, format: 'A4', printBackground: true });
          fs.writeFileSync(path.join(SALIDA, `${per}.json`), JSON.stringify(d, null, 1));
          if (ARCHIVAR) {
            try { d.archivo = archivar(cuit, per, d.pdf); } catch (e) { d.archivo_error = e.message; }
          }
          log(`${mmaaaa}: DF ${d.df} CF ${d.cf} ST ant ${d.st_ant} SLD ant ${d.sld_ant} ret ${d.ret} ${d.archivo || ''}`);
          if (!propia) await v.close().catch(() => {});
          break;
        } catch (e) {
          d.error = e.message.split('\n')[0]; log(`${mmaaaa}: intento ${intento} ERROR ${d.error}`);
          for (const pg of context.pages()) if (pg !== iva && pg !== page) await pg.close().catch(() => {});
        }
      }
      out.ddjj.push(d);
    }
    out.ok = out.ddjj.every((x) => x.falta || (!x.error && x.df != null));
  } catch (e) {
    out.code = e.code || 'error';
    out.error = e.code === 'captcha'
      ? 'ARCA pidio un captcha. Entra una vez a mano con esa clave, resolvelo, y reintenta.'
      : e.code === 'credenciales_invalidas'
        ? `ARCA rechazo las credenciales del CUIT ${cuit}. Verifica la fila en Claves_Organismos.xlsx; ojo con reintentar, ARCA bloquea la clave.`
        : e.message.split('\n')[0];
    log('ERROR ' + out.error);
  } finally { await browser.close().catch(() => {}); }
  return out;
}

(async () => {
  const resultados = [];
  for (const [i, pedida] of PEDIDAS.entries()) {
    prefijo = PEDIDAS.length > 1 ? ` [${i + 1}/${PEDIDAS.length}]` : '';
    const r = await unaSociedad(pedida);
    resultados.push(r);
    if (['captcha', 'credenciales_invalidas'].includes(r.code)) { log('corto el lote: ARCA pidio captcha o rechazo la clave'); break; }
  }
  prefijo = '';
  const ok = resultados.length === PEDIDAS.length && resultados.every((r) => r.ok);
  console.log(JSON.stringify(PEDIDAS.length === 1 ? resultados[0] : {
    ok, desde: DESDE, hasta: HASTA, procesadas: resultados.length,
    con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, error: r.error })),
    resultados,
  }, null, 1));
  process.exit(ok ? 0 : 1);
})();
