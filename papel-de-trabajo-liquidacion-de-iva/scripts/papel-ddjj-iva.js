#!/usr/bin/env node
/**
 * papel-ddjj-iva.js — etapa 2 del papel de trabajo (DDJJ de IVA) de punta a punta, sobre el
 * MISMO papel "AAAAMM- papeles de trabajo- <Farmacia>.xlsx" que armo la etapa 1 (ventas).
 *
 *   node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202609
 *   node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202609 --desde-paso=7 --cuentas="<CUIT>=HONORARIOS"
 *   node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202610 --desde-paso=2 --portal=<libro bajado a mano>
 *
 * Pasos (cada uno con el script de scripts/; el papel se respalda antes de tocarlo):
 *   1. Zetti 5.6.4 (Playwright, login automatico) -> "AAAAMM - ZETTI - COMPRAS.xlsx" -> solapa "Compras Zetti".
 *   2. Libro IVA Compras del Portal IVA (con NCR) -> "AAAAMM - PORTAL IVA - COMPRAS.xlsx" -> solapa "compras con NCR"
 *      (tal cual, sin tocar) y, al lado, "Compras portal iva" = ese libro MENOS las NCR (notas de credito de recupero) de
 *      Monroe y Suizo de la conciliacion NCR del periodo (quitar_ncr.py): las compras como se presentan en el Portal IVA.
 *      Si el periodo no tiene conciliacion NCR, la arma con el plugin conciliacion-ncr (NCR del SharePoint + conciliar;
 *      NO elimina nada en ARCA).
 *   3. cruce_compras.py -> solapa "cruce compras".
 *   4. filtrar_servicios.py (AySA, EDESUR, METROGAS en Compras Zetti).
 *   5. agregar_servicios_portal.py (las agrega a Compras portal iva con su alicuota) y vuelve a correr el cruce.
 *   6. titulo_proveedores.py (AG1 = PROVEEDORES).
 *   7. clasificar_proveedores.py con el Maestro (+ asignar_cuentas.py si se pasa --cuentas), cuadro_compras.py y,
 *      debajo del cuadro, el detalle de las NCR ELIMINADAS con su control (bloque_ncr.py).
 *      Si queda algun NO CLASIFICADO, para (codigo 4) y lista los proveedores: volver a correr con
 *      --desde-paso=7 --cuentas="CUIT=CUENTA;CUIT=CUENTA".
 *   8. imputacion_compras.py -> solapa "imputacion compras" (resumen CMV / SERVICIOS y detalle por cuenta de AG).
 *
 * Periodo ya liquidado (la DDJJ de IVA esta guardada en la carpeta del mes o en IVA/: AAAAMM-IVA-ACUSE/DDJJ-*.pdf):
 *   usa la informacion que ya esta en la carpeta (Compras Zetti y el Libro Compras que todavia trae las NCR); lo que
 *   no este guardado lo baja de su fuente. Periodo nuevo (sin liquidar): baja todo de cada fuente, como siempre.
 *
 * Argumentos:
 *   --cuit, --periodo (AAAAMM o MM/AAAA).
 *   --descargar     Baja todo de las fuentes aunque el periodo este liquidado y la informacion este guardada.
 *   --sin-ncr       No eliminar NCR (la sociedad no tiene, o se decide asi): "Compras portal iva" = "compras con NCR".
 *   --portal        De donde sale el Libro IVA Compras (paso 2):
 *                     auto (default): periodo abierto -> importa desde ARCA al borrador del Libro Compras y lo
 *                       baja (escribe en el borrador de ARCA, no presenta nada). Si la DDJJ ya esta presentada,
 *                       baja la de la DDJJ presentada (solo lectura, controlada contra la vista previa).
 *                       Resguardo NCR: si la carpeta ncr/ del periodo ya tiene conciliacion o eliminacion de
 *                       NCR, no importa (las volveria a traer): presentada o, si no hay, codigo 4.
 *                     borrador: solo el borrador (si esta presentada, corta). Importa aunque haya NCR conciliadas.
 *                     presentada: solo la DDJJ presentada.
 *                     iva: usa el "AAAAMM - PORTAL IVA - COMPRAS.xlsx" que ya esta en la carpeta IVA del mes.
 *                     <ruta.xlsx o .csv>: usa ese archivo (ej. el que el usuario bajo a mano del Portal).
 *   --subcarpeta    El papel esta en 01-Impuestos Mensuales/<subcarpeta>/AAAAMM (ej. controles). Default: carpeta del mes.
 *   --papel         Ruta del papel (en vez de buscarlo).
 *   --desde-paso / --hasta-paso   Para retomar o cortar (1 a 8).
 *   --cuentas       "CUIT=CUENTA;..." para los NO CLASIFICADOS (paso 7). Solo el papel, no el Maestro.
 *   --antes-del-11  Corre aunque todavia no sea el dia 11 del mes siguiente al periodo.
 *   --ver           Navegador visible (solo si se pide).
 *
 * Nunca escribe en Zetti; en ARCA solo importa al borrador del Libro Compras (paso 2, periodo abierto). Salida: JSON por stdout (y en
 * BVA-salidas/papel-ddjj-iva/); progreso por stderr. Codigos: 0 OK, 1 fallo un paso, 2 argumentos o
 * papel no listo, 3 corte de ARCA (captcha/clave), 4 falta una decision (NO CLASIFICADOS, NCR ya conciliadas con el periodo abierto), 5 el control del cuadro o de la imputacion difiere en $1 o mas.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const rutas = require('../lib/rutas-bva');
const P = require('../lib/proceso');

const args = parseArgs();
loadEnv();
if ('ver' in args) process.env.PLAYWRIGHT_HEADLESS = 'false';
const log = (m) => process.stderr.write(`[papel-ddjj-iva ${new Date().toTimeString().slice(0, 5)}] ${m}\n`);

const CUIT = String(args.cuit || '').replace(/\D/g, '');
if (CUIT.length !== 11) fail('Falta --cuit (11 digitos)');
const PER = P.normalizarPeriodo(args.periodo);
if (!PER) fail('--periodo tiene que ser AAAAMM o MM/AAAA');
const DESDE = Number(args['desde-paso'] || 1), HASTA = Number(args['hasta-paso'] || 8);
if (!(DESDE >= 1 && HASTA <= 8 && DESDE <= HASTA)) fail('--desde-paso / --hasta-paso: de 1 a 8');
const PORTAL = String(args.portal || 'auto');
const DESCARGAR = 'descargar' in args;
const SIN_NCR = 'sin-ncr' in args;
const SCRIPTS = __dirname;
const SALIDAS = path.join(process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'papel-ddjj-iva');

// La etapa 2 no se corre antes del dia 11 del mes siguiente al periodo.
const desde11 = new Date(Number(PER.slice(0, 4)), Number(PER.slice(4)), 11);
if (new Date() < desde11 && !('antes-del-11' in args)) {
  fail(`La etapa 2 de ${PER} se corre desde el ${desde11.toLocaleDateString('es-AR')} (dia 11 del mes siguiente). Para correrla igual: --antes-del-11.`, { code: 'antes_del_11' });
}

let socDir;
try { socDir = rutas.carpetaSociedad(CUIT); } catch (e) { fail(e.message, { code: 'sin_carpeta' }); }
const FARM = P.nombreFarmacia(socDir, CUIT);
const MENSUALES = path.join(socDir, '01-Impuestos Mensuales');
const SUB = String(args.subcarpeta || '').trim();
const D = args.papel ? path.dirname(String(args.papel)) : SUB ? path.join(MENSUALES, SUB, PER) : rutas.carpetaMes(MENSUALES, PER);
const PA = args.papel ? String(args.papel) : path.join(D, `${PER}- papeles de trabajo- ${FARM}.xlsx`);

// Periodo ya liquidado: la DDJJ de IVA (acuse o formulario) esta guardada en la carpeta del mes o en IVA/.
function dirIva() {
  try { return rutas.carpetaIva(CUIT, PER, { crear: false }); } catch (e) { return null; }
}
function estaLiquidado() {
  const re = new RegExp(`^${PER}-IVA-(ACUSE|DDJJ)`, 'i');
  return [D, dirIva()].filter((d) => d && fs.existsSync(d)).some((d) => fs.readdirSync(d).some((f) => re.test(f)));
}
const LIQUIDADO = estaLiquidado();

const salida = { ok: false, cuit: CUIT, sociedad: path.basename(socDir), periodo: PER, papel: PA, liquidado: LIQUIDADO, pasos: {} };
function terminar(codigo, extra = {}) {
  Object.assign(salida, extra);
  salida.ok = codigo === 0;
  fs.mkdirSync(SALIDAS, { recursive: true });
  const f = path.join(SALIDAS, `${CUIT}-${PER}-${new Date().toISOString().replace(/[:T]/g, '-').slice(0, 19)}.json`);
  fs.writeFileSync(f, JSON.stringify(salida, null, 1));
  salida.resultado_en = f;
  console.log(JSON.stringify(salida, null, 1));
  process.exit(codigo);
}

const errorDe = (r) => (r.json && !Array.isArray(r.json) && (r.json.error || (r.json.libros || []).filter((l) => !l.ok).map((l) => l.error).join(' | ')))
  || r.stderr.trim().split(/\r?\n/).slice(-2).join(' ') || `salio con codigo ${r.code}`;
const codeDe = (r) => (r.json && !Array.isArray(r.json) && (r.json.code || ((r.json.libros || []).find((l) => !l.ok) || {}).code)) || null;

const py = (script, extra = []) => P.correr(P.PYTHON, [path.join(SCRIPTS, script), ...extra], { cwd: SCRIPTS, log, prefijo: '   ' });
async function pyPaso(nombre, script, extra = []) {
  const r = await py(script, extra);
  if (r.code !== 0 || !r.json || r.json.error) {
    salida.pasos[nombre] = { ok: false, error: errorDe(r), ...(r.json && typeof r.json === 'object' ? { detalle: r.json } : {}) };
    terminar(1, { paso_fallido: nombre });
  }
  return r.json;
}
const sinPapel = ({ papel, solapas, ...resto }) => resto; // eslint-disable-line no-unused-vars

async function paso1() {
  const guardado = path.join(D, `${PER} - ZETTI - COMPRAS.xlsx`);
  if (LIQUIDADO && !DESCARGAR && fs.existsSync(guardado)) {
    log(`1/8 Compras Zetti: periodo liquidado, uso el que ya esta guardado (${path.basename(guardado)})`);
    const pegado = await pyPaso('1_compras_zetti', 'copiar_solapa.py', [guardado, PA, 'Compras Zetti']);
    salida.pasos['1_compras_zetti'] = { ok: true, modo: 'archivo guardado (periodo liquidado)', archivo: guardado,
      filas_solapa: pegado.filas, solapas: pegado.solapas };
    return;
  }
  log('1/8 Compras Zetti (5.6.4)');
  let r;
  for (let i = 1; i <= 2; i++) {
    r = await P.correr(process.execPath, [path.join(SCRIPTS, 'bajar-compras-zetti.js'), `--cuit=${CUIT}`, `--periodo=${PER}`, `--dir=${D}`, '--si-existe=reemplazar'], { log, prefijo: '   ' });
    const z = Array.isArray(r.json) ? r.json[0] : null;
    if (z && z.ok) {
      const pegado = await pyPaso('1_compras_zetti', 'copiar_solapa.py', [z.archivo, PA, 'Compras Zetti']);
      salida.pasos['1_compras_zetti'] = { ok: true, archivo: z.archivo, comprobantes: z.comprobantes, por_tc: z.por_tc, totales: z.totales,
        control_total_reporte: z.control_total_reporte, fechas_fuera_de_periodo: z.fechas_fuera_de_periodo,
        ...(z.aviso_nro ? { aviso_nro: z.aviso_nro } : {}), solapas: pegado.solapas };
      return;
    }
    if (z && ['control', 'sin_datos', 'sin_carpeta'].includes(z.code)) break;
    log(`   Zetti fallo${i < 2 ? ', reintento' : ''}: ${(z && z.error) || errorDe(r)}`);
  }
  const z = Array.isArray(r.json) ? r.json[0] : null;
  salida.pasos['1_compras_zetti'] = { ok: false, error: (z && z.error) || errorDe(r), ...(z && z.detalle ? { detalle: z.detalle } : {}) };
  terminar(1, { paso_fallido: '1_compras_zetti' });
}

// Archivos del circuito conciliacion-ncr en la carpeta ncr/ del periodo: la conciliacion
// (AAAAMM-COMPRAS-SIN-NCR-<sociedad>.xlsx) y el log de una eliminacion ejecutada
// (AAAAMM-ELIMINACION-NCR-<sello>.csv; las simulaciones no quedan ahi).
function rastrosNcr() {
  let dir;
  try { dir = rutas.carpetaNcr(CUIT, PER, { crear: false }); } catch (e) {
    // Mas de una carpeta ncr/ = no se sabe cual vale: se frena igual. Sin carpeta IVA = no hay NCR.
    return /carpetas ncr\//.test(e.message) ? [`carpeta ncr/ ambigua: ${e.message}`] : [];
  }
  if (!dir || !fs.existsSync(dir)) return [];
  const re = new RegExp(`^${PER}-(COMPRAS-SIN-NCR-.+\\.xlsx|ELIMINACION-NCR-\\d{8}-\\d{4}\\.csv)$`, 'i');
  return fs.readdirSync(dir).filter((f) => !f.startsWith('~$') && re.test(f));
}

// Baja (o toma de --portal) el Libro IVA Compras como siempre. Devuelve { origen, info }.
async function bajarLibro() {
  const destino = path.join(D, `${PER} - PORTAL IVA - COMPRAS.xlsx`);
  let origen, info;
  if (['auto', 'borrador', 'presentada'].includes(PORTAL)) {
    // auto (default): periodo abierto -> importa desde ARCA al borrador y lo baja (caso normal);
    // si ARCA ya tiene acuse de la DDJJ (ddjj_presentada), baja la DDJJ presentada en solo lectura.
    const pPortal = P.pluginDir('portal-iva-descarga');
    let presentada = PORTAL === 'presentada';
    // Resguardo NCR: el circuito conciliacion-ncr elimina las NCR del borrador DESPUES de la ultima
    // importacion; importar de nuevo las vuelve a traer. Si el periodo ya tiene conciliacion o
    // eliminacion de NCR, auto no importa: busca la DDJJ presentada y, si no hay, frena (codigo 4).
    // --portal=borrador importa igual (decision explicita del usuario).
    const ncr = PORTAL === 'auto' ? rastrosNcr() : [];
    if (ncr.length) {
      log(`   el periodo ya tiene NCR conciliadas/eliminadas (${ncr.join(', ')}): no importo al borrador`);
      presentada = true;
    }
    let r;
    for (let i = 1; i <= 3; i++) {
      log(`2/8 Libro IVA Compras del Portal IVA (${presentada ? 'DDJJ presentada, solo lectura' : 'borrador: importa desde ARCA'})`);
      r = await P.correr(process.execPath, ['scripts/portal-iva-descarga.js', `--cuit=${CUIT}`, `--periodo=${P.mmaaaa(PER)}`, '--libros=compras', '--no-archivar',
        '--si-existe=reemplazar', ...(presentada ? ['--presentada'] : [])], { cwd: pPortal, log, prefijo: '   ' });
      if (r.code === 0) break;
      const c = codeDe(r);
      if (['captcha', 'credenciales_invalidas'].includes(c)) {
        salida.pasos['2_compras_portal'] = { ok: false, error: errorDe(r), code: c };
        terminar(3, { paso_fallido: '2_compras_portal' });
      }
      if (c === 'ddjj_presentada' && PORTAL === 'auto' && !presentada) {
        log('   la DDJJ del periodo ya esta presentada: bajo el Libro Compras de la DDJJ presentada');
        presentada = true;
        continue;
      }
      if (c === 'sin_presentacion' && presentada && ncr.length) {
        salida.pasos['2_compras_portal'] = { ok: false, code: 'ncr_ya_conciliadas', error: errorDe(r) };
        terminar(4, { paso_fallido: '2_compras_portal', ncr: ncr, falta_decidir:
          `La DDJJ de IVA ${PER} no esta presentada y el periodo ya tiene NCR conciliadas o eliminadas del borrador ` +
          `(${ncr.join(', ')}). Importar de nuevo desde ARCA las volveria a traer. Volver a correr con --desde-paso=2 y ` +
          '--portal=iva (el libro de la carpeta IVA del mes), --portal=<ruta del libro bajado a mano> o, solo con el OK del ' +
          'usuario y sabiendo que hay que volver a eliminar las NCR despues, --portal=borrador.' });
      }
      if (c === 'sin_presentacion' && presentada) {
        salida.pasos['2_compras_portal'] = { ok: false, code: 'sin_presentacion', error: errorDe(r) };
        terminar(4, { paso_fallido: '2_compras_portal', falta_decidir:
          `La DDJJ de IVA ${PER} no esta presentada. Volver a correr con --desde-paso=2 y --portal=auto (importa al borrador), ` +
          '--portal=iva (el que ya esta en la carpeta IVA del mes) o --portal=<ruta.xlsx>.' });
      }
      if (['ddjj_presentada', 'otro_borrador_abierto'].includes(c) || /no coincide con la DDJJ/i.test(errorDe(r))) break;
    }
    const libro = r.json && (r.json.libros || [])[0];
    if (r.code !== 0 || !libro || !libro.archivo) {
      salida.pasos['2_compras_portal'] = { ok: false, error: errorDe(r), code: codeDe(r) };
      terminar(1, { paso_fallido: '2_compras_portal', ...(PORTAL !== 'presentada' && !presentada ? { falta_decidir:
        'No se pudo bajar el Libro Compras del Portal IVA. Si el usuario lo baja a mano, volver a correr con --desde-paso=2 ' +
        '--portal=<ruta del .xlsx o .csv>; o --portal=iva si ya esta en la carpeta IVA del mes.' } : {}) });
    }
    origen = libro.archivo;
    info = { modo: presentada ? 'presentada' : 'borrador', ddjj: r.json.ddjj, filas: libro.filas,
      filas_fuera_de_periodo: libro.filasFueraDePeriodo, control_ddjj: libro.control_ddjj, importacion: libro.importacion };
  } else {
    origen = PORTAL === 'iva' ? path.join(rutas.carpetaIva(CUIT, PER, { crear: false }), `${PER} - PORTAL IVA - COMPRAS.xlsx`) : PORTAL;
    if (!fs.existsSync(origen)) {
      salida.pasos['2_compras_portal'] = { ok: false, error: `No existe ${origen}` };
      terminar(2, { paso_fallido: '2_compras_portal' });
    }
    log(`2/8 Libro IVA Compras desde ${origen}`);
    info = { modo: 'archivo', origen };
    // Libro bajado a mano del Portal: el .zip o el .csv de ARCA se pasan a Excel con el mismo conversor
    // del plugin portal-iva-descarga (mismo formato que la descarga automatica).
    if (/\.(zip|csv)$/i.test(origen)) {
      let csv = origen;
      if (/\.zip$/i.test(origen)) {
        const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'papel-ddjj-iva-'));
        const z = await P.correr(P.PYTHON, ['-c', 'import sys,zipfile;z=zipfile.ZipFile(sys.argv[1]);n=[x for x in z.namelist() if x.lower().endswith(".csv")][0];print(z.extract(n,sys.argv[2]))', origen, tmp], { log });
        csv = (z.stdout || '').trim().split(String.fromCharCode(10)).pop().trim();
        if (z.code !== 0 || !csv || !fs.existsSync(csv)) {
          salida.pasos['2_compras_portal'] = { ok: false, error: `No encontre el CSV dentro de ${origen}` };
          terminar(2, { paso_fallido: '2_compras_portal' });
        }
      }
      const pPortal = P.pluginDir('portal-iva-descarga');
      const c = await P.correr(P.PYTHON, [path.join(pPortal, 'scripts', 'csv_a_excel.py'), csv, destino, '--libro=compras'], { log });
      if (c.code !== 0) {
        salida.pasos['2_compras_portal'] = { ok: false, error: `No pude convertir ${csv}: ${errorDe(c)}` };
        terminar(1, { paso_fallido: '2_compras_portal' });
      }
      origen = destino;
    }
  }
  return { origen, info };
}

// Conciliacion NCR del periodo (circuito conciliacion-ncr): ncr/AAAAMM-COMPRAS-SIN-NCR-<SOC>.xlsx, la mas nueva.
function conciliacionNcr() {
  let dir;
  try { dir = rutas.carpetaNcr(CUIT, PER, { crear: false }); } catch (e) { return null; }
  if (!dir || !fs.existsSync(dir)) return null;
  const re = new RegExp(`^${PER}-COMPRAS-SIN-NCR-.+\\.xlsx$`, 'i');
  const hay = fs.readdirSync(dir).filter((f) => !f.startsWith('~$') && re.test(f)).map((f) => path.join(dir, f));
  return hay.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)[0] || null;
}

// Libros de compras ya guardados en la carpeta del periodo (al lado del papel y en IVA/).
function librosGuardados() {
  const iva = dirIva();
  return [path.join(D, `${PER} - PORTAL IVA - COMPRAS.xlsx`), ...(iva ? [path.join(iva, `${PER} - PORTAL IVA - COMPRAS.xlsx`)] : [])]
    .filter((f, i, a) => fs.existsSync(f) && a.indexOf(f) === i);
}

// El libro trae TODAS las NCR de la conciliacion (el de la DDJJ presentada o el borrador ya depurado, no).
async function traeLasNcr(libro, conc) {
  const r = await py('quitar_ncr.py', [`--origen=${libro}`, `--ncr=${conc}`, '--verificar']);
  return !!(r.json && r.json.contiene_todas);
}

// Arma la conciliacion NCR con el plugin conciliacion-ncr: baja las NCR del SharePoint y concilia contra el libro.
// No elimina nada del borrador de ARCA (eso es el paso 4 de conciliacion-ncr y no se hace desde aca).
async function conciliarNcr(libro) {
  const parar = (code, error) => {
    salida.pasos['2_compras_portal'] = { ok: false, code, error };
    terminar(code === 'credenciales' ? 3 : 4, { paso_fallido: '2_compras_portal', falta_decidir:
      `No se pudo armar la conciliacion NCR de ${PER} (${error}). Correr la conciliacion con el plugin conciliacion-ncr ` +
      '(pasos 2 y 3) y volver a correr con --desde-paso=2, o, si la sociedad no tiene NCR para eliminar, --desde-paso=2 --sin-ncr.' });
  };
  let pNcr;
  try { pNcr = P.pluginDir('conciliacion-ncr'); } catch (e) { parar('falta_plugin', e.message); }
  log('2/8 conciliacion NCR: no esta en la carpeta ncr/, la armo (NCR del SharePoint + conciliar, sin eliminar nada en ARCA)');
  const d = await P.correr(process.execPath, ['scripts/descargar-ncr.js', `--cuit=${CUIT}`, `--periodo=${PER}`], { cwd: pNcr, log, prefijo: '   ' });
  if (d.code !== 0) parar(['captcha', 'credenciales_invalidas'].includes(codeDe(d)) ? 'credenciales' : (codeDe(d) || 'ncr_sharepoint'), errorDe(d));
  const c = await P.correr(P.PYTHON, ['scripts/conciliar-ncr.py', `--cuit=${CUIT}`, `--periodo=${PER}`, `--portal=${libro}`, '--si-existe=reemplazar'],
    { cwd: pNcr, log, prefijo: '   ' });
  if (c.code !== 0 || !c.json || c.json.error) parar('conciliar_ncr', errorDe(c));
  const conc = conciliacionNcr();
  if (!conc) parar('conciliar_ncr', 'conciliar-ncr.py no dejo AAAAMM-COMPRAS-SIN-NCR-*.xlsx en la carpeta ncr/');
  salida.conciliacion_ncr_armada = { archivo: conc, ...(c.json && typeof c.json === 'object' ? c.json : {}) };
  return conc;
}

async function paso2() {
  const destino = path.join(D, `${PER} - PORTAL IVA - COMPRAS.xlsx`);
  let conc = SIN_NCR ? null : conciliacionNcr();
  let origen = null, info = null;
  // 1) Periodo liquidado o NCR ya conciliadas: el libro CON NCR que ya esta guardado (no se baja ni se importa;
  //    importar al borrador volveria a traer las NCR eliminadas).
  if (!DESCARGAR && PORTAL === 'auto' && (LIQUIDADO || conc)) {
    for (const f of librosGuardados()) {
      if (conc ? await traeLasNcr(f, conc) : LIQUIDADO) {
        origen = f;
        info = { modo: 'archivo guardado', origen: f, motivo: LIQUIDADO ? 'periodo liquidado' : 'NCR ya conciliadas' };
        log(`2/8 Libro IVA Compras: uso el guardado ${f} (${info.motivo})`);
        break;
      }
    }
  }
  // 2) Si no esta guardado: de la fuente, como siempre.
  if (!origen) ({ origen, info } = await bajarLibro());
  // 3) Sin conciliacion NCR: se arma (salvo --sin-ncr).
  if (!SIN_NCR && !conc) conc = await conciliarNcr(origen);
  // 4) El libro tiene que traer las NCR que se van a eliminar; si no (ej. DDJJ presentada), uno guardado que si.
  if (conc && !(await traeLasNcr(origen, conc))) {
    let otro = null;
    for (const f of librosGuardados()) {
      if (path.resolve(f) !== path.resolve(origen) && await traeLasNcr(f, conc)) { otro = f; break; }
    }
    if (!otro) {
      salida.pasos['2_compras_portal'] = { ok: false, code: 'libro_sin_ncr', origen, conciliacion_ncr: conc };
      terminar(4, { paso_fallido: '2_compras_portal', falta_decidir:
        `El Libro Compras (${info.modo}) ya no trae las NCR de la conciliacion ${path.basename(conc)} y no hay otro guardado ` +
        'que las tenga. Pasar el libro con NCR con --desde-paso=2 --portal=<ruta del .xlsx/.zip/.csv>, o seguir sin eliminar ' +
        'NCR con --desde-paso=2 --sin-ncr.' });
    }
    log(`   el libro bajado no trae las NCR: uso el guardado ${otro}`);
    origen = otro;
    info = { ...info, modo: 'archivo guardado', origen: otro, motivo: 'el libro bajado ya no trae las NCR' };
  }
  if (path.resolve(origen) !== path.resolve(destino)) fs.copyFileSync(origen, destino);

  const conNcr = await pyPaso('2_compras_portal', 'copiar_solapa.py', [destino, PA, 'compras con NCR', '--antes=Compras portal iva']);
  const tmp = path.join(os.tmpdir(), `papel-ddjj-iva-${CUIT}-${PER}-sin-ncr.xlsx`);
  const q = await pyPaso('2_compras_portal', 'quitar_ncr.py', [`--origen=${destino}`, conc ? `--ncr=${conc}` : '--sin-ncr', `--salida=${tmp}`]);
  const pegado = await pyPaso('2_compras_portal', 'copiar_solapa.py', [tmp, PA, 'Compras portal iva']);
  try { fs.unlinkSync(tmp); } catch (e) { /* queda en temp */ }
  salida.pasos['2_compras_portal'] = { ok: true, archivo: destino, ...info, conciliacion_ncr: conc,
    compras_con_ncr: { comprobantes: q.comprobantes_con_ncr, importe_total: q.importe_total_con_ncr, filas_solapa: conNcr.filas },
    ncr_eliminadas: { cantidad: q.ncr_eliminadas, por_drogueria: q.ncr_por_drogueria, importe: q.importe_ncr },
    compras_portal_iva: { comprobantes: q.comprobantes_sin_ncr, importe_total: q.importe_total_sin_ncr },
    control: q.control, solapas: pegado.solapas };
}

async function paso3(nombre = '3_cruce_compras') {
  log(`${nombre === '3_cruce_compras' ? '3/8' : '5/8'} cruce compras (Zetti vs Portal IVA)`);
  const j = await pyPaso(nombre, 'cruce_compras.py', [`--papel=${PA}`]);
  salida.pasos[nombre] = { ok: true, ...sinPapel(j) };
}

async function paso7() {
  log('7/8 clasificacion de proveedores con el Maestro');
  let j = await pyPaso('7_clasificacion', 'clasificar_proveedores.py', [`--papel=${PA}`]);
  const clasif = { ok: true, maestro: j.maestro, comprobantes: j.comprobantes, por_cuenta: j.por_cuenta,
    ...(j.cuits_con_dos_cuentas_en_padron ? { cuits_con_dos_cuentas_en_padron: j.cuits_con_dos_cuentas_en_padron } : {}) };
  let pendientes = j.no_clasificados || [];
  if (pendientes.length && args.cuentas) {
    const a = await pyPaso('7_clasificacion', 'asignar_cuentas.py', [`--papel=${PA}`, `--cuentas=${args.cuentas}`]);
    clasif.asignados_a_mano = a.asignados;
    pendientes = a.no_clasificados;
  }
  clasif.no_clasificados = pendientes;
  salida.pasos['7_clasificacion'] = clasif;
  if (pendientes.length) {
    terminar(4, { paso_fallido: '7_clasificacion', falta_decidir:
      'Hay proveedores NO CLASIFICADOS: definir la cuenta de cada uno y volver a correr con --desde-paso=7 --cuentas="CUIT=CUENTA;CUIT=CUENTA" ' +
      '(y cargarlos en el Maestro de Proveedores para la proxima).' });
  }
  log('7/8 cuadro resumen de compras');
  const r = await py('cuadro_compras.py', [`--papel=${PA}`]);
  if (!r.json || r.json.error) {
    salida.pasos['7_cuadro'] = { ok: false, error: errorDe(r) };
    terminar(1, { paso_fallido: '7_cuadro' });
  }
  salida.pasos['7_cuadro'] = { ok: r.code === 0, ...sinPapel(r.json) };
  // Detalle de las NCR ELIMINADAS debajo del cuadro (va despues: cuadro_compras.py borra lo que hay abajo de la nota).
  const sol = (await P.correr(P.PYTHON, [path.join(SCRIPTS, 'solapas.py'), PA])).json;
  if (sol && (sol.solapas || []).includes('compras con NCR')) {
    const conc = SIN_NCR ? null : conciliacionNcr();
    log('7/8 NCR eliminadas debajo del cuadro');
    const b = await pyPaso('7_ncr_eliminadas', 'bloque_ncr.py', [`--papel=${PA}`, conc ? `--ncr=${conc}` : '--sin-ncr']);
    salida.pasos['7_ncr_eliminadas'] = { ok: b.ok !== false, ...sinPapel(b) };
  }
  if (r.code === 2) terminar(5, { paso_fallido: '7_cuadro', aviso: 'El control del cuadro difiere en $1 o mas: revisar antes de usar el papel.' });
}

async function paso8() {
  log('8/8 imputacion compras (resumen por clasificacion y detalle por cuenta)');
  const r = await py('imputacion_compras.py', [`--papel=${PA}`]);
  if (!r.json || r.json.error || (r.code !== 0 && r.code !== 2)) {
    salida.pasos['8_imputacion_compras'] = { ok: false, error: errorDe(r) };
    terminar(1, { paso_fallido: '8_imputacion_compras' });
  }
  salida.pasos['8_imputacion_compras'] = { ok: r.code === 0, ...sinPapel(r.json) };
  if (r.code === 2) terminar(5, { paso_fallido: '8_imputacion_compras', aviso: 'El control de la imputacion de compras no cierra: revisar antes de usar el papel.' });
}

(async () => {
  if (!fs.existsSync(PA)) fail(`No existe el papel ${PA}: la etapa 1 (papel-trabajo-ventas-tsh) no esta hecha${SUB ? '' : ' (si esta en controles: --subcarpeta=controles)'}.`, { code: 'sin_papel' });
  if (P.abiertoEnExcel(PA)) fail('El papel esta abierto en Excel: cerrarlo y volver a correr.', { code: 'papel_abierto' });
  const sol = (await P.correr(P.PYTHON, [path.join(SCRIPTS, 'solapas.py'), PA])).json;
  const solapas = (sol && sol.solapas) || [];
  if (!solapas.some((s) => /^cruce( ventas)?$/i.test(s))) fail('El papel no tiene la solapa "cruce ventas": la etapa 1 no esta terminada.', { code: 'etapa1_incompleta', solapas });
  const necesita = { 3: ['Compras Zetti', 'Compras portal iva'], 4: ['Compras Zetti'], 5: ['Compras Zetti', 'Compras portal iva'], 6: ['Compras portal iva'], 7: ['Compras portal iva'], 8: ['Compras portal iva'] };
  const faltan = (necesita[DESDE] || []).filter((s) => !solapas.includes(s));
  if (DESDE > 1 && faltan.length) fail(`Para empezar en el paso ${DESDE} el papel tiene que tener ${faltan.join(' y ')}.`, { code: 'faltan_solapas' });

  const dir = path.join(SALIDAS, 'respaldos');
  fs.mkdirSync(dir, { recursive: true });
  salida.respaldo = path.join(dir, `${new Date().toISOString().replace(/[:T]/g, '-').slice(0, 19)} - ${path.basename(PA)}`);
  fs.copyFileSync(PA, salida.respaldo);
  log(`${path.basename(socDir)} | ${PER} | pasos ${DESDE} a ${HASTA} | respaldo en ${salida.respaldo}`);

  const hace = (n) => DESDE <= n && n <= HASTA;
  if (hace(1)) await paso1();
  if (hace(2)) await paso2();
  if (hace(3)) await paso3();
  if (hace(4)) {
    log('4/8 servicios en Compras Zetti (AySA, EDESUR, METROGAS)');
    salida.pasos['4_servicios_zetti'] = { ok: true, ...sinPapel(await pyPaso('4_servicios_zetti', 'filtrar_servicios.py', [`--papel=${PA}`])) };
  }
  if (hace(5)) {
    log('5/8 servicios agregados a Compras portal iva');
    salida.pasos['5_servicios_portal'] = { ok: true, ...sinPapel(await pyPaso('5_servicios_portal', 'agregar_servicios_portal.py', [`--papel=${PA}`])) };
    if ((salida.pasos['5_servicios_portal'].agregados || []).length) await paso3('5_cruce_compras_de_nuevo');
  }
  if (hace(6)) {
    log('6/8 titulo PROVEEDORES en AG');
    salida.pasos['6_titulo_proveedores'] = { ok: true, ...sinPapel(await pyPaso('6_titulo_proveedores', 'titulo_proveedores.py', [`--papel=${PA}`])) };
  }
  if (hace(7)) await paso7();
  if (hace(8)) await paso8();
  log('OK');
  terminar(0);
})().catch((e) => { salida.error = e.message; terminar(1); });
