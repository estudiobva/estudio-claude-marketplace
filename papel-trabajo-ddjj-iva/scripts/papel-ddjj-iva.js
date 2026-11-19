#!/usr/bin/env node
/**
 * papel-ddjj-iva.js — etapa 2 del papel de trabajo (DDJJ de IVA) de punta a punta, sobre el
 * MISMO papel "AAAAMM- papeles de trabajo- <Farmacia>.xlsx" que armo la etapa 1 (ventas).
 *
 *   node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202609
 *   node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202609 --desde-paso=7 --cuentas="<CUIT>=HONORARIOS"
 *   node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202610 --portal=borrador
 *
 * Pasos (cada uno con el script de scripts/; el papel se respalda antes de tocarlo):
 *   1. Zetti 5.6.4 (Playwright, login automatico) -> "AAAAMM - ZETTI - COMPRAS.xlsx" -> solapa "Compras Zetti".
 *   2. Libro IVA Compras del Portal IVA -> "AAAAMM - PORTAL IVA - COMPRAS.xlsx" -> solapa "Compras portal iva".
 *   3. cruce_compras.py -> solapa "cruce compras".
 *   4. filtrar_servicios.py (AySA, EDESUR, METROGAS en Compras Zetti).
 *   5. agregar_servicios_portal.py (las agrega a Compras portal iva con su alicuota) y vuelve a correr el cruce.
 *   6. titulo_proveedores.py (AG1 = PROVEEDORES).
 *   7. clasificar_proveedores.py con el Maestro (+ asignar_cuentas.py si se pasa --cuentas) y cuadro_compras.py.
 *      Si queda algun NO CLASIFICADO, para (codigo 4) y lista los proveedores: volver a correr con
 *      --desde-paso=7 --cuentas="CUIT=CUENTA;CUIT=CUENTA".
 *   8. imputacion_compras.py -> solapa "imputacion compras" (resumen CMV / SERVICIOS y detalle por cuenta de AG).
 *
 * Argumentos:
 *   --cuit, --periodo (AAAAMM o MM/AAAA).
 *   --portal        De donde sale el Libro IVA Compras (paso 2):
 *                     auto (default): el de la DDJJ PRESENTADA (solo lectura, controlado contra la DDJJ);
 *                       si el periodo no esta presentado, para (codigo 4) y pide elegir.
 *                     borrador: importa desde ARCA al borrador del Libro Compras y lo baja (escribe en el
 *                       borrador de ARCA, no presenta nada). Ojo: si ya se eliminaron NCR del borrador
 *                       (conciliacion-ncr), importar puede volver a traerlas.
 *                     iva: usa el "AAAAMM - PORTAL IVA - COMPRAS.xlsx" que ya esta en la carpeta IVA del mes.
 *                     <ruta.xlsx>: usa ese archivo.
 *   --subcarpeta    El papel esta en 01-Impuestos Mensuales/<subcarpeta>/AAAAMM (ej. controles). Default: carpeta del mes.
 *   --papel         Ruta del papel (en vez de buscarlo).
 *   --desde-paso / --hasta-paso   Para retomar o cortar (1 a 8).
 *   --cuentas       "CUIT=CUENTA;..." para los NO CLASIFICADOS (paso 7). Solo el papel, no el Maestro.
 *   --antes-del-11  Corre aunque todavia no sea el dia 11 del mes siguiente al periodo.
 *   --ver           Navegador visible (solo si se pide).
 *
 * Nunca escribe en Zetti; en ARCA solo con --portal=borrador. Salida: JSON por stdout (y en
 * BVA-salidas/papel-ddjj-iva/); progreso por stderr. Codigos: 0 OK, 1 fallo un paso, 2 argumentos o
 * papel no listo, 3 corte de ARCA (captcha/clave), 4 falta una decision (NO CLASIFICADOS, periodo
 * sin presentar), 5 el control del cuadro o de la imputacion difiere en $1 o mas.
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

const salida = { ok: false, cuit: CUIT, sociedad: path.basename(socDir), periodo: PER, papel: PA, pasos: {} };
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

async function paso2() {
  const destino = path.join(D, `${PER} - PORTAL IVA - COMPRAS.xlsx`);
  let origen, info;
  if (PORTAL === 'auto' || PORTAL === 'borrador') {
    const pPortal = P.pluginDir('portal-iva-descarga');
    const presentada = PORTAL === 'auto';
    log(`2/8 Libro IVA Compras del Portal IVA (${presentada ? 'DDJJ presentada, solo lectura' : 'borrador: importa desde ARCA'})`);
    let r;
    for (let i = 1; i <= 2; i++) {
      r = await P.correr(process.execPath, ['scripts/portal-iva-descarga.js', `--cuit=${CUIT}`, `--periodo=${P.mmaaaa(PER)}`, '--libros=compras', '--no-archivar',
        ...(presentada ? ['--presentada'] : [])], { cwd: pPortal, log, prefijo: '   ' });
      if (r.code === 0) break;
      const c = codeDe(r);
      if (['captcha', 'credenciales_invalidas'].includes(c)) {
        salida.pasos['2_compras_portal'] = { ok: false, error: errorDe(r), code: c };
        terminar(3, { paso_fallido: '2_compras_portal' });
      }
      if (c === 'sin_presentacion' && presentada) {
        salida.pasos['2_compras_portal'] = { ok: false, code: 'sin_presentacion', error: errorDe(r) };
        terminar(4, { paso_fallido: '2_compras_portal', falta_decidir:
          `La DDJJ de IVA ${PER} no esta presentada. Elegir de donde sale el Libro IVA Compras y volver a correr con --desde-paso=2: ` +
          '--portal=iva (el que ya esta en la carpeta IVA del mes), --portal=<ruta.xlsx>, o --portal=borrador (importa al borrador de ARCA; ' +
          'si ya se eliminaron NCR del borrador, puede volver a traerlas).' });
      }
      if (['ddjj_presentada', 'otro_borrador_abierto'].includes(c) || /no coincide con la DDJJ/i.test(errorDe(r))) break;
    }
    const libro = r.json && (r.json.libros || [])[0];
    if (r.code !== 0 || !libro || !libro.archivo) {
      salida.pasos['2_compras_portal'] = { ok: false, error: errorDe(r), code: codeDe(r) };
      terminar(1, { paso_fallido: '2_compras_portal' });
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
  }
  if (path.resolve(origen) !== path.resolve(destino)) fs.copyFileSync(origen, destino);
  const pegado = await pyPaso('2_compras_portal', 'copiar_solapa.py', [destino, PA, 'Compras portal iva']);
  salida.pasos['2_compras_portal'] = { ok: true, archivo: destino, ...info, comprobantes: pegado.filas, solapas: pegado.solapas };
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
