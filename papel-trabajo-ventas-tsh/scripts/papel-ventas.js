#!/usr/bin/env node
/**
 * papel-ventas.js — etapa 1 del papel de trabajo de ventas (TSH / IIBB) de punta a punta.
 *
 *   node scripts/papel-ventas.js --cuit=<CUIT> --periodos=202509,202510,202511
 *   node scripts/papel-ventas.js --cuit=<CUIT> --desde=202601 --hasta=202609
 *   node scripts/papel-ventas.js --cuit=<CUIT> --periodos=202609 --subcarpeta=
 *   node scripts/papel-ventas.js --cuit=<CUIT> --solo-cruce            (todos los meses de la subcarpeta)
 *
 * Por cada periodo, en 01-Impuestos Mensuales/<subcarpeta>/AAAAMM de la sociedad:
 *   1. Zetti 5.6.5 (Playwright, login automatico) -> "AAAAMM- papeles de trabajo- <Farmacia>.xlsx"
 *      con las solapas Ventas Zetti y Ventas Zetti (depurado).
 *   2. Libro IVA Ventas de la DDJJ PRESENTADA (plugin portal-iva-descarga --presentada, solo lectura)
 *      -> "AAAAMM - PORTAL IVA - VENTAS.xlsx". Si no coincide con la DDJJ, el mes se frena.
 *   3. Mis Comprobantes emitidos (plugin cruce-ventas-3-fuentes, solo lectura)
 *      -> "AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx".
 *   4. armar_papel.py: pega Portal y MC, TOTAL de MC, depura PV y arma el cuadro resumen.
 * Al final, con todos los meses armados, el cruce de ventas (solapa "cruce ventas") de cada mes
 * y un resumen: Zetti - Portal IVA, estados y observaciones para revisar.
 *
 * Argumentos:
 *   --cuit          CUIT de la sociedad (la carpeta se busca por CUIT en la unidad compartida).
 *   --periodos      AAAAMM separados por coma. O --desde / --hasta (AAAAMM o MM/AAAA).
 *   --subcarpeta    Default "controles" (01-Impuestos Mensuales/controles/AAAAMM).
 *                   --subcarpeta= (vacio) deja el papel en la carpeta del mes (01-Impuestos Mensuales/AAAAMM).
 *   --si-existe     saltear (default: el mes con papel completo no se rehace) | reemplazar.
 *   --forzar        Rehace aunque el papel ya tenga solapas de la etapa 2 (Compras ...). Sin esto, no lo toca.
 *   --sin-cruce     No corre el cruce al final.
 *   --solo-cruce    Solo el cruce y el resumen (sin --periodos: todos los meses de la subcarpeta).
 *   --excepcion     Se pasa a depurar_pv / cruce (ej. "2:NC B").
 *   --ver           Navegador visible (solo si se pide).
 *
 * Nunca escribe en ARCA ni en Zetti. Antes de reemplazar un papel existente lo copia a
 * ~/Documents/BVA-salidas/papel-ventas-tsh/respaldos. Salida: JSON por stdout (y en
 * BVA-salidas/papel-ventas-tsh/<CUIT>-<fecha>.json); progreso por stderr.
 * Codigo de salida: 0 todo OK, 1 algun mes fallo, 2 argumentos, 3 corte de ARCA (captcha/clave).
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
const log = (m) => process.stderr.write(`[papel-ventas ${new Date().toTimeString().slice(0, 5)}] ${m}\n`);

const CUIT = String(args.cuit || '').replace(/\D/g, '');
if (CUIT.length !== 11) fail('Falta --cuit (11 digitos)');
const SUB = 'subcarpeta' in args ? String(args.subcarpeta).trim() : 'controles';
const SI_EXISTE = String(args['si-existe'] || 'saltear').toLowerCase();
if (!['saltear', 'reemplazar'].includes(SI_EXISTE)) fail('--si-existe tiene que ser saltear o reemplazar');
const FORZAR = 'forzar' in args;
const SOLO_CRUCE = 'solo-cruce' in args;
const SIN_CRUCE = 'sin-cruce' in args;
const EXC = args.excepcion ? [`--excepcion=${args.excepcion}`] : [];
const SCRIPTS = __dirname;
const SALIDAS = path.join(process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'papel-ventas-tsh');

const SOLAPAS_ETAPA1 = ['Ventas Zetti', 'Ventas portal iva', 'Ventas MCR'];
const CORTES_ARCA = ['captcha', 'credenciales_invalidas'];

let socDir;
try { socDir = rutas.carpetaSociedad(CUIT); } catch (e) { fail(e.message, { code: 'sin_carpeta' }); }
const FARM = P.nombreFarmacia(socDir, CUIT);
const MENSUALES = path.join(socDir, '01-Impuestos Mensuales');
const dirMes = (per) => (SUB ? path.join(MENSUALES, SUB, per) : rutas.carpetaMes(MENSUALES, per));
const papelDe = (per) => path.join(dirMes(per), `${per}- papeles de trabajo- ${FARM}.xlsx`);

function periodos() {
  if (args.periodos) {
    const l = String(args.periodos).split(/[,;\s]+/).filter(Boolean).map(P.normalizarPeriodo);
    if (l.some((x) => !x)) fail('--periodos: cada uno AAAAMM o MM/AAAA');
    return [...new Set(l)].sort();
  }
  if (args.desde || args.hasta) {
    const d = P.normalizarPeriodo(args.desde), h = P.normalizarPeriodo(args.hasta || args.desde);
    if (!d || !h || d > h) fail('--desde / --hasta: AAAAMM o MM/AAAA, desde <= hasta');
    return P.rangoPeriodos(d, h);
  }
  if (SOLO_CRUCE) {
    const base = SUB ? path.join(MENSUALES, SUB) : MENSUALES;
    const anios = fs.existsSync(base) ? fs.readdirSync(base).filter((d) => /^\d{4}$/.test(d)).map((a) => path.join(base, a)) : [];
    const meses = [base, ...anios].flatMap((d) => (fs.existsSync(d) ? fs.readdirSync(d) : []).filter((x) => /^\d{6}$/.test(x)));
    return [...new Set(meses)].sort().filter((per) => fs.existsSync(papelDe(per)));
  }
  fail('Faltan --periodos (o --desde/--hasta)');
}

async function solapas(xlsx) {
  const r = await P.correr(P.PYTHON, [path.join(SCRIPTS, 'solapas.py'), xlsx]);
  return (r.json && r.json.solapas) || [];
}

function respaldar(xlsx) {
  const dir = path.join(SALIDAS, 'respaldos');
  fs.mkdirSync(dir, { recursive: true });
  const dest = path.join(dir, `${new Date().toISOString().replace(/[:T]/g, '-').slice(0, 19)} - ${path.basename(xlsx)}`);
  fs.copyFileSync(xlsx, dest);
  return dest;
}

const errorDe = (r) => (r.json && (r.json.error || (r.json.libros || []).filter((l) => !l.ok).map((l) => l.error).join(' | ')))
  || r.stderr.trim().split(/\r?\n/).slice(-2).join(' ') || `salio con codigo ${r.code}`;
const codeDe = (r) => (r.json && (r.json.code || ((r.json.libros || []).find((l) => !l.ok) || {}).code)) || null;

async function unMes(per, pPortal, pMc) {
  const D = dirMes(per), PA = papelDe(per);
  const fila = { periodo: per, ok: false, papel: PA };
  if (fs.existsSync(PA)) {
    if (P.abiertoEnExcel(PA)) return { ...fila, paso: 'papel', error: 'El papel esta abierto en Excel: cerrarlo y volver a correr' };
    const sol = await solapas(PA);
    const etapa2 = sol.filter((s) => /^compras|cruce compras/i.test(s));
    if (etapa2.length && !FORZAR) {
      return { ...fila, paso: 'papel', error: `El papel ya tiene solapas de la etapa 2 (${etapa2.join(', ')}): no lo rehago. Usar --forzar si hay que rehacerlo (se pierden esas solapas; queda respaldo).` };
    }
    if (SI_EXISTE === 'saltear' && SOLAPAS_ETAPA1.every((s) => sol.includes(s))) {
      log(`${per}: el papel ya esta completo, lo salteo (--si-existe=reemplazar para rehacerlo)`);
      return { ...fila, ok: true, salteado: true };
    }
    fila.respaldo = respaldar(PA);
    log(`${per}: respaldo del papel anterior en ${fila.respaldo}`);
  }
  fs.mkdirSync(D, { recursive: true });

  // 1. Zetti
  let r;
  for (let i = 1; i <= 2; i++) {
    log(`${per}: 1/4 Zetti 5.6.5${i > 1 ? ' (reintento)' : ''}`);
    r = await P.correr(process.execPath, [path.join(SCRIPTS, 'bajar-ventas-zetti.js'), `--cuit=${CUIT}`, `--periodo=${per}`, `--dir=${D}`, '--si-existe=reemplazar'], { log, prefijo: '   ' });
    if (r.code === 0 && fs.existsSync(PA)) break;
    if (codeDe(r) === 'control') break; // el reporte no cierra: reintentar no cambia nada
  }
  if (!(r.code === 0 && fs.existsSync(PA))) return { ...fila, paso: 'zetti', error: errorDe(r), code: codeDe(r) };
  fila.zetti = r.json && { comprobantes: r.json.comprobantes, totales: r.json.totales, control_total_reporte: r.json.control_total_reporte,
    por_jurisdiccion: r.json.por_jurisdiccion, locales: r.json.locales, ...(r.json.aviso_nro ? { aviso_nro: r.json.aviso_nro } : {}) };

  // 2. Portal IVA, DDJJ presentada
  for (let i = 1; i <= 2; i++) {
    log(`${per}: 2/4 Libro IVA Ventas de la DDJJ presentada${i > 1 ? ' (reintento)' : ''}`);
    r = await P.correr(process.execPath, ['scripts/portal-iva-descarga.js', `--cuit=${CUIT}`, `--periodo=${P.mmaaaa(per)}`, '--presentada', '--libros=ventas', '--no-archivar'], { cwd: pPortal, log, prefijo: '   ' });
    if (r.code === 0) break;
    const c = codeDe(r);
    if (CORTES_ARCA.includes(c)) return { ...fila, paso: 'portal', error: errorDe(r), code: c, corte: true };
    if (c === 'sin_presentacion' || /no coincide con la DDJJ/i.test(errorDe(r))) break;
  }
  const libro = r.json && (r.json.libros || [])[0];
  if (r.code !== 0 || !libro || !libro.archivo) {
    return { ...fila, paso: 'portal', error: errorDe(r), code: codeDe(r),
      nota: 'Si el libro no coincide con la DDJJ presentada, no forzarlo: avisar con los dos importes. Ojo: el control compara el neto gravado de facturas y tiques sin restar NC.' };
  }
  const portal = path.join(D, `${per} - PORTAL IVA - VENTAS.xlsx`);
  fs.copyFileSync(libro.archivo, portal);
  fila.portal = { ddjj: r.json.ddjj, filas: libro.filas, control_ddjj: libro.control_ddjj };

  // 3. Mis Comprobantes emitidos
  for (let i = 1; i <= 2; i++) {
    log(`${per}: 3/4 Mis Comprobantes emitidos${i > 1 ? ' (reintento)' : ''}`);
    r = await P.correr(process.execPath, ['scripts/mis-comprobantes-emitidos.js', `--cuit=${CUIT}`, `--periodo=${P.mmaaaa(per)}`, '--no-archivar', '--si-existe=reemplazar'], { cwd: pMc, log, prefijo: '   ' });
    if (r.code === 0) break;
    const c = codeDe(r);
    if (CORTES_ARCA.includes(c)) return { ...fila, paso: 'mc', error: errorDe(r), code: c, corte: true };
  }
  if (r.code !== 0 || !r.json || !r.json.archivo) return { ...fila, paso: 'mc', error: (r.json && r.json.aviso) || errorDe(r), code: codeDe(r) };
  const mc = path.join(D, `${per} - MIS COMPROBANTES - EMITIDOS.xlsx`);
  fs.copyFileSync(r.json.archivo, mc);
  fila.mc = { filas: r.json.filas, aviso: r.json.aviso };

  // 4. Armar el papel (una apertura y un guardado)
  log(`${per}: 4/4 armando el papel (Portal, MC, depurado, cuadro resumen)`);
  r = await P.correr(P.PYTHON, [path.join(SCRIPTS, 'armar_papel.py'), PA, portal, mc, ...EXC], { cwd: SCRIPTS, log, prefijo: '   ' });
  if (r.code !== 0) return { ...fila, paso: 'armado', error: errorDe(r) };
  const pasos = r.stdout.split(/\r?\n/).filter((l) => l.startsWith('{')).map((l) => { try { return JSON.parse(l); } catch { return null; } }).filter(Boolean);
  const cuadro = pasos.find((p) => 'total_cuadro' in p);
  const dep = pasos.find((p) => 'fila_total_zetti' in p);
  fila.armado = {
    cuadro_ok: cuadro ? cuadro.ok : null,
    cuadro_diferencia: cuadro && cuadro.diferencia,
    codigos_no_reconocidos: cuadro && cuadro.no_reconocidos,
    depurado: dep && Object.fromEntries(Object.entries(dep).filter(([k]) => !/^fila_/.test(k))),
  };
  log(`${per}: OK`);
  return { ...fila, ok: true };
}

async function cruce(per) {
  const PA = papelDe(per);
  if (!fs.existsSync(PA)) return { periodo: per, ok: false, error: 'no hay papel' };
  if (P.abiertoEnExcel(PA)) return { periodo: per, ok: false, error: 'el papel esta abierto en Excel' };
  log(`${per}: cruce de ventas`);
  const r = await P.correr(P.PYTHON, [path.join(SCRIPTS, 'cruce_3_fuentes.py'), `--papel=${PA}`, ...EXC], { cwd: SCRIPTS, log, prefijo: '   ' });
  if (r.code !== 0 || !r.json || r.json.error) return { periodo: per, ok: false, error: errorDe(r) };
  const tmp = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'bva-cruce-')), 'cruce.json');
  fs.writeFileSync(tmp, JSON.stringify(r.json));
  const s = await P.correr(P.PYTHON, [path.join(SCRIPTS, 'resumen_cruce.py'), PA, tmp], { cwd: SCRIPTS });
  fs.rmSync(path.dirname(tmp), { recursive: true, force: true });
  return {
    periodo: per, ok: true,
    zetti_menos_portal: r.json.zetti_menos_portal, mc_menos_portal: r.json.mc_menos_portal,
    eliminados_zetti: r.json.eliminados_zetti, otros_periodos: r.json.otros_periodos,
    estados: s.json && s.json.estados, para_revisar: s.json && s.json.para_revisar, para_revisar_total: s.json && s.json.para_revisar_total,
    correlatividad: (r.json.correlatividad || []).length,
  };
}

(async () => {
  const PERS = periodos();
  if (!PERS.length) fail('No hay periodos para procesar');
  log(`${path.basename(socDir)} | ${PERS.join(', ')} | carpeta ${SUB ? `01-Impuestos Mensuales/${SUB}` : '01-Impuestos Mensuales (del mes)'}`);
  const salida = { ok: true, cuit: CUIT, sociedad: path.basename(socDir), farmacia: FARM, subcarpeta: SUB || null, meses: [], cruce: [] };
  let codigo = 0;

  if (!SOLO_CRUCE) {
    const pPortal = P.pluginDir('portal-iva-descarga');
    const pMc = P.pluginDir('cruce-ventas-3-fuentes');
    for (const per of PERS) {
      const t0 = Date.now();
      let r;
      try { r = await unMes(per, pPortal, pMc); } catch (e) { r = { periodo: per, ok: false, paso: 'error', error: e.message }; }
      r.minutos = Math.round((Date.now() - t0) / 6000) / 10;
      if (!r.ok) log(`${per}: FALLO ${r.paso}: ${r.error}`);
      salida.meses.push(r);
      if (r.corte) { log('corto: ARCA pidio captcha o rechazo la clave'); codigo = 3; break; }
    }
  }

  if (!SIN_CRUCE && codigo !== 3) {
    // El cruce mira el mes anterior y el siguiente: se corre con todos los meses ya armados.
    const aCruzar = SOLO_CRUCE ? PERS : PERS.filter((p) => (salida.meses.find((m) => m.periodo === p) || {}).ok);
    for (const per of aCruzar) salida.cruce.push(await cruce(per));
  }

  salida.ok = salida.meses.every((m) => m.ok) && salida.cruce.every((c) => c.ok);
  if (!salida.ok && codigo === 0) codigo = 1;
  fs.mkdirSync(SALIDAS, { recursive: true });
  const archivo = path.join(SALIDAS, `${CUIT}-${new Date().toISOString().replace(/[:T]/g, '-').slice(0, 19)}.json`);
  fs.writeFileSync(archivo, JSON.stringify(salida, null, 1));
  salida.resultado_en = archivo;
  console.log(JSON.stringify(salida, null, 1));
  process.exit(codigo);
})().catch((e) => fail(e.message, { code: e.code || 'error' }));
