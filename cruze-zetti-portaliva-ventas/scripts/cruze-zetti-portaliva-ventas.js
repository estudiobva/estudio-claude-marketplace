#!/usr/bin/env node
/**
 * cruze-zetti-portaliva-ventas.js — circuito completo de ventas del mes de una
 * sociedad: Portal IVA -> Zetti -> hojas del papel de trabajo -> cruze.
 *
 *   node scripts/cruze-zetti-portaliva-ventas.js --nombre="<sociedad>" --periodo=09/2026
 *   node scripts/cruze-zetti-portaliva-ventas.js --cuit=<CUIT>
 *   node scripts/cruze-zetti-portaliva-ventas.js --nombres="<A>;<B>" --periodo=09/2026
 *
 * Por sociedad:
 *   1. Portal IVA: AAAAMM - PORTAL IVA - VENTAS.xlsx de la carpeta IVA. Si no esta,
 *      o se bajo antes del dia 6 del mes siguiente (todavia sin todos los tiques), lo
 *      baja con portal-iva-descarga.js --libros=ventas. Si la DDJJ ya esta presentada,
 *      baja el libro presentado (--presentada).
 *   2. Zetti: AAAAMM - ZETTI - VENTAS.xlsx, mismo criterio, con zetti-iva-ventas.js.
 *   3. Papel de trabajo del mes (carpeta AAAAMM, "AAAAMM-papeldetrabajo-<SOC>.xlsx"):
 *      escribe las hojas "Ventas Zetti", "Ventas Portal IVA" y "cruze" con
 *      papel_ventas.py. Si el papel ya existe se respalda y se reescriben solo esas
 *      hojas; si alguna ya tiene datos corta salvo --si-existe=reemplazar.
 *   4. El cruce lo hace cruzar_ventas.py (las NC B del controlador fiscal van a la
 *      hoja "NC controlador (fuera)").
 *
 * Argumentos:
 *   --nombre / --cuit, --nombres / --cuits   sociedad(es), como en los otros scripts.
 *   --periodo       MM/AAAA o AAAAMM. Default: el mes anterior.
 *   --si-existe     abortar (default) | reemplazar: las hojas del papel que ya tienen datos.
 *   --rebajar       Baja Portal IVA y Zetti aunque ya esten en la carpeta.
 *   --papel=<ruta>  Papel de trabajo a usar, si en la carpeta del mes hay mas de uno.
 *
 * Salida: JSON por stdout; progreso por stderr. Navegador oculto (lo manejan los
 * scripts de descarga).
 */

const fs   = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const rutas = require('../lib/rutas-bva');

loadEnv();
const args = parseArgs();
const SI_EXISTE = String(args['si-existe'] || 'abortar').toLowerCase();
const REBAJAR = 'rebajar' in args;
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
if (!['abortar', 'reemplazar'].includes(SI_EXISTE)) fail('--si-existe tiene que ser abortar o reemplazar');

function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)))          return s;
  return null;
}
function mesAnterior(hoy = new Date()) {
  const d = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
}
const PERIODO = args['periodo'] ? normalizarPeriodo(args['periodo']) : mesAnterior();
if (!PERIODO) fail('--periodo tiene que ser MM/AAAA o AAAAMM');
const MM_AAAA = `${PERIODO.slice(4)}/${PERIODO.slice(0, 4)}`;
// Un libro bajado antes del dia 6 del mes siguiente puede no tener todos los
// tiques (mismo criterio que ncr-lote para dar un mes por cerrado).
const CORTE = new Date(Number(PERIODO.slice(0, 4)), Number(PERIODO.slice(4)), 6);

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length) fail('Falta --nombre, --cuit, --nombres o --cuits');
if (args['papel'] && PEDIDAS.length !== 1) fail('--papel es para una sola sociedad');

let prefijo = '';
const log = (m) => process.stderr.write(`[cruze-ventas]${prefijo} ${m}\n`);

function cuitDe(pedida) {
  if (pedida.cuit) return claves.normCuit(pedida.cuit);
  const { rows } = claves.loadSync();
  const { row, candidatos } = claves.buscar(rows, { nombre: pedida.nombre });
  if (candidatos) {
    throw new Error(`"${pedida.nombre}" coincide con ${candidatos.length} titulares: ` +
      candidatos.slice(0, 8).map((r) => `${r.titular} (${r.cuit})`).join(', ') + '. Afina el nombre o pasa --cuit.');
  }
  return row.cuit;
}

// Corre otro script y devuelve el JSON que imprime (el ultimo objeto de stdout).
function correr(cmd, argv, etiqueta) {
  log(`${etiqueta}: ${path.basename(argv[0])} ${argv.slice(1).join(' ')}`);
  const p = spawnSync(cmd, argv, { encoding: 'utf-8', maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'pipe'] });
  for (const l of String(p.stderr || '').split(/\r?\n/).filter(Boolean)) process.stderr.write(`    ${l}\n`);
  const out = String(p.stdout || '').trim();
  let json = null;
  try { json = JSON.parse(out); } catch {
    const i = out.lastIndexOf('\n{');
    try { json = JSON.parse(i >= 0 ? out.slice(i + 1) : out); } catch { /* no era JSON */ }
  }
  return { status: p.status, json, stdout: out };
}

const fresco = (f) => fs.existsSync(f) && fs.statSync(f).mtime >= CORTE;

function libroPortal(cuit, iva) {
  const archivo = path.join(iva, `${PERIODO} - PORTAL IVA - VENTAS.xlsx`);
  if (!REBAJAR && fresco(archivo)) return { archivo, bajado: false };
  const base = [path.join(__dirname, 'portal-iva-descarga.js'), `--cuit=${cuit}`, `--periodo=${MM_AAAA}`, '--libros=ventas', '--si-existe=reemplazar'];
  let r = correr(process.execPath, base, 'Portal IVA');
  let modo = 'borrador';
  if (r.json && r.json.code === 'ddjj_presentada') {
    modo = 'presentada';
    log('Portal IVA: la DDJJ ya esta presentada, bajo el libro presentado');
    r = correr(process.execPath, [...base, '--presentada'], 'Portal IVA');
  }
  if (!r.json || !r.json.ok) {
    const det = r.json ? (r.json.error || (r.json.libros || []).map((l) => l.error).filter(Boolean).join(' | ')) : r.stdout.slice(-300);
    throw Object.assign(new Error(`Portal IVA: ${det}`), { code: (r.json && r.json.code) || 'portal' });
  }
  return { archivo, bajado: true, modo };
}

function libroZetti(cuit, iva) {
  const archivo = path.join(iva, `${PERIODO} - ZETTI - VENTAS.xlsx`);
  if (!REBAJAR && fresco(archivo)) return { archivo, bajado: false };
  const r = correr(process.execPath, [path.join(__dirname, 'zetti-iva-ventas.js'), `--cuit=${cuit}`, `--periodo=${MM_AAAA}`, '--si-existe=reemplazar'], 'Zetti');
  if (!r.json || !r.json.ok) {
    throw Object.assign(new Error(`Zetti: ${r.json ? r.json.error : r.stdout.slice(-300)}`), { code: (r.json && r.json.code) || 'zetti' });
  }
  return { archivo, bajado: true, comprobantes: r.json.comprobantes };
}

// Papel de trabajo del mes: el .xlsx que diga "papel" en la carpeta AAAAMM (no
// los de agentes de recaudacion ni los marcados "NO VA"). Si no hay, uno nuevo
// "AAAAMM-papeldetrabajo-<SOC>.xlsx" (formato compacto del papel).
function papelDelMes(cuit) {
  if (args['papel']) return { papel: path.resolve(args['papel']), aviso: null };
  const soc = rutas.carpetaSociedad(cuit);
  const mes = rutas.carpetaMes(path.join(soc, '01-Impuestos Mensuales'), PERIODO);
  fs.mkdirSync(mes, { recursive: true });
  const todos = fs.readdirSync(mes).filter((f) => /papel/i.test(f) && !/^~\$/.test(f) && !/agente/i.test(f) && !/no va/i.test(f));
  const xlsx = todos.filter((f) => /\.xlsx$/i.test(f));
  if (xlsx.length > 1) {
    throw Object.assign(new Error(`Hay ${xlsx.length} papeles de trabajo en ${mes}: ${xlsx.join(' | ')}. Indica cual con --papel.`), { code: 'papel_ambiguo' });
  }
  if (xlsx.length === 1) return { papel: path.join(mes, xlsx[0]), aviso: null };
  const sigla = path.basename(soc).split(' - ').slice(1, -1).join(' ')
    .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  const nuevo = path.join(mes, `${PERIODO}-papeldetrabajo-${sigla}.xlsx`);
  const xls = todos.filter((f) => /\.xls$/i.test(f));
  return {
    papel: nuevo,
    aviso: xls.length ? `El papel del mes es .xls (${xls.join(', ')}) y no se puede escribir: las hojas quedan en ${path.basename(nuevo)}.` : null,
  };
}

function unaSociedad(pedida) {
  const fila = { sociedad: pedida.nombre || pedida.cuit, cuit: null, periodo: PERIODO, ok: false };
  try {
    const cuit = cuitDe(pedida);
    fila.cuit = cuit;
    const iva = rutas.carpetaIva(cuit, PERIODO);

    fila.portal = libroPortal(cuit, iva);
    log(`Portal IVA: ${fila.portal.bajado ? 'bajado' : 'ya estaba'} ${fila.portal.archivo}`);
    fila.zetti = libroZetti(cuit, iva);
    log(`Zetti: ${fila.zetti.bajado ? 'bajado' : 'ya estaba'} ${fila.zetti.archivo}`);

    const { papel, aviso } = papelDelMes(cuit);
    if (aviso) { fila.aviso = aviso; log(aviso); }
    const r = correr(PYTHON, [path.join(__dirname, 'papel_ventas.py'), `--zetti=${fila.zetti.archivo}`,
      `--portal=${fila.portal.archivo}`, `--papel=${papel}`, `--cuit=${cuit}`, `--periodo=${PERIODO}`,
      `--si-existe=${SI_EXISTE}`], 'Papel y cruze');
    if (!r.json || !r.json.ok) {
      throw Object.assign(new Error(r.json ? r.json.error : r.stdout.slice(-300)), { code: 'papel' });
    }
    const c = r.json.cruce;
    log(`papel ${r.json.papel_nuevo ? 'nuevo' : 'actualizado'}: ${r.json.papel}`);
    log(`cruze: ${c.filas_diferencia} filas con diferencia, Zetti ${c.total_zetti} vs Portal ${c.total_portal} (dif ${c.diferencia})`);
    Object.assign(fila, { ok: true, papel: r.json.papel, papel_nuevo: r.json.papel_nuevo, respaldo: r.json.respaldo,
      hojas: r.json.hojas, filas_portal: r.json.filas_portal, filas_zetti: r.json.filas_zetti, cruce: c });
  } catch (e) {
    log(`ERROR ${e.message}`);
    Object.assign(fila, { ok: false, code: e.code || 'error', error: e.message });
  }
  return fila;
}

const resultados = [];
for (const [i, pedida] of PEDIDAS.entries()) {
  prefijo = PEDIDAS.length > 1 ? ` [${i + 1}/${PEDIDAS.length}]` : '';
  resultados.push(unaSociedad(pedida));
}
prefijo = '';
const ok = resultados.every((r) => r.ok);
console.log(JSON.stringify(PEDIDAS.length === 1 ? resultados[0] : {
  ok, periodo: PERIODO, procesadas: resultados.length,
  con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, error: r.error })),
  resultados,
}, null, 2));
process.exit(ok ? 0 : 1);
