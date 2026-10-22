#!/usr/bin/env node
/**
 * cruce-ventas-3-fuentes.js — ventas del mes de una sociedad desde tres fuentes
 * y su cruce:
 *
 *   1. Mis Comprobantes (ARCA), emitidos -> mis-comprobantes-emitidos.js
 *   2. Portal IVA, Libro IVA Ventas      -> portal-iva-descarga.js --libros=ventas
 *   3. Zetti, Subdiario IVA Ventas 5.6.5 -> zetti-iva-ventas.js
 *   4. Cruce de las tres                 -> cruzar_ventas_3.py
 *
 *   node scripts/cruce-ventas-3-fuentes.js --nombre="<sociedad>" --periodo=09/2026
 *   node scripts/cruce-ventas-3-fuentes.js --cuit=<CUIT>
 *   node scripts/cruce-ventas-3-fuentes.js --nombres="<A>;<B>" --periodo=09/2026
 *
 * Todo queda en la carpeta IVA del mes:
 *   AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx
 *   AAAAMM - PORTAL IVA - VENTAS.xlsx
 *   AAAAMM - ZETTI - VENTAS.xlsx
 *   AAAAMM - CRUCE VENTAS MC PORTAL ZETTI.xlsx   (Resumen / Diferencias / Detalle)
 *
 * Cada fuente se reusa si ya esta en la carpeta y se bajo a partir del dia 6 del
 * mes siguiente (antes puede faltar algo); si no, se baja. Si la DDJJ de IVA ya
 * esta presentada, el Portal baja el libro presentado (--presentada). Si una
 * fuente no se puede bajar (por ejemplo, la sociedad no usa Zetti) se cruzan
 * las otras dos y se avisa.
 *
 * Argumentos:
 *   --nombre / --cuit, --nombres / --cuits   sociedad(es).
 *   --periodo       MM/AAAA o AAAAMM. Default: el mes anterior.
 *   --rebajar       Baja las tres fuentes aunque ya esten en la carpeta.
 *   --si-existe     abortar | reemplazar (default): el archivo del cruce.
 *   --tolerancia    por componente, default 0.5.
 *
 * No toca el papel de trabajo (eso es cruze-zetti-portaliva-ventas.js). Solo
 * lectura en ARCA y en Zetti. Navegador oculto.
 * Salida: JSON por stdout; progreso por stderr.
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
const REBAJAR = 'rebajar' in args;
const SI_EXISTE = String(args['si-existe'] || 'reemplazar').toLowerCase();
const TOLERANCIA = args['tolerancia'] || '0.5';
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
// Mismo criterio que ncr-lote y cruze-zetti-portaliva-ventas para dar un mes por cerrado.
const CORTE = new Date(Number(PERIODO.slice(0, 4)), Number(PERIODO.slice(4)), 6);

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length) fail('Falta --nombre, --cuit, --nombres o --cuits');

let prefijo = '';
const log = (m) => process.stderr.write(`[ventas-3]${prefijo} ${m}\n`);

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
const errorDe = (r) => (r.json ? (r.json.error || (r.json.libros || []).map((l) => l.error).filter(Boolean).join(' | ')) : r.stdout.slice(-300)) || 'error sin detalle';

function fuenteMc(cuit, iva) {
  const archivo = path.join(iva, `${PERIODO} - MIS COMPROBANTES - EMITIDOS.xlsx`);
  if (!REBAJAR && fresco(archivo)) return { archivo, bajado: false };
  const r = correr(process.execPath, [path.join(__dirname, 'mis-comprobantes-emitidos.js'), `--cuit=${cuit}`, `--periodo=${MM_AAAA}`, '--si-existe=reemplazar'], 'Mis Comprobantes');
  if (!r.json || !r.json.ok) throw Object.assign(new Error(`Mis Comprobantes: ${errorDe(r)}`), { code: (r.json && r.json.code) || 'mis_comprobantes' });
  if (!r.json.archivo) return { archivo: null, bajado: true, aviso: r.json.aviso };
  return { archivo, bajado: true, comprobantes: r.json.comprobantes, aviso: r.json.aviso };
}

function fuentePortal(cuit, iva) {
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
  if (!r.json || !r.json.ok) throw Object.assign(new Error(`Portal IVA: ${errorDe(r)}`), { code: (r.json && r.json.code) || 'portal' });
  return { archivo, bajado: true, modo };
}

function fuenteZetti(cuit, iva) {
  const archivo = path.join(iva, `${PERIODO} - ZETTI - VENTAS.xlsx`);
  if (!REBAJAR && fresco(archivo)) return { archivo, bajado: false };
  const r = correr(process.execPath, [path.join(__dirname, 'zetti-iva-ventas.js'), `--cuit=${cuit}`, `--periodo=${MM_AAAA}`, '--si-existe=reemplazar'], 'Zetti');
  if (!r.json || !r.json.ok) throw Object.assign(new Error(`Zetti: ${errorDe(r)}`), { code: (r.json && r.json.code) || 'zetti' });
  return { archivo, bajado: true, comprobantes: r.json.comprobantes };
}

function unaSociedad(pedida) {
  const fila = { sociedad: pedida.nombre || pedida.cuit, cuit: null, periodo: PERIODO, ok: false, fuentes: {}, errores: [] };
  try {
    const cuit = cuitDe(pedida);
    fila.cuit = cuit;
    const iva = rutas.carpetaIva(cuit, PERIODO);

    for (const [clave, fn, nombre] of [['mc', fuenteMc, 'Mis Comprobantes'], ['portal', fuentePortal, 'Portal IVA'], ['zetti', fuenteZetti, 'Zetti']]) {
      try {
        const f = fn(cuit, iva);
        fila.fuentes[clave] = f;
        log(`${nombre}: ${f.archivo ? (f.bajado ? 'bajado' : 'ya estaba') + ' ' + path.basename(f.archivo) : 'sin comprobantes'}`);
        if (f.aviso) log(`${nombre}: ${f.aviso}`);
      } catch (e) {
        fila.errores.push({ fuente: clave, code: e.code, error: e.message });
        log(`${nombre}: ERROR ${e.message}`);
        // con la clave de ARCA bloqueada no tiene sentido seguir intentando
        if (['captcha', 'credenciales_invalidas'].includes(e.code)) throw e;
      }
    }

    const disponibles = Object.entries(fila.fuentes).filter(([, f]) => f.archivo);
    if (disponibles.length < 2) {
      throw Object.assign(new Error(`Solo hay ${disponibles.length} fuente(s) para cruzar: ` +
        fila.errores.map((e) => e.error).join(' | ')), { code: 'fuentes_insuficientes' });
    }
    const argv = [path.join(__dirname, 'cruzar_ventas_3.py'), `--cuit=${cuit}`, `--periodo=${PERIODO}`,
      `--tolerancia=${TOLERANCIA}`, `--si-existe=${SI_EXISTE}`, ...disponibles.map(([k, f]) => `--${k}=${f.archivo}`)];
    if (pedida.nombre) argv.push(`--sociedad=${pedida.nombre}`);
    const r = correr(PYTHON, argv, 'Cruce');
    if (!r.json || !r.json.ok) throw Object.assign(new Error(`Cruce: ${errorDe(r)}`), { code: (r.json && r.json.code) || 'cruce' });
    const c = r.json;
    log(`cruce: ${c.renglones_con_diferencia} renglones con diferencia de ${c.renglones} -> ${c.salida}`);
    for (const [k, t] of Object.entries(c.totales)) log(`  total ${k}: ${t.total}`);
    Object.assign(fila, {
      ok: true, completo: fila.errores.length === 0, cruce: c.salida, totales: c.totales,
      nc_controlador_fuera: c.nc_controlador_fuera, renglones: c.renglones,
      renglones_con_diferencia: c.renglones_con_diferencia, por_estado: c.por_estado, avisos: c.avisos,
      diferencias: c.diferencias,
    });
  } catch (e) {
    log(`ERROR ${e.message}`);
    Object.assign(fila, { ok: false, code: e.code || 'error', error: e.message });
  }
  return fila;
}

const resultados = [];
for (const [i, pedida] of PEDIDAS.entries()) {
  prefijo = PEDIDAS.length > 1 ? ` [${i + 1}/${PEDIDAS.length}]` : '';
  const r = unaSociedad(pedida);
  resultados.push(r);
  if (['captcha', 'credenciales_invalidas'].includes(r.code)) { log('corto el lote: ARCA pidio captcha o rechazo la clave'); break; }
}
prefijo = '';
const ok = resultados.length === PEDIDAS.length && resultados.every((r) => r.ok);
console.log(JSON.stringify(PEDIDAS.length === 1 ? resultados[0] : {
  ok, periodo: PERIODO, procesadas: resultados.length,
  con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, error: r.error })),
  incompletas: resultados.filter((r) => r.ok && !r.completo).map((r) => ({ sociedad: r.sociedad, errores: r.errores })),
  resultados,
}, null, 2));
process.exit(ok ? 0 : 1);
