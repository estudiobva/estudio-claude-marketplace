#!/usr/bin/env node
/**
 * ncr-lote.js — para varias sociedades de un periodo, corre en fila los pasos
 * 1, 2, 3 y la SIMULACION del 4 del circuito NCR, y devuelve una tabla para aprobar:
 *
 *   compras del Portal IVA -> descargar-ncr.js -> conciliar-ncr.py -> eliminar-ncr.js (simulacion)
 *
 * Paso 1: primero busca las compras en la carpeta del periodo (misma regla que
 * conciliar-ncr.py). El equipo suele bajar el Portal IVA y liquidar IVA semanas
 * despues: si el archivo ya esta, se usa ese y no se entra a ARCA. Solo si falta
 * se baja con portal-iva-descarga.js (--libros=compras).
 * Antes de usar uno que ya estaba se mira que dia se bajo (fecha de modificacion
 * del archivo): bajado durante el mes del periodo, o hasta el dia 5 del mes
 * siguiente, es PARCIAL; bajado despues del dia 5 del mes siguiente es un mes
 * cerrado. Si es parcial y es el "AAAAMM - PORTAL IVA - COMPRAS.xlsx" del script,
 * se vuelve a bajar y se reemplaza; si es otro archivo, la sociedad queda con
 * error (no se borra nada del equipo).
 *
 *   node scripts/ncr-lote.js --periodo=08/2026 --nombres="<sociedad A>;<sociedad B>"
 *   node scripts/ncr-lote.js --periodo=202608 --cuits=<CUIT A>,<CUIT B>
 *   node scripts/ncr-lote.js --periodo=08/2026 --nombres="<sociedad>" --sin-simular
 *
 * Argumentos:
 *   --periodo      MM/AAAA o AAAAMM.
 *   --nombres      Titulares del Excel de claves, separados por ";".
 *   --cuits        CUITs separados por coma (se pueden combinar con --nombres).
 *   --bundle       Paquete del SharePoint (default: el que busca descargar-ncr,
 *                  ~/Downloads/SP-NCR-AAAAMM.bundle). Se arma una vez por mes con
 *                  scripts/sharepoint/bundle-ncr.js en la pestaña del SharePoint.
 *   --actualizar   Pasa --actualizar a descargar-ncr (reemplaza NCR que cambiaron).
 *   --reconciliar  Rehace la conciliacion aunque ya exista. Sin esto se rehace
 *                  solo si descargar-ncr trajo algo nuevo o reemplazado.
 *   --sin-simular  No simula en ARCA: solo compras, descarga y concilia.
 *   --no-bajar-compras  Si las compras no estan en la carpeta (o estan parciales),
 *                  no las baja de ARCA: la sociedad queda con error.
 *   --aceptar-compras-parciales  Concilia igual con un archivo parcial (queda aviso).
 *   --pausa=MS     Espera entre sociedades antes de entrar a ARCA. Default 5000.
 *   --ver          Navegador visible en la simulacion.
 *
 * NUNCA elimina: no acepta --ejecutar. La eliminacion se corre despues, de a una
 * sociedad y con confirmacion, con eliminar-ncr.js --ejecutar.
 *
 * Si ARCA rechaza una clave o pide captcha, no entra a ARCA con ninguna sociedad
 * mas (varias comparten apoderado y reintentar bloquea la clave); las que faltan
 * se descargan y concilian igual.
 *
 * Deja el resumen en ~/Documents/BVA-salidas/ncr-lote/ ($BVA_SALIDAS_PATH)AAAAMM-<fecha-hora>.json y .csv.
 */

const fs   = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const { parseArgs, fail } = require('../lib/args');
const { resolveCredentials } = require('../lib/env');
const rutas = require('../lib/rutas-bva');

const args = parseArgs();
if ('ejecutar' in args) fail('ncr-lote.js no elimina nada. Para eliminar: eliminar-ncr.js --ejecutar, de a una sociedad.');
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const PAUSA = args['pausa'] != null ? Number(args['pausa']) : 5000;
const log = (m) => process.stderr.write(`[ncr-lote] ${m}\n`);

function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)))          return s;
  return null;
}
const PERIODO = normalizarPeriodo(args['periodo']);
if (!PERIODO) fail('Falta --periodo (MM/AAAA o AAAAMM)');

// ── A quien ────────────────────────────────────────────────────────────────
const objetivos = [];
for (const n of String(args['nombres'] || '').split(';').map((x) => x.trim()).filter(Boolean)) {
  try {
    const { cuit, titular } = resolveCredentials({ nombre: n });
    objetivos.push({ pedido: n, cuit, titular: titular || n });
  } catch (e) { objetivos.push({ pedido: n, cuit: null, titular: n, error: e.message }); }
}
for (const c of String(args['cuits'] || '').split(',').map((x) => x.replace(/\D/g, '')).filter(Boolean)) {
  let titular = null;
  try { ({ titular } = resolveCredentials({ cuit: c })); } catch { /* el titular es solo para mostrar */ }
  objetivos.push({ pedido: c, cuit: c, titular: titular || c });
}
if (!objetivos.length) fail('Indica las sociedades con --nombres="A;B" y/o --cuits=X,Y');
const repetidos = objetivos.filter((o, i) => o.cuit && objetivos.findIndex((x) => x.cuit === o.cuit) !== i);
if (repetidos.length) fail(`Sociedades repetidas: ${repetidos.map((o) => o.pedido).join(', ')}`);

// ── Procesos hijos ─────────────────────────────────────────────────────────
// Corre un script, reenvia su stderr como progreso y devuelve el JSON de stdout.
function correr(cmd, argv, prefijo) {
  return new Promise((resolve) => {
    const env = { ...process.env, BVA_PYTHON: PYTHON };
    const ch = spawn(cmd, argv, { cwd: path.join(__dirname, '..'), env });
    let out = '', err = '';
    ch.stdout.on('data', (d) => { out += d; });
    ch.stderr.on('data', (d) => {
      err += d;
      for (const l of String(d).split(/\r?\n/).filter(Boolean)) process.stderr.write(`   ${prefijo} ${l}\n`);
    });
    ch.on('close', (code) => {
      let json = null;
      try { json = JSON.parse(out); }
      catch { const m = out.match(/\{[\s\S]*\}\s*$/); if (m) { try { json = JSON.parse(m[0]); } catch { /* sin JSON */ } } }
      resolve({ code, json, stderr: err });
    });
    ch.on('error', (e) => resolve({ code: -1, json: { error: `No pude correr ${cmd}: ${e.message}` }, stderr: '' }));
  });
}
const node = (script, argv) => correr(process.execPath, [path.join(__dirname, script), ...argv], `[${script.replace('.js', '')}]`);
const py   = (script, argv) => correr(PYTHON, [path.join(__dirname, script), ...argv], `[${script.replace('.py', '')}]`);
const errorDe = (r) => (r.json && r.json.error) || (r.code !== 0 ? `salio con codigo ${r.code}: ${r.stderr.trim().split('\n').pop() || ''}` : null);
const dormir = (ms) => new Promise((r) => setTimeout(r, ms));

// La DDJJ de IVA del periodo ya se presento si en la carpeta del mes (o una
// subcarpeta) hay un acuse "AAAAMM-IVA-ACUSE-...". eliminar-ncr entra por
// "Nueva declaracion jurada": con el periodo presentado eso puede abrir una
// rectificativa, asi que el lote no entra a ARCA con esa sociedad.
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

// ── Compras del Portal IVA ya bajadas ───────────────────────────────────────
// Misma regla que es_compras() / carpetas_compras() de conciliar-ncr.py, para
// que el archivo que se encuentra aca sea el mismo que despues usa la conciliacion.
const normNombre = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();

function esCompras(nombre) {
  const n = normNombre(nombre);
  if (nombre.startsWith('~$') || n.includes('SIN-NCR') || n.includes('SIN NCR') || n.includes('ELIMINACION')) return false;
  const ext = path.extname(nombre).toLowerCase();
  if (['.csv', '.zip'].includes(ext) && n.includes(`COMPROBANTES_PERIODO_${PERIODO}_COMPRAS`)) return true;
  return ['.csv', '.zip', '.xlsx'].includes(ext) && n.includes('COMPRAS') && n.replace(/ /g, '').includes('PORTAL');
}

// Devuelve { archivo, modificado, dias } del archivo de compras, o null si no hay.
// Si hay mas de uno en la misma carpeta, corta: conciliar-ncr.py tambien cortaria.
function comprasEnCarpeta(cuit) {
  const mens = path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales');
  const mes = rutas.carpetaMes(mens, PERIODO);
  const dirs = [];
  try { const ncr = rutas.carpetaNcr(cuit, PERIODO, { crear: false }); dirs.push(ncr, path.dirname(ncr)); } catch { /* sin ncr/ todavia */ }
  dirs.push(path.join(mes, `${PERIODO}-IVA`), path.join(mens, `${PERIODO}-IVA`), path.join(mes, 'IVA'), mes);
  for (const d of [...new Set(dirs)].filter((x) => fs.existsSync(x))) {
    const hits = fs.readdirSync(d).filter((f) => fs.statSync(path.join(d, f)).isFile() && esCompras(f)).sort();
    if (hits.length > 1) throw new Error(`Hay ${hits.length} archivos de compras en ${d}: ${hits.join(' | ')}. Deja uno solo.`);
    if (hits.length === 1) {
      const archivo = path.join(d, hits[0]);
      const modificado = fs.statSync(archivo).mtime;
      return { archivo, modificado, dias: Math.floor((Date.now() - modificado) / 86400000) };
    }
  }
  return null;
}

// Mes cerrado = bajado despues del dia 5 del mes siguiente al periodo. Antes de
// eso (durante el mes, o del 1 al 5 del siguiente) todavia puede faltar algo.
const DIA_CIERRE = 5;
function limiteCierre() {
  const anio = Number(PERIODO.slice(0, 4)), mes = Number(PERIODO.slice(4, 6));   // mes 1-12
  return new Date(anio, mes, DIA_CIERRE, 23, 59, 59, 999);                        // dia 5 del mes siguiente
}

const fechaCorta = (d) => `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;

// ── Una sociedad ───────────────────────────────────────────────────────────
let arcaBloqueado = null;
let primeraEnArca = true;

async function procesar(o) {
  const fila = { sociedad: o.titular, cuit: o.cuit, pasos: {} };
  if (o.error) return { ...fila, error: o.error };
  const base = [`--cuit=${o.cuit}`, `--periodo=${PERIODO}`];
  fila.avisos = [];

  // 1. Compras del Portal IVA: si ya estan en la carpeta y cubren el mes, se usan esas.
  let compras;
  try { compras = comprasEnCarpeta(o.cuit); } catch (e) { return { ...fila, error: `compras: ${e.message}` }; }
  let bajar = !compras, reemplazar = false, porQue = 'no estan en la carpeta';
  if (compras) {
    const nombre = path.basename(compras.archivo);
    fila.compras_origen = `ya estaba en la carpeta (bajada el ${fechaCorta(compras.modificado)}, hace ${compras.dias} dias)`;
    log(`   compras: ${nombre} · ${fila.compras_origen}`);
    if (compras.modificado <= limiteCierre()) {
      const motivo = `${nombre} es una descarga parcial: se bajo el ${fechaCorta(compras.modificado)}, ` +
        `antes del ${fechaCorta(limiteCierre())} (cierre del periodo)`;
      log(`   compras: PARCIAL · bajada el ${fechaCorta(compras.modificado)}`);
      if ('aceptar-compras-parciales' in args) {
        fila.avisos.push(`${motivo}. Se concilio igual (--aceptar-compras-parciales).`);
      } else if (nombre !== rutas.nombreArchivo(PERIODO, 'compras', 'xlsx')) {
        return { ...fila, error: `compras: ${motivo}. No es el archivo del script y no lo toco: ` +
          'borralo o renombralo (sacale "PORTAL" o "COMPRAS" del nombre) para que se vuelva a bajar, o --aceptar-compras-parciales.' };
      } else {
        bajar = true; reemplazar = true; porQue = 'la descarga que estaba es parcial';
      }
    }
  }
  if (bajar) {
    const acuse = acuseIva(o.cuit);
    if (acuse) return { ...fila, error: `compras: ${porQue} y la DDJJ de IVA ya esta presentada (${acuse}): no entro a ARCA a bajarlas.` };
    if ('no-bajar-compras' in args) return { ...fila, error: `compras: ${porQue} (--no-bajar-compras).` };
    if (arcaBloqueado) return { ...fila, error: `compras: ${porQue} y no se entra a ARCA: ${arcaBloqueado}` };
    if (!primeraEnArca && PAUSA > 0) await dormir(PAUSA);
    primeraEnArca = false;
    log(`   compras: ${porQue}, las bajo del Portal IVA${reemplazar ? ' y reemplazo el archivo' : ''}`);
    const b = await node('portal-iva-descarga.js', [...base, '--libros=compras',
      ...(reemplazar ? ['--si-existe=reemplazar'] : []), ...('ver' in args ? ['--ver'] : [])]);
    fila.pasos.compras = b.json;
    const eB = errorDe(b) || (b.json && !b.json.ok ? (b.json.libros || []).map((l) => l.error).filter(Boolean).join(' | ') || 'no se pudo bajar' : null);
    if (eB) {
      const code = b.json && b.json.code;
      if (code === 'captcha' || code === 'credenciales_invalidas') arcaBloqueado = `se corto el lote en ${o.titular} (${eB})`;
      return { ...fila, error: `compras (portal-iva-descarga): ${eB}` };
    }
    fila.compras_origen = reemplazar ? 'bajada ahora del Portal IVA (la anterior era parcial)' : 'bajada ahora del Portal IVA';
    if (new Date() <= limiteCierre()) {
      fila.avisos.push(`Las compras se bajaron hoy, antes del ${fechaCorta(limiteCierre())}: todavia no es el mes cerrado.`);
    }
  }

  // 2. NCR del paquete del SharePoint
  const d = await node('descargar-ncr.js', [...base,
    ...(args['bundle'] ? [`--bundle=${args['bundle']}`] : []), ...('actualizar' in args ? ['--actualizar'] : [])]);
  const eD = errorDe(d);
  fila.pasos.descargar = d.json;
  if (eD) return { ...fila, error: `descargar-ncr: ${eD}` };
  const estados = (d.json.archivos || []).map((a) => a.estado);
  const cambio = estados.some((e) => e === 'nuevo' || e === 'reemplazado');
  fila.monroe = (d.json.archivos || []).filter((a) => a.leidas != null)
    .map((a) => `${a.archivo.replace(/\.xlsx$/i, '')} ${a.leidas}/${a.declaradas}${a.pie_ok ? '' : ' TRUNCADA'}`);
  if (estados.includes('distinto_no_tocado')) {
    fila.avisos.push('El SharePoint tiene NCR distintas a las de ncr/ y no se copiaron: correr con --actualizar.');
  }

  // 3. Conciliacion
  const si = 'reconciliar' in args || cambio ? 'reemplazar' : 'saltear';
  let c = await py('conciliar-ncr.py', [...base, `--si-existe=${si}`]);
  if (!errorDe(c) && c.json.salteada) {
    fila.conciliacion_existente = true;
    c = { ...(await py('conciliar-ncr.py', [...base, '--solo-revisar'])), salida: c.json.salida };
    if (!errorDe(c)) c.json.salida = c.salida;
  }
  const eC = errorDe(c);
  fila.pasos.conciliar = c.json;
  if (eC) return { ...fila, error: `conciliar-ncr: ${eC}` };
  const k = c.json;
  Object.assign(fila, {
    compras: k.portal_comprobantes, compras_total: k.portal_total,
    ncr_monroe: k.ncr_monroe, ncr_suizo: k.ncr_suizo,
    a_eliminar: k.eliminados, a_eliminar_total: k.eliminados_total,
    compras_sin_ncr: k.resultantes, compras_sin_ncr_total: k.resultantes_total,
    ncr_no_encontradas: k.ncr_no_encontradas, diferencias: k.diferencias,
    sin_respaldo: k.nc_sin_respaldo, sin_respaldo_total: k.nc_sin_respaldo_total,
    salida: k.salida, compras_archivo: k.compras,
  });
  fila.avisos.push(...(k.avisos || []));
  if (!k.eliminados) { fila.simulacion = 'nada para eliminar'; return fila; }

  // 4. Simulacion en ARCA (no elimina)
  if ('sin-simular' in args) { fila.simulacion = 'no pedida'; return fila; }
  if (arcaBloqueado) { fila.simulacion = `no se corrio: ${arcaBloqueado}`; return fila; }
  const acuse = acuseIva(o.cuit);
  if (acuse) { fila.simulacion = `no se corrio: la DDJJ de IVA ya esta presentada (${acuse})`; return fila; }
  if (!primeraEnArca && PAUSA > 0) await dormir(PAUSA);
  primeraEnArca = false;
  const s = await node('eliminar-ncr.js', [...base, `--archivo=${k.salida}`, ...('ver' in args ? ['--ver'] : [])]);
  const eS = errorDe(s);
  fila.pasos.simular = s.json;
  if (eS) {
    const code = s.json && s.json.code;
    if (code === 'captcha' || /rechaz\w* las credenciales/i.test(eS)) {
      arcaBloqueado = `se corto el lote en ${o.titular} (${eS})`;
    }
    return { ...fila, error: `eliminar-ncr (simulacion): ${eS}` };
  }
  const pa = s.json.por_accion || {};
  fila.sim_eliminar = pa.eliminar || 0;
  fila.sim_no_esta = pa.no_esta_en_libro || 0;
  fila.sim_diferencia = pa.diferencia || 0;
  fila.sim_ambigua = pa.ambigua || 0;
  const lib = s.stderr.match(/libro:\s*(\d+)\s*comprobantes\s*·\s*total\s*(-?[\d.]+)/);
  if (lib) {
    fila.libro = Number(lib[1]); fila.libro_total = Number(lib[2]);
    if (fila.libro !== fila.compras || Math.abs(fila.libro_total - fila.compras_total) > 0.01) {
      fila.avisos.push(`El libro en ARCA (${fila.libro} por ${fila.libro_total}) no coincide con el archivo de compras ` +
        `(${fila.compras} por ${fila.compras_total}): se reimporto o ya se eliminaron NCR.`);
    }
  }
  fila.simulacion = (fila.sim_diferencia || fila.sim_ambigua) ? 'REVISAR: hay diferencias o ambiguas'
    : (fila.sim_eliminar === fila.a_eliminar ? 'OK' : `OK parcial: ${fila.sim_no_esta} ya no estan en el libro`);
  fila.log_simulacion = s.json.log;
  return fila;
}

// ── Principal ──────────────────────────────────────────────────────────────
(async () => {
  log(`periodo ${PERIODO} · ${objetivos.length} sociedades · ${'sin-simular' in args ? 'sin simulacion' : 'con simulacion en ARCA'}`);
  const filas = [];
  for (const [i, o] of objetivos.entries()) {
    log(`(${i + 1}/${objetivos.length}) ${o.titular}`);
    let f;
    try { f = await procesar(o); } catch (e) { f = { sociedad: o.titular, cuit: o.cuit, error: e.message }; }
    filas.push(f);
    log(`   -> ${f.error ? 'ERROR ' + f.error : `a eliminar ${f.a_eliminar} · simulacion ${f.simulacion}`}`);
  }

  const d = new Date(), z = (n) => String(n).padStart(2, '0');
  const sello = `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`;
  const dir = path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'ncr-lote');
  fs.mkdirSync(dir, { recursive: true });
  const base = path.join(dir, `${PERIODO}-${sello}`);
  fs.writeFileSync(`${base}.json`, JSON.stringify({ periodo: PERIODO, generado: d.toISOString(), filas }, null, 1), 'utf-8');

  const cols = ['sociedad', 'cuit', 'compras_origen', 'compras', 'compras_total', 'ncr_monroe', 'ncr_suizo', 'a_eliminar', 'a_eliminar_total',
    'compras_sin_ncr', 'compras_sin_ncr_total', 'ncr_no_encontradas', 'diferencias', 'sin_respaldo', 'sin_respaldo_total',
    'libro', 'sim_eliminar', 'sim_no_esta', 'sim_diferencia', 'sim_ambigua', 'simulacion', 'monroe', 'avisos', 'error'];
  const celda = (v) => {
    const s = Array.isArray(v) ? v.join(' | ') : (v == null ? '' : (typeof v === 'number' ? String(v).replace('.', ',') : String(v)));
    return /[;"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  fs.writeFileSync(`${base}.csv`, '﻿' + [cols.join(';'), ...filas.map((f) => cols.map((c) => celda(f[c])).join(';'))].join('\r\n'), 'utf-8');

  const tabla = filas.map((f) => ({
    sociedad: f.sociedad, cuit: f.cuit, compras_origen: f.compras_origen, compras: f.compras, a_eliminar: f.a_eliminar, a_eliminar_total: f.a_eliminar_total,
    compras_sin_ncr_total: f.compras_sin_ncr_total, no_encontradas: f.ncr_no_encontradas, diferencias: f.diferencias,
    sin_respaldo: f.sin_respaldo, sin_respaldo_total: f.sin_respaldo_total, monroe: f.monroe,
    simulacion: f.simulacion, avisos: f.avisos, error: f.error,
    comando_para_eliminar: f.simulacion === 'OK' || /^OK parcial/.test(f.simulacion || '')
      ? `node scripts/eliminar-ncr.js --cuit=${f.cuit} --periodo=${PERIODO} --archivo="${f.salida}" --ejecutar` : null,
  }));
  console.log(JSON.stringify({
    ok: filas.every((f) => !f.error), periodo: PERIODO,
    listas_para_eliminar: tabla.filter((t) => t.comando_para_eliminar).length,
    con_error: filas.filter((f) => f.error).length,
    arca_cortado: arcaBloqueado, tabla, resumen: `${base}.json`, csv: `${base}.csv`,
  }, null, 2));
  if (filas.some((f) => f.error)) process.exitCode = 1;
})();
