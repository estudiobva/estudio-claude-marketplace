#!/usr/bin/env node
/**
 * eliminar-ncr.js — saca del borrador del Libro Compras (Portal IVA) las Notas
 * de Credito de Recupero que la skill conciliacion-ncr dejo en la hoja
 * "Eliminados". Es el paso manual que mas tiempo le lleva al estudio.
 *
 *   node scripts/eliminar-ncr.js --nombre="<sociedad>" --periodo=08/2026            (simulacion)
 *   node scripts/eliminar-ncr.js --nombre="<sociedad>" --periodo=08/2026 --ejecutar --limite=5
 *   node scripts/eliminar-ncr.js --nombre="<sociedad>" --periodo=08/2026 --ejecutar
 *
 * Argumentos:
 *   --nombre / --cuit   Sociedad (igual que portal-iva.js).
 *   --periodo           MM/AAAA o AAAAMM.
 *   --archivo           Salida de conciliacion-ncr. Default: la unica
 *                       [AAAAMM]-COMPRAS-SIN-NCR-*.xlsx de la carpeta ncr/ del periodo.
 *   --ejecutar          Elimina de verdad. SIN este flag solo simula: entra,
 *                       cruza cada NCR contra el libro y dice que haria.
 *   --limite=N          Elimina solo las primeras N (para pilotos).
 *   --pausa=MS          Espera entre eliminaciones. Default 400.
 *   --ver               Navegador visible.
 *
 * Como elimina: igual que el tachito rojo de cada fila. Ese boton abre un
 * "¿Eliminar el comprobante…?" y, al aceptar, manda ajax.do?f=eliminarComprobante&id=<idReg>.
 * Aca se manda ese mismo pedido, con el idReg de la fila que matcheo. Nunca se
 * usa "Eliminar todos" (eliminarComprobantesCategoria) ni se presenta nada.
 *
 * Una NCR se elimina solo si en el libro hay EXACTAMENTE UNA fila con el mismo
 * tipo (3), CUIT emisor, punto de venta y numero, y ademas coinciden importe
 * (al centavo) y fecha. Cualquier otra cosa se informa y no se toca.
 *
 * Es reanudable: si se corta, se vuelve a correr y las ya eliminadas aparecen
 * como "no esta en el libro".
 */

const fs   = require('fs');
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
const EJECUTAR = 'ejecutar' in args;
const LIMITE = args['limite'] != null ? Number(args['limite']) : Infinity;
const PAUSA  = args['pausa'] != null ? Number(args['pausa']) : 400;
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
if ('ver' in args && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';
if (!(LIMITE > 0)) fail('--limite tiene que ser un numero mayor a 0');

const log = (m) => process.stderr.write(`[eliminar-ncr] ${m}\n`);

function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)))          return s;
  return null;
}
const PERIODO = normalizarPeriodo(args['periodo']);
if (!PERIODO) fail('Falta --periodo (MM/AAAA o AAAAMM)');

let CUIT, CUIT_LOGIN, PASSWORD, TITULAR;
try {
  ({ cuit: CUIT, cuitLogin: CUIT_LOGIN, password: PASSWORD, titular: TITULAR } =
    resolveCredentials({ cuit: args['cuit'], nombre: args['nombre'] }));
} catch (e) {
  fail(e.message, { sugerencia: `Revisa el .env (${envFile() || 'no encontrado'}) o el Excel de claves.` });
}

// ── 1. La conciliacion ─────────────────────────────────────────────────────
function ubicarConciliacion() {
  if (args['archivo']) return path.resolve(String(args['archivo']));
  const ncr = rutas.carpetaNcr(CUIT, PERIODO, { crear: false });
  if (!fs.existsSync(ncr)) throw new Error(`No existe la carpeta ${ncr}. Primero hay que correr conciliacion-ncr.`);
  const re = new RegExp(`^${PERIODO}-COMPRAS-SIN-NCR-.*\\.xlsx$`, 'i');
  const hits = fs.readdirSync(ncr).filter((f) => re.test(f) && !f.startsWith('~$'));
  if (hits.length === 0) throw new Error(`No hay ningun ${PERIODO}-COMPRAS-SIN-NCR-*.xlsx en ${ncr}. Corre primero conciliacion-ncr.`);
  if (hits.length > 1) throw new Error(`Hay ${hits.length} salidas de conciliacion en ${ncr} (${hits.join(' | ')}). Indica cual con --archivo.`);
  return path.join(ncr, hits[0]);
}

let ARCHIVO, CONC;
try {
  ARCHIVO = ubicarConciliacion();
  CONC = JSON.parse(execFileSync(PYTHON, [path.join(__dirname, 'leer_conciliacion_ncr.py'), ARCHIVO],
                                 { encoding: 'utf-8', maxBuffer: 64 * 1024 * 1024 }));
} catch (e) {
  fail(`No pude leer la conciliacion: ${(e.stderr || e.message).toString().trim().split('\n').pop()}`);
}
if (CONC.errores.length) fail('La hoja Eliminados tiene problemas; no elimino nada.', { errores: CONC.errores });
if (!CONC.items.length) fail('La hoja Eliminados esta vacia: no hay nada para eliminar.');

// ── helpers de pagina ─────────────────────────────────────────────────────
const TIPOS_NC = new Set([3, 8, 13, 21, 38, 43, 48, 53, 90, 110, 112, 113, 114, 119, 203, 208, 213]);
const clave = (cuit, tipo, pv, nro) => `${cuit}|${Number(tipo)}|${Number(pv)}|${Number(nro)}`;
const isoDesdeDMY = (s) => { const m = String(s).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? `${m[3]}-${m[2]}-${m[1]}` : null; };

// Lee TODAS las filas del libro desde el DataTable (gServerSide=false: estan en memoria).
// Indices de la fila, sacados de compras.js de ARCA (v=20260826):
//   0 fecha dd/mm/aaaa · 1 tipo · 3 PV · 4 numero · 11 CUIT emisor · 12 denominacion · 23 total · 25 idReg
// Como ARCA puede reordenarlos, antes se valida contra lo que la grilla MUESTRA.
async function leerLibro(page) {
  // La grilla tiene que estar dibujada (sirve para validar indices contra lo visible).
  await page.waitForFunction(() => window.jQuery && jQuery.fn.dataTable &&
    jQuery.fn.dataTable.isDataTable('#tablaDataTables') &&
    (document.querySelector('#tablaDataTables tbody tr .botonEliminarComprobante') ||
     /Ning.n dato|No hay/i.test(document.body.innerText)),
  null, { timeout: 180000, polling: 1000 });

  return page.evaluate(async () => {
    const dt = jQuery('#tablaDataTables').DataTable();
    const get = (data) => new Promise((res, rej) => jQuery.ajax({ url: 'ajax.do', data, dataType: 'json',
      timeout: 120000, success: res, error: (x) => rej(new Error(`HTTP ${x.status} leyendo el libro`)) }));

    // Libros chicos: gServerSide=false y todas las filas ya estan en memoria.
    // Libros grandes (~5.600 comprobantes): gServerSide=true y la grilla pide
    // de a una pagina a ajax.do?f=listaComprobantesIncluidosSS. Se pide igual, por tandas.
    let rows, totalDeclarado;
    if (!window.gServerSide) {
      rows = dt.rows().data().toArray();
      totalDeclarado = rows.length;
    } else {
      rows = [];
      const TANDA = 500;
      for (let draw = 1; ; draw++) {
        const r = await get({ f: 'listaComprobantesIncluidosSS', c: window.gTipoOperacion, draw,
                              start: rows.length, length: TANDA,
                              'search[value]': '', 'search[regex]': false,
                              'order[0][column]': 0, 'order[0][dir]': 'asc' });
        const d = (r && r.datos) || {};
        const data = d.data || [];
        totalDeclarado = Number(d.recordsTotal ?? r.recordsTotal ?? d.recordsFiltered ?? r.recordsFiltered ?? NaN);
        rows.push(...data);
        if (!data.length || rows.length >= totalDeclarado || data.length < TANDA) break;
        if (draw > 200) throw new Error('Demasiadas paginas leyendo el libro.');
      }
    }

    // El total que ARCA muestra en pantalla ("de un total de N") es la referencia.
    const enPantalla = document.body.innerText.match(/de un total de ([\d.]+)/i);
    if (enPantalla) totalDeclarado = Number(enPantalla[1].replace(/\./g, ''));

    let validacion = 'sin filas';
    const tr = document.querySelector('#tablaDataTables tbody tr');
    const btn = tr && tr.querySelector('.botonEliminarComprobante');
    if (btn) {
      const r = dt.row(tr).data();
      const nroVisible = tr.children[2].innerText.replace(/\D/g, '');
      const ok = String(r[25]) === String(btn.getAttribute('data-id-reg')) &&
                 Number(nroVisible) === Number(String(r[3]) + String(r[4]).padStart(8, '0')) &&
                 /^\d{11}$/.test(String(r[11])) && /^\d{2}\/\d{2}\/\d{4}$/.test(String(r[0]));
      validacion = ok ? 'ok' : `no coincide: fila=${JSON.stringify(r).slice(0, 300)} nroVisible=${nroVisible}`;
    }
    const ids = new Set(rows.map((r) => String(r[25])));
    if (ids.size !== rows.length) validacion = `idReg repetidos en la lectura (${rows.length} filas, ${ids.size} ids)`;
    if (rows.length !== totalDeclarado) validacion = `lei ${rows.length} filas y ARCA declara ${totalDeclarado}`;
    return {
      validacion, serverSide: !!window.gServerSide,
      estado: window.gEstadoPresentacion, cuit: window.gUserEmpresaRepresentadaCUIT,
      mesInicio: window.gConfigFechas && window.gConfigFechas.fechaMesInicio,
      filas: rows.map((r) => ({ fecha: r[0], tipo: Number(r[1]), pv: Number(r[3]), nro: Number(r[4]),
                                cuit: String(r[11]), nombre: r[12], total: Number(r[23]), idReg: String(r[25]) })),
    };
  });
}

function totalConSigno(filas) {
  return Math.round(filas.reduce((s, f) => s + (TIPOS_NC.has(f.tipo) ? -Math.abs(f.total) : f.total), 0) * 100) / 100;
}

// ── principal ──────────────────────────────────────────────────────────────
(async () => {
  const browser = await chromium.launch(launchOptions());
  const context = await browser.newContext({ locale: 'es-AR' });
  const page = await context.newPage();
  const resultados = [];
  let resumen = null;

  try {
    log(`${TITULAR || CUIT} · CUIT ${CUIT} · periodo ${PERIODO} · ${EJECUTAR ? 'MODO EJECUCION' : 'SIMULACION (no elimina nada)'}`);
    log(`conciliacion: ${path.basename(ARCHIVO)} · ${CONC.items.length} NCR para eliminar`);

    await login(page, { cuitLogin: CUIT_LOGIN, password: PASSWORD });
    const iva = await pi.abrirPortalIva(context, page);
    const repr = await pi.asegurarRepresentacion(iva, CUIT);
    await pi.abrirBorrador(iva, PERIODO);
    if (/verDatosInicialesPresentacion/i.test(iva.url())) {
      throw new Error('El borrador del periodo no esta iniciado (ARCA pide datos iniciales): no hay compras importadas para depurar.');
    }
    await pi.abrirLibro(iva, 'compras');
    await pi.cerrarModales(iva);

    // ── 2. Controles antes de tocar nada ──
    const antes = await leerLibro(iva);
    const mmAAAA = `${PERIODO.slice(4, 6)}/${PERIODO.slice(0, 4)}`;
    if (antes.estado !== 'BO') throw new Error(`El libro no esta en Borrador (estado ${antes.estado}). No se puede eliminar.`);
    if (String(antes.cuit) !== String(CUIT)) throw new Error(`El libro abierto es del CUIT ${antes.cuit}, no de ${CUIT}.`);
    if (!String(antes.mesInicio || '').endsWith(mmAAAA)) throw new Error(`El libro abierto es de ${antes.mesInicio}, no de ${mmAAAA}.`);
    if (antes.validacion !== 'ok') throw new Error(`La estructura de la grilla de ARCA cambio (${antes.validacion}). Hay que revisar los indices.`);
    log(`libro: ${antes.filas.length} comprobantes · total ${totalConSigno(antes.filas)}${antes.serverSide ? ' (leido por paginas)' : ''}`);

    // ── 3. Cruce NCR -> fila del libro ──
    const indice = new Map();
    for (const f of antes.filas) {
      const k = clave(f.cuit, f.tipo, f.pv, f.nro);
      indice.set(k, [...(indice.get(k) || []), f]);
    }
    for (const it of CONC.items) {
      const hits = indice.get(clave(it.cuit, it.tipo, it.pv, it.nro)) || [];
      const r = { ...it, idReg: null, libro_total: null, libro_fecha: null, accion: null, detalle: '' };
      if (hits.length === 0) { r.accion = 'no_esta_en_libro'; r.detalle = 'ya eliminada o nunca importada'; }
      else if (hits.length > 1) { r.accion = 'ambigua'; r.detalle = `${hits.length} filas con la misma clave`; }
      else {
        const f = hits[0];
        Object.assign(r, { idReg: f.idReg, libro_total: f.total, libro_fecha: isoDesdeDMY(f.fecha) });
        const difImp = Math.abs(Math.abs(f.total) - Math.abs(it.importe)) > 0.01;
        const difFecha = r.libro_fecha !== it.fecha;
        if (difImp || difFecha) {
          r.accion = 'diferencia';
          r.detalle = [difImp && `importe libro ${f.total} vs conciliacion ${it.importe}`,
                       difFecha && `fecha libro ${r.libro_fecha} vs conciliacion ${it.fecha}`].filter(Boolean).join('; ');
        } else r.accion = 'eliminar';
      }
      resultados.push(r);
    }
    const aEliminar = resultados.filter((r) => r.accion === 'eliminar').slice(0, LIMITE);
    const cuenta = (a) => resultados.filter((r) => r.accion === a).length;
    log(`cruce: eliminar=${cuenta('eliminar')} no_esta=${cuenta('no_esta_en_libro')} diferencia=${cuenta('diferencia')} ambigua=${cuenta('ambigua')}`);

    // ── 4. Eliminacion, de a una ──
    let eliminadas = 0;
    if (EJECUTAR) {
      if (LIMITE !== Infinity) log(`limite: solo las primeras ${aEliminar.length}`);
      for (const [i, r] of aEliminar.entries()) {
        const resp = await iva.evaluate((id) => new Promise((res) => {
          jQuery.ajax({ url: 'ajax.do', data: { f: 'eliminarComprobante', id }, dataType: 'json', timeout: 30000,
            success: (x) => res({ ok: x && x.estado === 'ok', x: JSON.stringify(x).slice(0, 200) }),
            error: (xhr, st) => res({ ok: false, x: `HTTP ${xhr.status} ${st}` }) });
        }), r.idReg);
        if (!resp.ok) {
          r.accion = 'error'; r.detalle = resp.x;
          log(`ERROR en ${r.drogueria} ${r.pv}-${r.nro}: ${resp.x}. Corto aca; las anteriores quedaron eliminadas.`);
          break;
        }
        r.accion = 'eliminada'; eliminadas++;
        if ((i + 1) % 25 === 0 || i + 1 === aEliminar.length) log(`eliminadas ${i + 1}/${aEliminar.length}`);
        await iva.waitForTimeout(PAUSA);
      }

      // ── 5. Verificacion: se recarga el libro desde ARCA, no se confia en la memoria ──
      await iva.goto(iva.url().replace(/verCompras\.do.*/, 'verCompras.do?t=21'), { waitUntil: 'domcontentloaded' });
      await pi.cerrarModales(iva);
      const despues = await leerLibro(iva);
      const idsDespues = new Set(despues.filas.map((f) => f.idReg));
      const siguen = resultados.filter((r) => r.accion === 'eliminada' && idsDespues.has(r.idReg));
      for (const r of siguen) { r.accion = 'error'; r.detalle = 'ARCA dijo ok pero sigue en el libro'; }
      resumen = {
        antes: { comprobantes: antes.filas.length, total: totalConSigno(antes.filas) },
        despues: { comprobantes: despues.filas.length, total: totalConSigno(despues.filas) },
        eliminadas, siguenEnLibro: siguen.length,
        cuadraCantidad: antes.filas.length - eliminadas === despues.filas.length,
        // Informativo: solo cuadra si el libro no cambio desde que se bajo el CSV de la conciliacion.
        cuadraConConciliacion: CONC.compras_sin_ncr
          ? despues.filas.length === CONC.compras_sin_ncr.comprobantes &&
            Math.abs(totalConSigno(despues.filas) - CONC.compras_sin_ncr.total) <= 0.01
          : null,
      };
      log(`verificacion: ${resumen.antes.comprobantes} -> ${resumen.despues.comprobantes} comprobantes ` +
          `(${eliminadas} eliminadas) · ${resumen.cuadraCantidad && !siguen.length ? 'OK' : 'NO CUADRA'}`);
    }

    // ── 6. Log: archivo nuevo, nunca se pisa ──
    const d = new Date(), z = (n) => String(n).padStart(2, '0');
    const sello = `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`;
    const nombreLog = `${PERIODO}-ELIMINACION-NCR-${sello}${EJECUTAR ? '' : '-SIMULACION'}.csv`;
    const dirLog = EJECUTAR ? path.dirname(ARCHIVO) : path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'ncr-simulaciones');
    fs.mkdirSync(dirLog, { recursive: true });
    const cols = ['fila', 'drogueria', 'farmacia', 'cuit', 'tipo', 'pv', 'nro', 'importe', 'fecha',
                  'idReg', 'libro_total', 'libro_fecha', 'accion', 'detalle'];
    const csv = [cols.join(';'), ...resultados.map((r) => cols.map((c) => String(r[c] ?? '').replace(/;/g, ',')).join(';'))].join('\r\n');
    const rutaLog = path.join(dirLog, nombreLog);
    if (fs.existsSync(rutaLog)) throw new Error(`Ya existe ${rutaLog}; no lo piso.`);
    fs.writeFileSync(rutaLog, '﻿' + csv, 'utf-8');

    const porAccion = {};
    for (const r of resultados) porAccion[r.accion] = (porAccion[r.accion] || 0) + 1;
    console.log(JSON.stringify({
      ok: !resultados.some((r) => r.accion === 'error') && (!resumen || resumen.cuadraCantidad),
      modo: EJECUTAR ? 'ejecucion' : 'simulacion', titular: TITULAR || null, cuit: CUIT, periodo: PERIODO,
      representando: repr.linea, conciliacion: ARCHIVO, ncr_en_conciliacion: CONC.items.length,
      por_accion: porAccion, verificacion: resumen, compras_sin_ncr_esperado: CONC.compras_sin_ncr,
      log: rutaLog,
    }, null, 2));
  } catch (e) {
    if (e.code === 'captcha') fail('ARCA pidio un captcha.', { code: 'captcha' });
    if (e.code === 'credenciales_invalidas') fail(`ARCA rechazo las credenciales del CUIT ${CUIT}. No reintentes: bloquea la clave.`);
    fail(e.message, { code: e.code || 'error', procesadas: resultados.filter((r) => r.accion === 'eliminada').length });
  } finally {
    await browser.close().catch(() => {});
  }
})();
