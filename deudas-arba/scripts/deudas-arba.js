#!/usr/bin/env node
/**
 * deudas-arba.js — el equivalente en ARBA de deudas-ddjj-sct: para cada
 * sociedad del archivo de claves releva la deuda (IIBB y agente de
 * recaudación), las DDJJ que ARBA da por no presentadas, las notificaciones
 * DFE sin leer y el estado del agente, y arma UN Excel consolidado.
 *
 *   node scripts/deudas-arba.js                      (todas las que tienen CIT de ARBA)
 *   node scripts/deudas-arba.js --cuit=<CUIT>         (una)
 *   node scripts/deudas-arba.js --solo=7,12,30       (posiciones del Excel de claves)
 *   node scripts/deudas-arba.js --consolidar         (solo rearma el Excel con lo de hoy)
 *
 * Argumentos:
 *   --anios=5            Años hacia atrás para el detalle por período (default 5).
 *   --detalle-agentes    Cruza declarado vs pagado de TODOS los agentes. Por
 *                        defecto solo se hace cuando ARBA marca deuda, DDJJ sin
 *                        presentar o no pudo calcular la situación del agente.
 *   --rehacer            Vuelve a relevar las ya hechas hoy (por defecto se saltean).
 *   --salida             Carpeta base (default: $BVA_SALIDAS_PATH/deudas-arba,
 *                        o ~/Documents/BVA-salidas/deudas-arba).
 *   --ver                Navegador visible.
 *
 * Un login por sociedad (cada una tiene su CIT). Si ARBA rechaza la clave NO se
 * reintenta: 13 CIT de la planilla ya venían rechazadas en septiembre y
 * reintentarlas puede bloquearlas. Las sociedades sin CIT igual salen en el
 * Excel con el estado público del agente.
 *
 * Esto solo LEE ARBA. Nunca genera boletas, planes, liquidaciones ni DDJJ.
 */

const fs   = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const { chromium, request: pwRequest } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const arba = require('../lib/arba-deudas');

const args = parseArgs();
const VER = 'ver' in args;
const REHACER = 'rehacer' in args;
const SOLO_CONSOLIDAR = 'consolidar' in args;
const DETALLE_TODOS = 'detalle-agentes' in args;
const ANIOS = Math.max(1, Number(args['anios'] || 5));
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const log = (m) => process.stderr.write(`[arba] ${m}\n`);

const HOY = new Date();
const iso = [HOY.getFullYear(), String(HOY.getMonth() + 1).padStart(2, '0'),
             String(HOY.getDate()).padStart(2, '0')].join('-');
const DESDE_ANIO = HOY.getFullYear() - ANIOS + 1;
loadEnv(); // antes de BASE: BVA_SALIDAS_PATH puede venir del .env
const BASE = args['salida'] || path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'deudas-arba');
const DIR = path.join(BASE, iso);
const CRUDO = path.join(DIR, '_crudo');
const RESUMEN = path.join(DIR, 'lote-resumen.json');
// CIT que ARBA ya rechazó, con una huella de la clave: mientras la planilla
// tenga la misma clave no se vuelve a intentar (cada intento fallido suma para
// el bloqueo de la CIT, que el equipo usa a mano).
const RECHAZADAS = path.join(BASE, 'cit-rechazadas.json');
const huella = (clave) => crypto.createHash('sha256').update(String(clave)).digest('hex').slice(0, 16);

// ── A quién relevar ─────────────────────────────────────────────────────────
let rows;
try { ({ rows } = claves.loadSync()); }
catch (e) { fail(`No pude leer el archivo de claves: ${e.message}`); }

const todos = [];
const vistos = new Set();
for (const r of rows) {
  if (!r.cuit || vistos.has(r.cuit)) continue;
  vistos.add(r.cuit);
  const c = claves.credenciales(r, 'ARBA');
  todos.push({ pos: r.pos, titular: r.titular, cuit: r.cuit,
               usuario: claves.normCuit(c && c.usuario) || r.cuit, clave: c && c.clave ? c.clave : null });
}
const solo = args['solo'] ? new Set(String(args['solo']).split(',').map(x => x.trim())) : null;
const objetivos = SOLO_CONSOLIDAR ? [] : todos.filter(o =>
  args['cuit'] ? o.cuit === String(args['cuit']).replace(/\D/g, '') : (!solo || solo.has(String(o.pos).trim())));
if (!objetivos.length && !SOLO_CONSOLIDAR) fail('No hay ninguna sociedad para relevar con ese filtro.');

// ── Estado (reanudable) ─────────────────────────────────────────────────────
fs.mkdirSync(CRUDO, { recursive: true });
const estado = fs.existsSync(RESUMEN) ? JSON.parse(fs.readFileSync(RESUMEN, 'utf-8')) : { fecha: iso, sociedades: {} };
const guardar = () => fs.writeFileSync(RESUMEN, JSON.stringify(estado, null, 1), 'utf-8');
function registrar(o, datos) {
  estado.sociedades[o.cuit] = { pos: o.pos, sociedad: o.titular, cuit: o.cuit, ...datos, cuando: new Date().toISOString() };
  guardar();
}
const yaHecha = (o) => !REHACER && estado.sociedades[o.cuit] && estado.sociedades[o.cuit].ok;
const rechazadas = fs.existsSync(RECHAZADAS) ? JSON.parse(fs.readFileSync(RECHAZADAS, 'utf-8')) : {};
const guardarRechazo = (o, msg) => {
  rechazadas[o.cuit] = { huella: huella(o.clave), fecha: iso, sociedad: o.titular, mensaje: msg };
  fs.writeFileSync(RECHAZADAS, JSON.stringify(rechazadas, null, 1), 'utf-8');
};
const yaRechazada = (o) => o.clave && rechazadas[o.cuit] && rechazadas[o.cuit].huella === huella(o.clave);

// ── Una sociedad ────────────────────────────────────────────────────────────
async function relevar(browser, req, o) {
  const agenteEstado = await arba.estadoAgente(req, o.cuit, HOY);
  if (!o.clave) {
    return { ok: true, sinCit: true, agenteEstado, objetos: [], alertas: [], periodos: [], faltantes: [], iniciadas: [] };
  }
  const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1366, height: 900 } });
  try {
    const cred = { usuario: o.usuario, clave: o.clave };
    const sit = await arba.leerSituacion(context, cred);
    const r = { ok: true, razonSocial: sit.razonSocial, agenteEstado, objetos: sit.objetos, alertas: sit.alertas,
                periodos: [], faltantes: [], iniciadas: [], observaciones: [] };

    // Agentes: detalle por período si hay algo que explicar (o si se pidió para todos).
    const agentes = sit.objetos.filter(x => x.codImpuesto === 7);
    const conAlgo = agentes.filter(x => x.error || x.deudaTotal > 0 || x.ddjjSinPresentar > 0 || !x.consultada);
    if (agentes.length && (DETALLE_TODOS || conAlgo.length)) {
      const conSinDdjj = agentes.filter(x => x.ddjjSinPresentar > 0);
      const det = await arba.saldosAgente(context, cred, {
        desdeAnio: DESDE_ANIO, hoy: HOY,
        codigos: [...new Set(agentes.map(x => x.actividad))],
        activos: conSinDdjj.filter(x => !/CESADO/.test(x.vinculo)).map(x => x.actividad),
        cesados: conSinDdjj.filter(x => /CESADO/.test(x.vinculo)).map(x => x.actividad),
      });
      r.periodos = det.periodos;
      r.faltantes.push(...det.faltantes);
      r.iniciadas = det.iniciadas;
      r.observaciones.push(...det.errores);
    }
    // IIBB: meses sin DDJJ cuando ARBA marca alguna, o cuando ARBA no pudo
    // calcular la situación (salvo Convenio Multilateral: esas DDJJ van por SIFERE).
    const iibb = sit.objetos.find(x => x.codImpuesto === 2);
    const cm = iibb && /CONVENIO/i.test(iibb.vinculo);
    if (iibb && (iibb.ddjjSinPresentar > 0 || (iibb.error && !cm))) {
      const vd = iibb.vigenteDesde && /^\d{8}$/.test(iibb.vigenteDesde)
        ? new Date(Number(iibb.vigenteDesde.slice(0, 4)), Number(iibb.vigenteDesde.slice(4, 6)) - 1, 1) : null;
      const desde = new Date(Math.max(new Date(DESDE_ANIO, 0, 1), vd || 0));
      r.faltantes.push(...await arba.faltantesIibb(context, cred, { desde, hoy: HOY }));
    }
    if (iibb && iibb.error && cm) {
      r.observaciones.push('IIBB en Convenio Multilateral: ARBA no calcula la deuda acá; revisar en SIFERE');
    } else if (iibb && iibb.error && /IMPUESTO INVALIDO/i.test(iibb.error)) {
      r.observaciones.push('ARBA no pudo calcular la deuda de IIBB ("No se pudo calcular la situación del objeto"); las DDJJ de anticipo se revisaron en IBPresentaciones');
    }
    if (sit.objetos.some(x => /proceso de actualizaci/i.test(x.error || ''))) {
      r.observaciones.push('Cuenta en proceso de actualización en ARBA: volver a correr mañana');
    }
    fs.writeFileSync(path.join(CRUDO, `arba-${o.cuit}.json`), JSON.stringify(r, null, 1), 'utf-8');
    return r;
  } finally {
    await context.close().catch(() => {});
  }
}

const SIN_REINTENTO = new Set(['arba_login_rechazado']);

(async () => {
  const total = objetivos.length;
  log(`${total} sociedad(es) · detalle desde ${DESDE_ANIO} · salida: ${DIR}`);
  if (total) {
    const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 150 } : {}));
    const req = await pwRequest.newContext({ extraHTTPHeaders: { 'User-Agent': 'Mozilla/5.0' } });
    try {
      let n = 0;
      for (const o of objetivos) {
        n++;
        if (yaHecha(o)) continue;
        if (yaRechazada(o)) {
          const r0 = rechazadas[o.cuit];
          log(`[${n}/${total}] #${o.pos} ${o.titular}: CIT rechazada el ${r0.fecha} con esta misma clave — no se reintenta`);
          registrar(o, { ok: false, code: 'arba_login_rechazado',
                         error: `ARBA rechazó esta CIT el ${r0.fecha} y la planilla tiene la misma clave: actualizar Claves_Organismos (no se reintentó para no bloquearla)`,
                         agenteEstado: await arba.estadoAgente(req, o.cuit, HOY) });
          continue;
        }
        log(`[${n}/${total}] #${o.pos} ${o.titular} (${o.cuit})${o.clave ? '' : ' — sin CIT: solo estado de agente'}`);
        let ultimo = null;
        for (let intento = 1; intento <= 2; intento++) {
          try {
            const r = await relevar(browser, req, o);
            registrar(o, r);
            const deuda = (r.objetos || []).reduce((a, x) => a + (x.deudaTotal || 0), 0);
            const sinDdjj = (r.objetos || []).reduce((a, x) => a + (x.ddjjSinPresentar || 0), 0);
            log(`   deuda $ ${deuda.toLocaleString('es-AR', { minimumFractionDigits: 2 })} · DDJJ sin presentar ${sinDdjj}` +
                ` · ${r.periodos.length} período(s) con saldo · agente: ${r.agenteEstado.estado || r.agenteEstado.error || '-'}`);
            ultimo = null;
            break;
          } catch (e) {
            ultimo = e;
            if (!browser.isConnected()) throw Object.assign(new Error('Se cerró el navegador. Volvé a correrlo y retoma donde quedó.'), { code: 'navegador_cerrado' });
            if (SIN_REINTENTO.has(e.code) || intento === 2) break;
            log(`   falló (${e.code || 'error'}): ${e.message.split('\n')[0]} — reintento`);
          }
        }
        if (ultimo) {
          const msg = ultimo.message.split('\n')[0];
          log(`   ERROR (${ultimo.code || 'error'}): ${msg}`);
          if (ultimo.code === 'arba_login_rechazado') guardarRechazo(o, msg);
          const agenteEstado = await arba.estadoAgente(req, o.cuit, HOY);
          registrar(o, { ok: false, error: msg, code: ultimo.code || 'error', agenteEstado });
        }
      }
    } finally {
      await req.dispose().catch(() => {});
      await browser.close().catch(() => {});
    }
  }

  // ── Consolidado ─────────────────────────────────────────────────────────
  const orden = todos.map(o => estado.sociedades[o.cuit]).filter(Boolean);
  const consolidado = path.join(DIR, `Deudas_DDJJ_ARBA_CONSOLIDADO_${iso}.xlsx`);
  const tmp = path.join(CRUDO, 'lote.json');
  fs.writeFileSync(tmp, JSON.stringify({ fecha: iso, desdeAnio: DESDE_ANIO, sociedades: orden }), 'utf-8');
  let excelError = null;
  try {
    execFileSync(PYTHON, [path.join(__dirname, 'deudas_arba_a_excel.py'), tmp, consolidado],
                 { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
  } catch (e) {
    excelError = String(e.stderr || e.message).trim().split(/\r?\n/).pop();
    log(`no pude armar el consolidado: ${excelError}`);
  }

  const ok = orden.filter(s => s.ok);
  const deuda = (s) => (s.objetos || []).reduce((a, x) => a + (x.deudaTotal || 0), 0);
  const sinDdjj = (s) => (s.objetos || []).reduce((a, x) => a + (x.ddjjSinPresentar || 0), 0);
  console.log(JSON.stringify({
    ok: !excelError,
    fecha: iso,
    consolidado: excelError ? null : consolidado,
    relevadas: ok.filter(s => !s.sinCit).length,
    sinCit: ok.filter(s => s.sinCit).map(s => s.sociedad),
    conDeuda: ok.filter(s => deuda(s) > 0).map(s => ({ sociedad: s.sociedad, cuit: s.cuit, deuda: Math.round(deuda(s) * 100) / 100 })),
    conDdjjSinPresentar: ok.filter(s => sinDdjj(s) > 0).map(s => ({ sociedad: s.sociedad, cuit: s.cuit, cantidad: sinDdjj(s) })),
    agentesSuspendidos: orden.filter(s => s.agenteEstado && s.agenteEstado.suspendido).map(s => s.sociedad),
    errores: orden.filter(s => !s.ok).map(s => ({ pos: s.pos, sociedad: s.sociedad, cuit: s.cuit, code: s.code, error: s.error })),
  }, null, 2));
})().catch(e => fail(e.message, { code: e.code || 'error' }));
