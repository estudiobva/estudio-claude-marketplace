#!/usr/bin/env node
/**
 * dfe-arba.js — revisa el Domicilio Fiscal Electrónico (DFE) de ARBA de las
 * sociedades: entra con la CIT de cada una (planilla Claves_Organismos), lee
 * la bandeja completa (intimaciones y notificaciones, leídas o no) y arma UN
 * Excel consolidado. Antes cruza el Maestro de Sociedades contra la planilla de
 * claves para que se vea a quién no se pudo revisar y por qué.
 *
 *   node scripts/dfe-arba.js                      (todas las que tienen CIT de ARBA)
 *   node scripts/dfe-arba.js --cuit=<CUIT>         (una)
 *   node scripts/dfe-arba.js --solo=3,7,12        (posiciones del Excel de claves)
 *   node scripts/dfe-arba.js --consolidar         (solo rearma el Excel con lo de hoy)
 *
 * Argumentos:
 *   --dias=30     Ventana para "recibidas en los últimos N días" (default 30).
 *   --rehacer     Vuelve a relevar las ya hechas hoy (por defecto se saltean).
 *   --salida      Carpeta base (default: $BVA_SALIDAS_PATH/dfe-arba, o ~/Documents/BVA-salidas/dfe-arba).
 *   --ver         Navegador visible (solo si se pide).
 *
 * Solo LEE el listado de la bandeja. NUNCA abre "Visualizar" (eso marca la
 * comunicación como leída y puede adelantar la notificación).
 *
 * Un login por sociedad (cada una tiene su CIT). Si ARBA rechaza la clave NO se
 * reintenta, y tampoco se prueban las CIT que deudas-arba ya registró como
 * rechazadas con la misma clave (cada intento suma para el bloqueo).
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const rutas = require('../lib/rutas-bva');
const { readWorkbook } = require('../lib/xlsx-min');
const dfe = require('../lib/arba-dfe');

const args = parseArgs();
const VER = 'ver' in args;
const REHACER = 'rehacer' in args;
const SOLO_CONSOLIDAR = 'consolidar' in args;
const DIAS = Math.max(1, Number(args['dias'] || 30));
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const log = (m) => process.stderr.write(`[dfe-arba] ${m}\n`);

loadEnv();
const HOY = new Date();
const iso = [HOY.getFullYear(), String(HOY.getMonth() + 1).padStart(2, '0'),
             String(HOY.getDate()).padStart(2, '0')].join('-');
const SALIDAS = process.env.BVA_SALIDAS_PATH && process.env.BVA_SALIDAS_PATH.trim();
const BASE = args['salida'] || path.join(SALIDAS || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'dfe-arba');
const DIR = path.join(BASE, iso);
const CRUDO = path.join(DIR, '_crudo');
const RESUMEN = path.join(DIR, 'lote-resumen.json');
const RECHAZADAS = path.join(BASE, 'cit-rechazadas.json');
// Las de deudas-arba (mismo login, misma CIT): se consultan, no se escriben.
const RECHAZADAS_OTRAS = [
  SALIDAS && path.join(SALIDAS, 'deudas-arba', 'cit-rechazadas.json'),
  path.join(os.homedir(), 'Documents', 'BVA-salidas', 'deudas-arba', 'cit-rechazadas.json'),
  path.join(__dirname, '..', 'descargas', 'deudas-arba', 'cit-rechazadas.json'),
].filter(Boolean);
const huella = (clave) => crypto.createHash('sha256').update(String(clave)).digest('hex').slice(0, 16);
// Intimaciones y notificaciones (las "Comunicaciones" son avisos que no se notifican).
const formal = (c) => !/^Comunicaci/i.test(c.tipo || '');
const leerJson = (p) => { try { return JSON.parse(fs.readFileSync(p, 'utf-8')); } catch { return {}; } };

// ── Planilla de claves ──────────────────────────────────────────────────────
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

// ── Maestro de Sociedades ───────────────────────────────────────────────────
function leerMaestro() {
  const archivo = path.join(rutas.raiz(), '0 - ESTUDIO - 00000000000', 'Maestro de Sociedades.xlsx');
  const filas = readWorkbook(fs.readFileSync(archivo)).rows('MAESTRO');
  const iEnc = filas.findIndex(f => f.some(c => /^CUIT$/i.test(String(c).trim())));
  if (iEnc < 0) throw new Error('No encuentro la fila de encabezados en la hoja MAESTRO.');
  const enc = filas[iEnc].map(c => String(c).trim());
  const col = (re) => enc.findIndex(c => re.test(c));
  const cNom = col(/^NOMBRE$/i), cCuit = col(/^CUIT$/i), cTipo = col(/^TIPO$/i);
  const cIibb = col(/^IIBB/i), cArba = col(/^ARBA$/i), cMuni = col(/^MUNICIPIO$/i);
  return filas.slice(iEnc + 1)
    .filter(f => String(f[cCuit] || '').replace(/\D/g, '').length === 11)
    .map(f => ({
      nombre: String(f[cNom] || '').trim(), cuit: String(f[cCuit]).replace(/\D/g, ''),
      tipo: String(f[cTipo] || '').trim(), iibb: String(f[cIibb] || '').trim(),
      arba: cArba >= 0 ? String(f[cArba] || '').trim() : '', municipio: cMuni >= 0 ? String(f[cMuni] || '').trim() : '',
    }));
}
let maestro = [];
let maestroError = null;
try { maestro = leerMaestro(); } catch (e) { maestroError = e.message; log(`no pude leer el Maestro: ${e.message}`); }

const solo = args['solo'] ? new Set(String(args['solo']).split(',').map(x => x.trim())) : null;
const objetivos = SOLO_CONSOLIDAR ? [] : todos.filter(o => o.clave && (
  args['cuit'] ? o.cuit === String(args['cuit']).replace(/\D/g, '') : (!solo || solo.has(String(o.pos).trim()))));
if (!objetivos.length && !SOLO_CONSOLIDAR) fail('No hay ninguna sociedad con CIT de ARBA para ese filtro.');

// ── Estado (reanudable) ─────────────────────────────────────────────────────
fs.mkdirSync(CRUDO, { recursive: true });
const estado = fs.existsSync(RESUMEN) ? leerJson(RESUMEN) : { fecha: iso, sociedades: {} };
estado.sociedades = estado.sociedades || {};
const guardar = () => fs.writeFileSync(RESUMEN, JSON.stringify(estado, null, 1), 'utf-8');
function registrar(o, datos) {
  estado.sociedades[o.cuit] = { pos: o.pos, sociedad: o.titular, cuit: o.cuit, ...datos, cuando: new Date().toISOString() };
  guardar();
}
const yaHecha = (o) => !REHACER && estado.sociedades[o.cuit] && estado.sociedades[o.cuit].ok;
const rechazadas = leerJson(RECHAZADAS);
const otrasRechazadas = Object.assign({}, ...RECHAZADAS_OTRAS.map(leerJson));
const guardarRechazo = (o, msg) => {
  rechazadas[o.cuit] = { huella: huella(o.clave), fecha: iso, sociedad: o.titular, mensaje: msg };
  fs.writeFileSync(RECHAZADAS, JSON.stringify(rechazadas, null, 1), 'utf-8');
};
const rechazoPrevio = (o) => [rechazadas[o.cuit], otrasRechazadas[o.cuit]]
  .find(r => r && o.clave && r.huella === huella(o.clave)) || null;

(async () => {
  const total = objetivos.length;
  log(`${total} sociedad(es) · salida: ${DIR}`);
  if (total) {
    const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 150 } : {}));
    try {
      let n = 0;
      for (const o of objetivos) {
        n++;
        if (yaHecha(o)) continue;
        const previo = rechazoPrevio(o);
        if (previo) {
          log(`[${n}/${total}] #${o.pos} ${o.titular}: CIT rechazada el ${previo.fecha} con esta misma clave — no se reintenta`);
          registrar(o, { ok: false, code: 'arba_login_rechazado',
                         error: `ARBA rechazó esta CIT el ${previo.fecha} y la planilla tiene la misma clave: actualizar Claves_Organismos (no se reintentó para no bloquearla)` });
          continue;
        }
        log(`[${n}/${total}] #${o.pos} ${o.titular} (${o.cuit})`);
        let ultimo = null;
        for (let intento = 1; intento <= 2; intento++) {
          const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1366, height: 900 } });
          try {
            const r = await dfe.leerBandeja(context, { usuario: o.usuario, clave: o.clave });
            fs.writeFileSync(path.join(CRUDO, `dfe-${o.cuit}.json`), JSON.stringify(r, null, 1), 'utf-8');
            registrar(o, { ok: true, ...r });
            const sinLeer = r.comunicaciones.filter(x => !x.estaLeida).length;
            const sinNotif = r.comunicaciones.filter(x => formal(x) && !x.estaNotificada).length;
            log(`   ${r.comunicaciones.length} comunicación(es) · ${sinLeer} sin leer · ${sinNotif} intimación/notificación sin notificar aún` +
                (r.comunicaciones[0] ? ` · última ${r.comunicaciones[0].puestaADisposicion}` : ''));
            ultimo = null;
            break;
          } catch (e) {
            ultimo = e;
            if (!browser.isConnected()) throw Object.assign(new Error('Se cerró el navegador. Volvé a correrlo y retoma donde quedó.'), { code: 'navegador_cerrado' });
            if (e.code === 'arba_login_rechazado' || intento === 2) break;
            log(`   falló (${e.code || 'error'}): ${e.message.split('\n')[0]} — reintento`);
          } finally {
            await context.close().catch(() => {});
          }
        }
        if (ultimo) {
          const msg = ultimo.message.split('\n')[0];
          log(`   ERROR (${ultimo.code || 'error'}): ${msg}`);
          if (ultimo.code === 'arba_login_rechazado') guardarRechazo(o, msg);
          registrar(o, { ok: false, error: msg, code: ultimo.code || 'error' });
        }
      }
    } finally {
      await browser.close().catch(() => {});
    }
  }

  // ── Consolidado ─────────────────────────────────────────────────────────
  const lote = {
    fecha: iso, dias: DIAS, maestroError,
    claves: todos.map(o => ({ pos: o.pos, titular: o.titular, cuit: o.cuit, tieneCit: Boolean(o.clave) })),
    maestro,
    sociedades: todos.map(o => estado.sociedades[o.cuit]).filter(Boolean),
  };
  const consolidado = path.join(DIR, `DFE_ARBA_CONSOLIDADO_${iso}.xlsx`);
  const tmp = path.join(CRUDO, 'lote.json');
  fs.writeFileSync(tmp, JSON.stringify(lote), 'utf-8');
  let excelError = null;
  try {
    execFileSync(PYTHON, [path.join(__dirname, 'dfe_arba_a_excel.py'), tmp, consolidado],
                 { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
  } catch (e) {
    excelError = String(e.stderr || e.message).trim().split(/\r?\n/).pop();
    log(`no pude armar el consolidado: ${excelError}`);
  }

  const desde = new Date(HOY.getTime() - DIAS * 864e5).toISOString().slice(0, 10);
  const ok = lote.sociedades.filter(s => s.ok);
  const resumen = (s) => ({
    sociedad: s.sociedad, cuit: s.cuit, total: s.comunicaciones.length,
    sinLeer: s.comunicaciones.filter(x => !x.estaLeida).length,
    formalesSinLeer: s.comunicaciones.filter(x => formal(x) && !x.estaLeida).length,
    sinNotificar: s.comunicaciones.filter(x => formal(x) && !x.estaNotificada).length,
    recientes: s.comunicaciones.filter(x => x.puestaADisposicion >= desde).length,
  });
  console.log(JSON.stringify({
    ok: !excelError,
    fecha: iso,
    consolidado: excelError ? null : consolidado,
    relevadas: ok.length,
    conComunicacionesRecientes: ok.map(resumen).filter(x => x.recientes > 0),
    sinNotificarAun: ok.map(resumen).filter(x => x.sinNotificar > 0),
    sinCit: todos.filter(o => !o.clave).map(o => `#${o.pos} ${o.titular}`),
    errores: lote.sociedades.filter(s => !s.ok).map(s => ({ pos: s.pos, sociedad: s.sociedad, cuit: s.cuit, code: s.code, error: s.error })),
  }, null, 2));
})().catch(e => fail(e.message, { code: e.code || 'error' }));
