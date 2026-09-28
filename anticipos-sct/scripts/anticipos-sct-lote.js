#!/usr/bin/env node
/**
 * anticipos-sct-lote.js — releva los anticipos del SCT de TODAS las sociedades
 * del archivo de claves (o de las que se pidan), un Excel por sociedad mas un
 * consolidado con la hoja de control.
 *
 *   node scripts/anticipos-sct-lote.js --ver
 *   node scripts/anticipos-sct-lote.js --solo=7,12,30        (posiciones del Excel de claves)
 *   node scripts/anticipos-sct-lote.js --rehacer             (vuelve a relevar las ya hechas hoy)
 *
 * Argumentos:
 *   --solo      Posiciones de Claves_Organismos.xlsx, separadas por coma.
 *   --rehacer   Por defecto, lo ya relevado hoy con exito se saltea (reanudable).
 *   --salida    Carpeta base (default: ~/Documents/BVA-salidas/anticipos, o
 *               $BVA_SALIDAS_PATH/anticipos). Se crea una
 *               subcarpeta por dia.
 *   --consolidar  No entra a ARCA: solo rearma el consolidado con lo relevado hoy.
 *   --ver       Navegador visible.
 *
 * Por que no llama a anticipos-sct.js en un loop: 49 sociedades entran con el
 * mismo CUIT de apoderado. Loguear 69 veces seguidas hace rebotar al SCT, y si
 * ARCA rechaza esa clave, reintentarla 49 veces la bloquea para todas. Aca se
 * loguea UNA vez por apoderado y se cambia de representado dentro del SCT; si
 * el login falla, se marca todo ese grupo y no se reintenta.
 *
 * Esto solo LEE el SCT. Nunca modifica, elimina, presenta ni paga nada.
 */

const fs   = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const { login } = require('../lib/arca-login');
const sctLib = require('../lib/sct');

const args = parseArgs();
const VER = 'ver' in args;
const REHACER = 'rehacer' in args;
const SOLO_CONSOLIDAR = 'consolidar' in args;
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const log = (m) => process.stderr.write(`[lote] ${m}\n`);

const HOY = new Date();
const iso = [HOY.getFullYear(), String(HOY.getMonth() + 1).padStart(2, '0'),
             String(HOY.getDate()).padStart(2, '0')].join('-');
const slug = (s) => String(s).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');

loadEnv(); // antes de BASE: BVA_SALIDAS_PATH puede venir del .env
const BASE = args['salida'] || path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'anticipos');
const DIR = path.join(BASE, iso);
const CRUDO = path.join(DIR, '_crudo');
const RESUMEN = path.join(DIR, 'lote-resumen.json');

// ── A quien relevar ─────────────────────────────────────────────────────────
let rows;
try { ({ rows } = claves.loadSync()); }
catch (e) { fail(`No pude leer el archivo de claves: ${e.message}`); }

const solo = args['solo'] ? new Set(String(args['solo']).split(',').map(x => x.trim())) : null;
const vistos = new Map();
const todos = [];      // todo el Excel de claves (para el consolidado)
const omitidos = [];
for (const r of rows) {
  const c = claves.credenciales(r, 'ARCA');
  if (!r.cuit || !c.clave) { omitidos.push({ pos: r.pos, titular: r.titular, motivo: 'sin CUIT o sin clave de ARCA' }); continue; }
  if (vistos.has(r.cuit)) {
    omitidos.push({ pos: r.pos, titular: r.titular, motivo: `CUIT repetido (ya esta como "${vistos.get(r.cuit)}")` });
    continue;
  }
  vistos.set(r.cuit, r.titular);
  todos.push({ pos: r.pos, titular: r.titular, cuit: r.cuit,
               login: claves.normCuit(c.usuario) || r.cuit, clave: c.clave });
}
// --solo filtra que se releva, no que entra al consolidado: el consolidado
// siempre junta todo lo relevado en el dia.
const objetivos = SOLO_CONSOLIDAR ? [] : todos.filter(o => !solo || solo.has(String(o.pos).trim()));
if (!objetivos.length && !SOLO_CONSOLIDAR) fail('No hay ninguna sociedad para relevar con ese filtro.');

// Agrupar por CUIT de login, respetando el orden del Excel.
const grupos = new Map();
for (const o of objetivos) {
  if (!grupos.has(o.login)) grupos.set(o.login, []);
  grupos.get(o.login).push(o);
}

// ── Estado (reanudable) ─────────────────────────────────────────────────────
fs.mkdirSync(CRUDO, { recursive: true });
const estado = fs.existsSync(RESUMEN) ? JSON.parse(fs.readFileSync(RESUMEN, 'utf-8')) : { fecha: iso, sociedades: {} };
const guardar = () => fs.writeFileSync(RESUMEN, JSON.stringify(estado, null, 1), 'utf-8');

function registrar(o, datos) {
  estado.sociedades[o.cuit] = { pos: o.pos, sociedad: o.titular, cuit: o.cuit, ...datos, cuando: new Date().toISOString() };
  guardar();
}

const yaHecha = (o) => {
  const s = estado.sociedades[o.cuit];
  return !REHACER && s && s.ok && s.archivo && fs.existsSync(s.archivo);
};

function excel(payload, destino) {
  const tmp = path.join(CRUDO, `anticipos-${payload.cuit}.json`);
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  try {
    const out = execFileSync(PYTHON, [path.join(__dirname, 'anticipos_a_excel.py'), tmp, destino],
                             { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
    return JSON.parse(out.trim().split(/\r?\n/).pop());
  } catch (err) {
    const e = new Error(`No pude armar el Excel (${PYTHON}): ` +
      String(err.stderr || err.message).trim().split(/\r?\n/).pop());
    e.code = 'excel_fallo';
    throw e;
  }
}

// ── Relevar una sociedad con el SCT ya posicionado en su CUIT ───────────────
async function relevar(sct, o) {
  const marco = sct.frames().find(f => /homeContrib/.test(f.url())) || sct.mainFrame();
  const { encabezados, filas, resumen } = await sctLib.leerDeudas(sct, marco);
  const { anticipos, inesperados, otrasDeudas } = sctLib.clasificar(filas, HOY);

  fs.writeFileSync(path.join(CRUDO, `deudas-${o.cuit}.json`),
                   JSON.stringify({ encabezados, filas }, null, 1), 'utf-8');

  const destino = path.join(DIR, `Anticipos_SCT_${slug(o.titular)}_${iso}.xlsx`);
  const payload = { sociedad: o.titular, cuit: o.cuit, fecha: iso, anticipos, inesperados, resumen };
  excel(payload, destino);

  const vencidos = anticipos.filter(a => a.estado === 'Vencido').length;
  return {
    ok: true, archivo: destino,
    anticipos: anticipos.length, vencidos, noVencidos: anticipos.length - vencidos,
    inesperados, otrasDeudas, resumenSct: resumen, detalle: anticipos,
  };
}

// Errores que no se arreglan reintentando.
const SIN_REINTENTO = new Set(['cuit_no_disponible', 'excel_fallo']);

(async () => {
  const total = objetivos.length;
  log(`${total} sociedad(es) en ${grupos.size} grupo(s) de login · salida: ${DIR}`);
  for (const om of omitidos) log(`omitida #${om.pos} ${om.titular}: ${om.motivo}`);
  let n = 0;

  for (const [cuitLogin, grupo] of grupos) {
    const pendientes = grupo.filter(o => !yaHecha(o));
    n += grupo.length - pendientes.length;
    if (!pendientes.length) continue;
    log(`── login ${cuitLogin.slice(0, 2)}…${cuitLogin.slice(-3)}: ${pendientes.length} sociedad(es)`);

    const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 150 } : {}));
    const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1500, height: 950 } });
    const page = await context.newPage();
    try {
      try {
        await login(page, { cuitLogin, password: pendientes[0].clave });
      } catch (e) {
        // Credenciales rechazadas o captcha: NO se reintenta (bloquearia la
        // clave fiscal de todo el grupo). Se marcan todas y se sigue.
        const motivo = e.code === 'credenciales_invalidas'
          ? 'ARCA rechazo las credenciales del apoderado (no se reintento para no bloquear la clave)'
          : e.code === 'captcha' ? 'ARCA pidio captcha en el login' : `login fallo: ${e.message}`;
        log(`   ${motivo}`);
        for (const o of pendientes) { n++; registrar(o, { ok: false, error: motivo, code: e.code || 'login' }); }
        continue;
      }

      let sct = null;
      for (const o of pendientes) {
        n++;
        log(`[${n}/${total}] #${o.pos} ${o.titular} (${o.cuit})`);
        let ultimo = null;
        for (let intento = 1; intento <= 2; intento++) {
          try {
            if (!sct || sct.isClosed()) ({ sct } = await sctLib.abrirSct(context, page, o.cuit));
            else await sctLib.posicionar(sct, o.cuit);
            const r = await relevar(sct, o);
            registrar(o, r);
            log(`   ${r.anticipos} anticipo(s) · ${r.vencidos} vencido(s)` +
                (r.inesperados.length ? ` · NOVEDAD: ${r.inesperados.length} de otro impuesto` : ''));
            ultimo = null;
            break;
          } catch (e) {
            ultimo = e;
            // Si se cerro el navegador (lo cerraron a mano o se corto el
            // proceso) no tiene sentido seguir: cada sociedad restante fallaria
            // igual. Se corta todo; al volver a correr, retoma desde aca.
            if (!browser.isConnected() || /has been closed/i.test(e.message)) {
              const f = new Error('Se cerro el navegador. Corto el lote; volve a correrlo y retoma donde quedo.');
              f.code = 'navegador_cerrado';
              throw f;
            }
            if (SIN_REINTENTO.has(e.code) || intento === 2) break;
            log(`   fallo (${e.code || 'error'}): ${e.message.split('\n')[0]} — reabro el SCT y reintento`);
            if (sct) await sct.close().catch(() => {});
            sct = null;
          }
        }
        if (ultimo) {
          const msg = ultimo.message.split('\n')[0];
          log(`   ERROR (${ultimo.code || 'error'}): ${msg}`);
          registrar(o, { ok: false, error: msg, code: ultimo.code || 'error' });
          if (sct) await sct.close().catch(() => {});
          sct = null;
        }
      }
    } finally {
      await browser.close().catch(() => {});
    }
  }

  // ── Consolidado ─────────────────────────────────────────────────────────
  const orden = todos.map(o => estado.sociedades[o.cuit]).filter(Boolean);
  const consolidado = path.join(DIR, `Anticipos_SCT_CONSOLIDADO_${iso}.xlsx`);
  const lote = { fecha: iso, sociedades: orden, omitidos };
  const tmp = path.join(CRUDO, 'lote.json');
  fs.writeFileSync(tmp, JSON.stringify(lote), 'utf-8');
  try {
    execFileSync(PYTHON, [path.join(__dirname, 'anticipos_a_excel.py'), '--consolidado', tmp, consolidado],
                 { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
  } catch (e) {
    log(`no pude armar el consolidado: ${String(e.stderr || e.message).trim().split(/\r?\n/).pop()}`);
  }

  const ok = orden.filter(s => s.ok);
  const err = orden.filter(s => !s.ok);
  const sum = (k) => ok.reduce((a, s) => a + (s[k] || 0), 0);
  console.log(JSON.stringify({
    ok: err.length === 0,
    fecha: iso,
    carpeta: DIR,
    consolidado,
    relevadas: ok.length,
    conAnticipos: ok.filter(s => s.anticipos > 0).length,
    sinAnticipos: ok.filter(s => s.anticipos === 0).length,
    anticipos: sum('anticipos'), vencidos: sum('vencidos'), noVencidos: sum('noVencidos'),
    novedades: ok.filter(s => s.inesperados && s.inesperados.length)
      .map(s => ({ sociedad: s.sociedad, cuit: s.cuit, inesperados: s.inesperados })),
    errores: err.map(s => ({ pos: s.pos, sociedad: s.sociedad, cuit: s.cuit, code: s.code, error: s.error })),
    omitidos,
  }, null, 2));
})().catch(e => fail(e.message, { code: e.code || 'error' }));
