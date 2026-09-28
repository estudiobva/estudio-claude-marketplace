#!/usr/bin/env node
/**
 * planes-facilidades.js — releva los planes de facilidades de pago (Mis
 * Facilidades de ARCA) de todas las sociedades, o de las que se pidan, y arma
 * un informe consolidado + un Excel por sociedad.
 *
 *   node scripts/planes-facilidades.js --ver                       (todas las del Excel de claves)
 *   node scripts/planes-facilidades.js --login=<apoderado> --ver   (solo las que entran con ese apoderado)
 *   node scripts/planes-facilidades.js --nombre="<sociedad>" --ver   (una sola)
 *   node scripts/planes-facilidades.js --solo=7,12,30              (posiciones del Excel de claves)
 *   node scripts/planes-facilidades.js --consolidar                (rearma el informe sin entrar a ARCA)
 *
 * Argumentos:
 *   --login       Apoderado: nombre del titular de login o su CUIT.
 *   --nombre / --cuit  Una sociedad puntual.
 *   --solo        Posiciones de Claves_Organismos.xlsx, separadas por coma.
 *   --historico   Ademas de los Vigentes, entra al detalle de los planes
 *                 cancelados/caducos (por defecto solo se listan).
 *   --rehacer     Por defecto, lo ya relevado hoy con exito se saltea (reanudable).
 *   --salida      Carpeta base (default: ~/Documents/BVA-salidas/planes, o $BVA_SALIDAS_PATH/planes). Subcarpeta por dia.
 *   --ver         Navegador visible.
 *
 * Por cada sociedad: Presentaciones Enviadas → por cada plan Vigente: Detalle →
 * Ver Pagos + Plan de Pago → cruza ambos y marca cada cuota como Pagada,
 * IMPAGA (vencida sin pago) o A vencer.
 *
 * Loguea UNA vez por apoderado (un apoderado puede entrar a decenas de sociedades). Si ARCA rechaza la
 * clave o pide captcha, marca todo ese grupo y NO reintenta (bloquearia la
 * clave fiscal de todas).
 *
 * Esto solo LEE Mis Facilidades. Nunca presenta, reformula ni paga nada.
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
const fac = require('../lib/facilidades');

const args = parseArgs();
const VER = 'ver' in args;
const REHACER = 'rehacer' in args;
const HISTORICO = 'historico' in args;
const SOLO_CONSOLIDAR = 'consolidar' in args;
if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const log = (m) => process.stderr.write(`[planes] ${m}\n`);

const HOY = new Date();
const iso = [HOY.getFullYear(), String(HOY.getMonth() + 1).padStart(2, '0'),
             String(HOY.getDate()).padStart(2, '0')].join('-');
const compacto = iso.replace(/-/g, '');
loadEnv(); // antes de BASE: BVA_SALIDAS_PATH puede venir del .env
const BASE = args['salida'] || path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'planes');
const DIR = path.join(BASE, iso);
const CRUDO = path.join(DIR, '_crudo');
const RESUMEN = path.join(DIR, 'lote-resumen.json');
const slug = (s) => String(s).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');

// ── A quien relevar ─────────────────────────────────────────────────────────
loadEnv();
let rows;
try { ({ rows } = claves.loadSync()); }
catch (e) { fail(`No pude leer el archivo de claves: ${e.message}`); }

const vistos = new Map();
const todos = [];
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

// Filtros
let objetivos = todos;
if (args['login']) {
  const q = String(args['login']);
  const dig = q.replace(/\D/g, '');
  let loginCuit = dig.length === 11 ? dig : null;
  if (!loginCuit) {
    const hit = rows.filter(r => claves.norm(r.titular).includes(claves.norm(q)) && todos.some(o => o.login === r.cuit));
    if (hit.length !== 1) fail(`--login="${q}": ${hit.length ? 'hay varios apoderados que coinciden' : 'no encontre un apoderado con ese nombre'}.`,
                               { candidatos: hit.map(h => h.titular) });
    loginCuit = hit[0].cuit;
  }
  objetivos = objetivos.filter(o => o.login === loginCuit);
  log(`apoderado ${loginCuit}: ${objetivos.length} sociedad(es)`);
}
if (args['nombre'] || args['cuit']) {
  let b;
  try { b = claves.buscar(rows, { nombre: args['nombre'], cuit: args['cuit'] }); }
  catch (e) { fail(e.message); }
  if (b.candidatos) fail('Hay varias sociedades que coinciden; afina el nombre o pasa --cuit.',
                         { candidatos: b.candidatos.map(c => `${c.titular} (${c.cuit})`) });
  objetivos = objetivos.filter(o => o.cuit === b.row.cuit);
}
if (args['solo']) {
  const solo = new Set(String(args['solo']).split(',').map(x => x.trim()));
  objetivos = objetivos.filter(o => solo.has(String(o.pos).trim()));
}
if (SOLO_CONSOLIDAR) objetivos = [];
if (!objetivos.length && !SOLO_CONSOLIDAR) fail('No hay ninguna sociedad para relevar con ese filtro.');

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
  estado.sociedades[o.cuit] = { pos: o.pos, sociedad: o.titular, cuit: o.cuit, login: o.login,
                                ...datos, cuando: new Date().toISOString() };
  guardar();
}
// version 2 = incluye CBU y obligaciones de cada plan. Lo relevado con una
// version anterior se vuelve a relevar aunque haya salido bien.
const VERSION = 2;
const yaHecha = (o) => {
  const s = estado.sociedades[o.cuit];
  return !REHACER && s && s.ok && s.version === VERSION;
};

function python(argsPy) {
  try {
    const out = execFileSync(PYTHON, [path.join(__dirname, 'planes_a_excel.py'), ...argsPy],
                             { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
    return JSON.parse(out.trim().split(/\r?\n/).pop());
  } catch (err) {
    const e = new Error(`No pude armar el Excel (${PYTHON}): ` +
      String(err.stderr || err.message).trim().split(/\r?\n/).pop());
    e.code = 'excel_fallo';
    throw e;
  }
}

// Limite de tiempo: una pantalla de ARCA que no responde puede dejar un
// evaluate() esperando para siempre (caso real: 25 min colgado). Con
// esto falla, se abre sesion nueva y se reintenta.
function conLimite(promesa, ms, que) {
  let t;
  const limite = new Promise((_, rej) => {
    t = setTimeout(() => rej(Object.assign(new Error(`${que}: sin respuesta de ARCA en ${Math.round(ms / 1000)} s`), { code: 'colgado' })), ms);
  });
  return Promise.race([promesa, limite]).finally(() => clearTimeout(t));
}

// ── Relevar una sociedad ────────────────────────────────────────────────────
async function relevar(context, portal, o) {
  const { mf, razon } = await conLimite(fac.abrirEnCuit(context, portal, o.cuit), 150000, 'Abrir Mis Facilidades');
  try {
    const presentaciones = await fac.leerPresentaciones(mf);
    const enAlcance = presentaciones.filter(p => HISTORICO || /vigente/i.test(p.situacion));
    log(`   ${presentaciones.length} presentacion(es) · ${presentaciones.filter(p => /vigente/i.test(p.situacion)).length} vigente(s)`);

    const planes = [];
    for (const fila of enAlcance) {
      // Si un plan no se puede leer, falla la sociedad entera y se reintenta:
      // un informe con un plan vigente faltante parece "al dia" y no lo esta.
      const crudo = await conLimite(fac.leerPlan(mf, fila), 240000, `Plan ${fila.numero}`);
      const a = fac.analizar({ fila, ...crudo }, HOY);
      fs.writeFileSync(path.join(CRUDO, `plan-${o.cuit}-${fila.numero}.json`), JSON.stringify(crudo, null, 1), 'utf-8');
      log(`   plan ${a.nroPlan} (${a.tipo.slice(0, 40)}): ${a.cuotasPagadas} pagada(s), ` +
          `${a.cuotasImpagas} IMPAGA(S), ${a.cuotasAVencer} a vencer` +
          (a.cuotasImpagas ? ` → ${a.impagas.map(i => `cuota ${i.cuota} vto ${i.vencimiento}`).join('; ')}` : ''));
      planes.push(a);
    }

    const vigentes = planes.filter(p => /vigente/i.test(p.situacion));
    const payload = {
      sociedad: o.titular, razonArca: razon, cuit: o.cuit, fecha: iso,
      presentaciones: presentaciones.map(({ botonId, ...p }) => ({ ...p, consolidado: fac.aNumero(p.consolidado) })),
      planes,
    };
    const jsonSoc = path.join(CRUDO, `sociedad-${o.cuit}.json`);
    fs.writeFileSync(jsonSoc, JSON.stringify(payload, null, 1), 'utf-8');
    const destino = path.join(DIR, `planes_facilidades_${slug(o.titular)}_${compacto}.xlsx`);
    python([jsonSoc, destino]);

    return {
      ok: true, version: VERSION, razonArca: razon, archivo: destino, datos: jsonSoc,
      presentaciones: presentaciones.length,
      vigentes: vigentes.length,
      cuotasImpagas: vigentes.reduce((a, p) => a + (p.cuotasImpagas || 0), 0),
      alerta: vigentes.some(p => p.alerta === 'RIESGO DE CADUCIDAD') ? 'RIESGO DE CADUCIDAD'
            : vigentes.some(p => p.alerta === 'CUOTA IMPAGA') ? 'CUOTA IMPAGA'
            : vigentes.some(p => p.alerta === 'SIN DATO DE PAGOS') ? 'LECTURA INCOMPLETA'
            : vigentes.length ? 'AL DIA' : 'SIN PLANES VIGENTES',
    };
  } finally {
    await mf.close().catch(() => {});
  }
}

const SIN_REINTENTO = new Set(['cuit_no_disponible', 'excel_fallo']);

(async () => {
  const total = objetivos.length;
  if (total) log(`${total} sociedad(es) en ${grupos.size} grupo(s) de login · salida: ${DIR}`);
  let n = 0;

  for (const [cuitLogin, grupo] of grupos) {
    const pendientes = grupo.filter(o => !yaHecha(o));
    n += grupo.length - pendientes.length;
    if (!pendientes.length) continue;
    log(`── login ${cuitLogin.slice(0, 2)}…${cuitLogin.slice(-3)}: ${pendientes.length} sociedad(es)`);

    const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 100 } : {}));
    let context = null, portal = null;
    // Sesion nueva (contexto limpio + login). Se usa al empezar el grupo y
    // cuando el portal deja de abrir Mis Facilidades: ARCA vence la sesion
    // despues de un rato largo y el portal queda muerto sin avisar. Las
    // credenciales ya se aceptaron al principio, asi que reloguear es seguro;
    // igual hay un tope para no insistir si algo anda mal.
    const abrirSesion = async () => {
      if (context) await context.close().catch(() => {});
      context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1400, height: 900 } });
      portal = await context.newPage();
      await login(portal, { cuitLogin, password: pendientes[0].clave });
    };
    const MAX_RELOGINS = 4;
    let relogins = 0;
    try {
      try {
        await abrirSesion();
      } catch (e) {
        const motivo = e.code === 'credenciales_invalidas'
          ? 'ARCA rechazo las credenciales del apoderado (no se reintento para no bloquear la clave)'
          : e.code === 'captcha' ? 'ARCA pidio captcha en el login' : `login fallo: ${e.message.split('\n')[0]}`;
        log(`   ${motivo}`);
        for (const o of pendientes) { n++; registrar(o, { ok: false, error: motivo, code: e.code || 'login' }); }
        continue;
      }

      let cortar = null;
      for (const o of pendientes) {
        n++;
        if (cortar) { registrar(o, { ok: false, error: cortar, code: 'sesion' }); continue; }
        log(`[${n}/${total}] #${o.pos} ${o.titular} (${o.cuit})`);
        let ultimo = null;
        for (let intento = 1; intento <= 2; intento++) {
          try {
            const r = await relevar(context, portal, o);
            registrar(o, r);
            log(`   → ${r.alerta}`);
            ultimo = null;
            break;
          } catch (e) {
            ultimo = e;
            if (!browser.isConnected()) {
              const f = new Error('Se cerro el navegador. Corto el lote; volve a correrlo y retoma donde quedo.');
              f.code = 'navegador_cerrado';
              throw f;
            }
            if (SIN_REINTENTO.has(e.code) || intento === 2) break;
            log(`   fallo (${e.code || 'error'}): ${e.message.split('\n')[0]}`);
            if (relogins < MAX_RELOGINS) {
              relogins++;
              log(`   abro sesion nueva en ARCA (${relogins}/${MAX_RELOGINS}) y reintento`);
              try { await abrirSesion(); }
              catch (le) {
                cortar = `se perdio la sesion y el re-login fallo (${le.code || le.message.split('\n')[0]}); no se insistio`;
                break;
              }
            } else {
              await portal.goto('https://portalcf.cloud.afip.gob.ar/portal/app/', { waitUntil: 'domcontentloaded' }).catch(() => {});
              await portal.waitForTimeout(3000);
            }
          }
        }
        if (cortar) { log(`   ${cortar}`); registrar(o, { ok: false, error: cortar, code: 'sesion' }); continue; }
        if (ultimo) {
          const msg = ultimo.message.split('\n')[0];
          log(`   ERROR (${ultimo.code || 'error'}): ${msg}`);
          registrar(o, { ok: false, error: msg, code: ultimo.code || 'error' });
        }
      }
    } finally {
      await browser.close().catch(() => {});
    }
  }

  const orden = todos.map(o => estado.sociedades[o.cuit]).filter(Boolean);

  // --consolidar vuelve a analizar lo leido (_crudo) con la logica actual y
  // regenera los Excel por sociedad, sin entrar a ARCA.
  if (SOLO_CONSOLIDAR) {
    for (const s of orden.filter(x => x.ok && x.datos && fs.existsSync(x.datos))) {
      const d = JSON.parse(fs.readFileSync(s.datos, 'utf-8'));
      d.planes = d.planes.map(p => {
        const f = path.join(CRUDO, `plan-${s.cuit}-${p.nroPlan}.json`);
        const fila = d.presentaciones.find(x => x.numero === p.nroPlan);
        if (!fs.existsSync(f) || !fila) return p;
        return fac.analizar({ fila, ...JSON.parse(fs.readFileSync(f, 'utf-8')) }, HOY);
      });
      fs.writeFileSync(s.datos, JSON.stringify(d, null, 1), 'utf-8');
      const vig = d.planes.filter(p => /vigente/i.test(p.situacion));
      s.vigentes = vig.length;
      s.cuotasImpagas = vig.reduce((a, p) => a + (p.cuotasImpagas || 0), 0);
      s.alerta = vig.some(p => p.alerta === 'RIESGO DE CADUCIDAD') ? 'RIESGO DE CADUCIDAD'
        : vig.some(p => p.alerta === 'CUOTA IMPAGA') ? 'CUOTA IMPAGA'
        : vig.some(p => p.alerta === 'SIN DATO DE PAGOS') ? 'LECTURA INCOMPLETA'
        : vig.length ? 'AL DIA' : 'SIN PLANES VIGENTES';
      if (s.archivo) python([s.datos, s.archivo]);
    }
    guardar();
    log('re-analizado desde _crudo');
  }

  // ── Informe consolidado ─────────────────────────────────────────────────
  const informe = path.join(DIR, `INFORME_PLANES_FACILIDADES_${compacto}.xlsx`);
  const lote = { fecha: iso, sociedades: orden, omitidos };
  const tmp = path.join(CRUDO, 'lote.json');
  fs.writeFileSync(tmp, JSON.stringify(lote), 'utf-8');
  let res = null;
  try { res = python(['--consolidado', tmp, informe]); }
  catch (e) { log(`no pude armar el informe: ${e.message}`); }
  let reporte = null;
  try {
    const out = execFileSync(process.execPath, [path.join(__dirname, 'planes-reporte.js'), `--fecha=${iso}`, `--salida=${BASE}`],
                             { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
    reporte = JSON.parse(out.trim().split(/\r?\n/).pop()).pdf;
  } catch (e) { log(`no pude armar el reporte PDF: ${String(e.stderr || e.message).trim().split(/\r?\n/).pop()}`); }

  const ok = orden.filter(s => s.ok);
  const err = orden.filter(s => !s.ok);
  const conImpagas = [];
  for (const s of ok) {
    if (!s.cuotasImpagas) continue;
    const d = JSON.parse(fs.readFileSync(s.datos, 'utf-8'));
    for (const p of d.planes.filter(p => /vigente/i.test(p.situacion) && p.cuotasImpagas)) {
      conImpagas.push({ sociedad: s.sociedad, cuit: s.cuit, plan: p.nroPlan, tipo: p.tipo, alerta: p.alerta,
                        cuotasImpagas: p.cuotasImpagas,
                        detalle: p.impagas.map(i => `cuota ${i.cuota} vto ${i.vencimiento} (${i.diasAtraso} dias) $${(i.importe || 0).toFixed(2)}`) });
    }
  }
  console.log(JSON.stringify({
    ok: err.length === 0,
    fecha: iso, carpeta: DIR, informe: res ? informe : null, reporte,
    relevadas: ok.length,
    conPlanesVigentes: ok.filter(s => s.vigentes > 0).length,
    planesVigentes: ok.reduce((a, s) => a + (s.vigentes || 0), 0),
    alDia: ok.filter(s => s.alerta === 'AL DIA').length,
    conCuotasImpagas: conImpagas,
    errores: err.map(s => ({ pos: s.pos, sociedad: s.sociedad, cuit: s.cuit, code: s.code, error: s.error })),
    omitidos,
  }, null, 2));
})().catch(e => fail(e.message, { code: e.code || 'error' }));
