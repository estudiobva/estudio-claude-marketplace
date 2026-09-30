#!/usr/bin/env node
/**
 * f931.js — baja los F.931 ya presentados de una sociedad, uno por periodo.
 *
 *   node scripts/f931.js --nombre="<sociedad>" --desde=01/2026 --hasta=05/2026
 *   node scripts/f931.js --nombre="<sociedad>" --periodo=05/2026
 *   node scripts/f931.js --nombre="<sociedad>" --listar
 *
 * Argumentos:
 *   --nombre     Titular en Claves_Organismos.xlsx.
 *   --cuit       Alternativa a --nombre.
 *   --desde      Primer periodo (MM/AAAA o AAAAMM).
 *   --hasta      Ultimo periodo. Sin --hasta, baja solo --desde.
 *   --periodo    Un solo periodo (atajo de --desde=X --hasta=X).
 *   --meses=N    Aborta si el rango no tiene exactamente N periodos. Usalo
 *                cuando el usuario diga "tienen que ser 12": es el chequeo que
 *                evita rehacer la tanda entera.
 *   --listar     Solo lista lo que hay presentado, no baja nada.
 *   --original   Ante una rectificativa, bajar igual la original (sec 000).
 *                Por default se baja la secuencia vigente (la mas alta).
 *   --salida     Carpeta destino (default: ~/Documents/BVA-salidas/f931/[CUIT],
 *                o $BVA_SALIDAS_PATH/f931/[CUIT]).
 *   --ver        Navegador visible.
 *
 * Salida: JSON por stdout; progreso por stderr.
 *
 * Esto solo consulta y descarga: nunca genera ni presenta una DDJJ.
 */

const fs   = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { resolveCredentials, envFile } = require('../lib/env');
const { login } = require('../lib/arca-login');
const f931 = require('../lib/f931');

const args = parseArgs();
const VER     = 'ver' in args;
const LISTAR  = 'listar' in args;
const ORIGINAL = 'original' in args;
if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';

const log = (m) => process.stderr.write(`[f931] ${m}\n`);

function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if (/^\d{6}$/.test(s)) return s;
  if ((m = s.match(/^(\d{4})[\/\-.](\d{1,2})$/))) return m[1] + String(m[2]).padStart(2, '0');
  return null;
}

// Expande un rango inclusivo de periodos AAAAMM.
function rango(desde, hasta) {
  const out = [];
  let a = Number(desde.slice(0, 4)), m = Number(desde.slice(4, 6));
  const fa = Number(hasta.slice(0, 4)), fm = Number(hasta.slice(4, 6));
  if (a * 12 + m > fa * 12 + fm) return null;
  while (a * 12 + m <= fa * 12 + fm) {
    out.push(`${a}${String(m).padStart(2, '0')}`);
    if (++m > 12) { m = 1; a++; }
  }
  return out;
}

let DESDE = null, HASTA = null;
if (!LISTAR) {
  DESDE = normalizarPeriodo(args['periodo'] || args['desde']);
  HASTA = normalizarPeriodo(args['periodo'] || args['hasta']) || DESDE;
  if (!DESDE) fail('Falta --desde / --periodo (MM/AAAA o AAAAMM). O usa --listar.');
}

let PERIODOS = null;
if (!LISTAR) {
  PERIODOS = rango(DESDE, HASTA);
  if (!PERIODOS) fail(`El rango esta al reves: --desde=${DESDE} es posterior a --hasta=${HASTA}.`);

  // El usuario casi siempre tiene en la cabeza un ejercicio concreto. Si dijo
  // cuantos meses son y el rango no da, frenar aca sale mucho mas barato que
  // descubrirlo con 14 PDFs ya bajados.
  if (args['meses']) {
    const esperados = Number(args['meses']);
    if (PERIODOS.length !== esperados) {
      fail(`El rango ${DESDE}-${HASTA} tiene ${PERIODOS.length} periodos, no ${esperados}.`, {
        code: 'rango_no_coincide',
        periodos: PERIODOS,
        que_hacer: 'Confirma con el usuario si queria el rango completo o un ejercicio de ' +
                   `${esperados} meses, y volve a correr con las fechas corregidas.`,
      });
    }
  }
}

let CUIT, CUIT_LOGIN, PASSWORD, TITULAR;
try {
  ({ cuit: CUIT, cuitLogin: CUIT_LOGIN, password: PASSWORD, titular: TITULAR } =
    resolveCredentials({
      cuit: args['cuit'], nombre: args['nombre'],
      cuitLogin: args['cuit-login'], password: args['password'],
    }));
} catch (e) {
  fail(e.message, { sugerencia: `Revisa el .env (${envFile() || 'no encontrado'}) o el Excel de claves.` });
}

const SALIDA = args['salida'] || path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'f931', CUIT);

(async () => {
  // launch.js lee PLAYWRIGHT_HEADLESS al cargarse, antes de que --ver lo toque.
  const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 250 } : {}));
  const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
  const page = await context.newPage();

  try {
    log(`${TITULAR || CUIT} · CUIT ${CUIT}`);
    log('entrando a ARCA…');
    await login(page, { cuitLogin: CUIT_LOGIN, password: PASSWORD });

    const del = await f931.abrirDeclaracionEnLinea(context, page, CUIT);
    log('leyendo el listado de DDJJ generadas…');
    const presentadas = await f931.listarDDJJ(del);
    log(`${presentadas.length} DDJJ en el historial de la sociedad`);

    if (LISTAR) {
      console.log(JSON.stringify({
        ok: true, titular: TITULAR || null, cuit: CUIT,
        total: presentadas.length,
        ddjj: presentadas.map(p => ({
          periodo: p.periodo, secuencia: p.secuencia, sec: p.sec,
          empleados: p.empleados, generada: p.generada,
        })),
      }, null, 2));
      return;
    }

    // Por periodo puede haber original + rectificativas. La vigente es la de
    // secuencia mas alta; --original fuerza la 000.
    const porPeriodo = new Map();
    for (const p of presentadas) {
      const prev = porPeriodo.get(p.periodo);
      if (!prev || Number(p.sec) > Number(prev.sec)) porPeriodo.set(p.periodo, p);
    }

    const faltantes = PERIODOS.filter(p => !porPeriodo.has(p));
    const conRectificativa = PERIODOS
      .filter(p => porPeriodo.has(p) && porPeriodo.get(p).esRectificativa)
      .map(p => `${p} (${porPeriodo.get(p).secuencia})`);

    if (faltantes.length) {
      log(`FALTAN ${faltantes.length} periodo(s) sin DDJJ presentada: ${faltantes.join(', ')}`);
    }
    if (conRectificativa.length) {
      log(`con rectificativa: ${conRectificativa.join(', ')} — bajo la ${ORIGINAL ? 'ORIGINAL' : 'VIGENTE'}`);
    }

    fs.mkdirSync(SALIDA, { recursive: true });
    const bajados = [], errores = [];

    for (const periodo of PERIODOS) {
      const fila = porPeriodo.get(periodo);
      if (!fila) continue;
      const sec = ORIGINAL ? '000' : fila.sec;
      const destino = path.join(SALIDA, `${periodo}.pdf`);
      try {
        const info = await f931.descargarF931(del, { periodo, sec }, destino);
        const chk = f931.verificarPdf(destino, periodo);
        if (!chk.unaPagina) {
          log(`AVISO ${periodo}: el PDF salio en ${chk.paginas} paginas (deberia ser 1)`);
        }
        log(`${periodo} · sec ${sec} · ${info.empleados || '?'} empleados · ` +
            `${chk.bytes} bytes · ${chk.paginas} pag.`);
        bajados.push({ ...info, archivo: destino, ...chk });
      } catch (e) {
        log(`ERROR en ${periodo}: ${e.message.split('\n')[0]}`);
        errores.push({ periodo, error: e.message.split('\n')[0] });
      }
    }

    console.log(JSON.stringify({
      ok: errores.length === 0,
      titular: TITULAR || null, cuit: CUIT,
      rango: { desde: DESDE, hasta: HASTA, periodos: PERIODOS.length },
      bajados: bajados.length,
      faltantes,
      conRectificativa,
      salida: SALIDA,
      detalle: bajados.map(b => ({
        periodo: b.periodo, sec: b.sec, empleados: b.empleados,
        sumaRem1: b.rem1, paginas: b.paginas, bytes: b.bytes,
      })),
      errores,
    }, null, 2));
  } catch (e) {
    if (e.code === 'captcha') {
      fail('ARCA pidio un captcha.', { code: 'captcha',
        que_hacer: 'Entra una vez a ARCA a mano con esa clave para resolver el captcha, y reintenta.' });
    }
    if (e.code === 'credenciales_invalidas') {
      fail(`ARCA rechazo las credenciales del CUIT ${CUIT}.`, { code: 'credenciales_invalidas',
        que_hacer: 'Verifica la fila en Claves_Organismos.xlsx. Ojo con reintentar: ARCA bloquea la clave.' });
    }
    fail(e.message, { code: e.code || 'error' });
  } finally {
    await browser.close().catch(() => {});
  }
})();
