#!/usr/bin/env node
/**
 * leer-borrador-compras.js — exporta el Libro Compras del borrador ACTUAL del
 * Portal IVA (ARCA) de una sociedad, tal como esta hoy.
 *
 *   node scripts/leer-borrador-compras.js --cuit=<CUIT> --periodo=09/2026 --destino=<carpeta>
 *   node scripts/leer-borrador-compras.js --cuit=<CUIT> --login-nombre=<titular>
 *
 * SOLO LECTURA: no importa desde ARCA (importar volveria a traer las NCR que
 * conciliacion-ncr ya elimino), no elimina, no carga y no presenta. Abre el
 * borrador del periodo, entra al Libro Compras y baja el CSV (ARCA lo entrega en
 * un ZIP o pelado: el que llama resuelve). Si el periodo ya no se puede abrir
 * como borrador porque la DDJJ esta presentada, baja el Libro Compras de la
 * ultima presentacion (modo "presentada"), sin abrir rectificativas.
 *
 * Argumentos:
 *   --cuit / --nombre   Sociedad (titular en Claves_Organismos.xlsx, o CUIT).
 *   --periodo           MM/AAAA o AAAAMM. Default: el mes anterior.
 *   --destino           Carpeta donde dejar el archivo (se crea si falta).
 *   --login-nombre      Titular de la planilla con cuya clave se entra cuando la
 *                       de la sociedad no tiene la representada en el Portal IVA.
 *                       Default: BVA_ARCA_LOGIN_NOMBRE del .env. Sin eso, si la
 *                       representada no aparece se corta con code=sin_representacion.
 *   --presentada        Va directo al libro de la DDJJ presentada.
 *
 * Salida: JSON por stdout ({ ok, modo, archivo, total_en_pantalla, login, ddjj? }); progreso
 * por stderr. Navegador oculto salvo PLAYWRIGHT_HEADLESS=false.
 */

const fs   = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { loadEnv, resolveCredentials } = require('../lib/env');
const { login } = require('../lib/arca-login');
const pi = require('../lib/portal-iva');

loadEnv();
const args = parseArgs();

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
if (!args['cuit'] && !args['nombre']) fail('Falta --cuit o --nombre');
const DESTINO = path.resolve(args['destino'] || process.cwd());
fs.mkdirSync(DESTINO, { recursive: true });
const LOGIN_NOMBRE = args['login-nombre'] || process.env.BVA_ARCA_LOGIN_NOMBRE || '';
const PRESENTADA = 'presentada' in args;

const log = (m) => process.stderr.write(`[borrador-compras] ${m}\n`);

// modo 'borrador': Nueva declaracion jurada -> periodo -> Libro Compras.
// modo 'presentada': Declaraciones juradas presentadas -> Libro IVA -> "Ver" de la ultima secuencia -> Libro
// Compras (sin IMPORTAR). Se usa cuando el periodo ya no se puede abrir como borrador porque esta presentado.
async function exportar(cred, cuitObjetivo, modo) {
  const browser = await chromium.launch(launchOptions());
  const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
  const page = await context.newPage();
  try {
    await login(page, { cuitLogin: cred.cuitLogin, password: cred.password });
    let iva = await pi.abrirPortalIva(context, page);
    await pi.asegurarRepresentacion(iva, cuitObjetivo);
    let ddjj = null;
    if (modo === 'presentada') {
      const p = await pi.abrirPresentada(context, iva, PERIODO);
      iva = p.vista;
      ddjj = { formulario: p.elegida.formulario, secuencia: p.elegida.secuencia, presentada_el: p.elegida.presentada,
               secuencias_del_periodo: p.secuencias.length, totales_vista_previa: p.totales.compras };
      await pi.irAlMenuPresentacion(iva);
      await pi.abrirLibro(iva, 'compras');
      await iva.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
      if (await iva.locator('a, button').filter({ hasText: /^\s*IMPORTAR\s*$/i }).first().isVisible().catch(() => false)) {
        throw Object.assign(new Error('El Libro Compras muestra IMPORTAR: no parece la DDJJ presentada. No sigo.'), { code: 'presentada' });
      }
    } else {
      // Si el periodo no esta en el combo de "Nueva declaracion jurada" es porque ya se presento.
      try {
        await pi.abrirBorrador(iva, PERIODO);
      } catch (e) {
        if (/selectOption|did not find some options/i.test(e.message)) {
          throw Object.assign(new Error(`El periodo ${PERIODO} no se puede abrir como borrador (DDJJ presentada)`), { code: 'ddjj_presentada' });
        }
        throw e;
      }
      if (/verDatosInicialesPresentacion/i.test(iva.url())) {
        throw Object.assign(new Error(`El borrador de ${PERIODO} no esta iniciado en el Portal IVA`), { code: 'borrador_no_iniciado' });
      }
      await pi.abrirLibro(iva, 'compras');
    }
    await pi.cerrarModales(iva);
    await iva.waitForTimeout(5000);
    const total = await iva.evaluate(() => (document.body.innerText.match(/de un total de ([\d.]+)/i) || [])[1]);
    const archivo = await pi.descargarCsv(iva, DESTINO);
    return { modo, archivo, total_en_pantalla: total ? Number(String(total).replace(/\./g, '')) : null, ...(ddjj ? { ddjj } : {}) };
  } finally {
    await browser.close();
  }
}

// Borrador; si el periodo ya esta presentado, el libro de la DDJJ presentada.
async function exportarLibro(cred, cuitObjetivo) {
  if (PRESENTADA) return exportar(cred, cuitObjetivo, 'presentada');
  try {
    return await exportar(cred, cuitObjetivo, 'borrador');
  } catch (e) {
    if (e.code !== 'ddjj_presentada') throw e;
    log(`${e.message}: bajo el Libro Compras de la DDJJ presentada`);
    return exportar(cred, cuitObjetivo, 'presentada');
  }
}

(async () => {
  const propia = resolveCredentials({ cuit: args['cuit'], nombre: args['nombre'] });
  const cuit = propia.cuit;
  try {
    const r = await exportarLibro(propia, cuit);
    console.log(JSON.stringify({ ok: true, cuit, periodo: PERIODO, login: 'propio', ...r }));
  } catch (e) {
    const sinRepr = /no aparece entre las representadas/i.test(e.message);
    if (!sinRepr) {
      console.log(JSON.stringify({ ok: false, cuit, periodo: PERIODO, code: e.code || 'portal', error: e.message }));
      process.exit(1);
    }
    if (!LOGIN_NOMBRE) {
      console.log(JSON.stringify({ ok: false, cuit, periodo: PERIODO, code: 'sin_representacion',
        error: `${e.message} Pasa --login-nombre=<titular de la planilla que la representa> (o BVA_ARCA_LOGIN_NOMBRE en el .env).` }));
      process.exit(1);
    }
    log(`la clave de la sociedad no tiene la representada; entro como "${LOGIN_NOMBRE}"`);
    try {
      const otra = resolveCredentials({ nombre: LOGIN_NOMBRE });
      const r = await exportarLibro(otra, cuit);
      console.log(JSON.stringify({ ok: true, cuit, periodo: PERIODO, login: LOGIN_NOMBRE, ...r }));
    } catch (e2) {
      console.log(JSON.stringify({ ok: false, cuit, periodo: PERIODO, code: e2.code || 'portal', error: e2.message }));
      process.exit(1);
    }
  }
})();
