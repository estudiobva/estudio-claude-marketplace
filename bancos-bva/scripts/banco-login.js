#!/usr/bin/env node
/**
 * banco-login.js — Ingreso al home banking de empresas, sin tipear nada.
 *
 * Lee usuario y clave de Claves_Bancos.xlsx (cifrado, via lib/claves-bancos.js)
 * y entra con Playwright. Es el primer paso de la skill descarga-extractos-bancarios.
 *
 *   node scripts/banco-login.js --banco=GALICIA --sociedad="<sociedad>" --quedate
 *   node scripts/banco-login.js --banco=GALICIA --inspeccionar
 *   node scripts/banco-login.js --banco=GALICIA --sociedad="<sociedad>" --oculto
 *
 * Argumentos:
 *   --banco         Banco tal como figura en Claves_Bancos.xlsx (GALICIA, FRANCES...).
 *   --cuit          Opcional (Galicia). CUIT de la sociedad: busca su login propio por CUIT.
 *   --sociedad      Opcional. En Galicia usa el login propio de esa sociedad (bloque
 *                   Galicia del archivo); sin esto usa el login unico del banco.
 *   --quedate       Deja el navegador abierto y logueado hasta que apretes Enter.
 *   --inspeccionar  Despues de entrar guarda captura, HTML y la lista de menus/links
 *                   en ~/bva-inspeccion/<banco>-<fecha>/, para armar la navegacion
 *                   hasta los resumenes. Esos archivos traen datos de cuentas:
 *                   quedan en la PC, no en la unidad compartida.
 *   --oculto        Corre sin ventana. Por defecto es VISIBLE: los bancos tienen
 *                   antifraude y en headless es mas facil que bloqueen.
 *
 * Si el banco pide token o codigo (segundo factor) y la ventana esta visible, el
 * script espera a que lo ingreses a mano y apretes Enter; despues sigue.
 *
 * Se intenta UNA sola vez: si el banco rechaza la clave no se reintenta.
 *
 * Salida: JSON por stdout; el progreso va por stderr. NUNCA imprime credenciales.
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const { chromium } = require('playwright');

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const cb = require('../lib/claves-bancos');
const bancos = require('../lib/bancos');

loadEnv();
const args = parseArgs();
const QUEDATE      = 'quedate' in args;
const INSPECCIONAR = 'inspeccionar' in args;
const OCULTO       = 'oculto' in args && !QUEDATE;

// launch.js lee esto del entorno, asi que se setea antes de pedir las opciones.
if (process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = OCULTO ? 'true' : 'false';
const { launchOptions, HEADLESS } = require('../lib/launch');

const log = (m) => process.stderr.write(`[banco-login] ${m}\n`);

function esperarEnter(mensaje) {
  process.stderr.write(`\n${mensaje}\n`);
  return new Promise((resolve) => {
    process.stdin.resume();
    process.stdin.once('data', () => { process.stdin.pause(); resolve(); });
  });
}

// Chrome instalado se parece mas a un navegador normal que el Chromium de
// Playwright; si no esta, se usa Chromium.
async function abrirNavegador() {
  try { return await chromium.launch(launchOptions({ channel: 'chrome' })); }
  catch { log('Chrome no disponible, uso Chromium de Playwright.'); return chromium.launch(launchOptions()); }
}

async function inspeccionar(page, banco) {
  const sello = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const dir = path.join(os.homedir(), 'bva-inspeccion', `${banco.toLowerCase()}-${sello}`);
  fs.mkdirSync(dir, { recursive: true });
  await page.waitForTimeout(3000);
  await page.screenshot({ path: path.join(dir, 'captura.png'), fullPage: true }).catch(() => {});
  fs.writeFileSync(path.join(dir, 'pagina.html'), await page.content().catch(() => ''));
  const estructura = await page.evaluate(() => {
    const visible = (e) => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
    const txt = (e) => (e.innerText || e.getAttribute('aria-label') || e.title || '').trim().replace(/\s+/g, ' ').slice(0, 80);
    return {
      url: location.href,
      titulo: document.title,
      links: [...document.querySelectorAll('a, [role="menuitem"], [role="tab"], button')]
        .filter(visible).map(e => ({ tag: e.tagName, texto: txt(e), href: e.getAttribute('href') || null,
          id: e.id || null, testid: e.getAttribute('data-testid') || null }))
        .filter(x => x.texto),
      selects: [...document.querySelectorAll('select, [role="combobox"], [role="listbox"]')]
        .filter(visible).map(e => ({ id: e.id || null, name: e.getAttribute('name'), texto: txt(e) })),
    };
  }).catch((e) => ({ error: e.message }));
  fs.writeFileSync(path.join(dir, 'estructura.json'), JSON.stringify(estructura, null, 2));
  return dir;
}

(async () => {
  if (!args['banco']) fail('Falta --banco. Ej: --banco=GALICIA');

  let adaptador, cred;
  try {
    adaptador = bancos.adaptador(args['banco']);
    cred = await cb.buscar({ banco: bancos.nombreEnClaves(args['banco']), sociedad: args['sociedad'], cuit: args['cuit'] });
  } catch (e) {
    fail(e.message, { code: e.code || 'credenciales_no_resueltas' });
  }

  log(`banco: ${adaptador.banco}`);
  log(`credenciales: ${cred.origen === 'galicia-por-sociedad'
    ? `login propio de "${cred.sociedad}"` : 'login unico del banco'} (Claves_Bancos.xlsx)`);
  if (HEADLESS) log('corriendo oculto: si el banco bloquea, proba sin --oculto.');

  const browser = await abrirNavegador();
  const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
  const page = await context.newPage();

  let intervencion = false;
  const pedirAccion = HEADLESS ? null : (msg) => { intervencion = true; return esperarEnter(`  ${msg}`); };
  try {
    log('entrando…');
    await adaptador.login(page, cred, { pedirAccion });
    cred = null; // no dejar las credenciales vivas mas de lo necesario
    log('adentro.');

    const carpetaInspeccion = INSPECCIONAR ? await inspeccionar(page, adaptador.banco) : null;
    if (carpetaInspeccion) log(`inspeccion guardada en ${carpetaInspeccion}`);

    console.log(JSON.stringify({
      ok: true,
      banco: adaptador.banco,
      sociedad: args['sociedad'] || null,
      intervencionManual: intervencion,
      url: page.url(),
      inspeccion: carpetaInspeccion,
    }, null, 2));

    if (QUEDATE) {
      await esperarEnter(
        '  El navegador queda abierto y logueado. Segui a mano desde ahi.\n' +
        '  Cuando termines, volve a esta consola y apreta Enter para cerrarlo.');
    }
    await browser.close().catch(() => {});
  } catch (e) {
    const queHacer = {
      credenciales_invalidas: `Revisa la fila de ${adaptador.banco} en Claves_Bancos.xlsx. ` +
                              'No reintentes seguido: el banco bloquea el usuario.',
      bloqueado: 'El usuario esta bloqueado: hay que desbloquearlo con el banco. No reintentar.',
      segundo_factor: 'Corre sin --oculto para ingresar el codigo a mano en la ventana.',
      captcha: 'Corre sin --oculto para tildar el captcha a mano en la ventana.',
      sin_credenciales: 'Completa el dato que falta en Claves_Bancos.xlsx.',
      login_sin_confirmar: 'Corre con --quedate para ver en que pantalla quedo.',
    }[e.code];
    await browser.close().catch(() => {});
    fail(e.message, { code: e.code || 'error', ...(queHacer ? { que_hacer: queHacer } : {}) });
  }
})();
