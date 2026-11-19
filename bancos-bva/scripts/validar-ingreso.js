#!/usr/bin/env node
/**
 * validar-ingreso.js — Prueba con que bancos entra el script y con cuales no.
 *
 * Para cada banco (y en Galicia, para cada sociedad del bloque Galicia) intenta
 * UNA vez el login con Claves_Bancos.xlsx, oculto, y si entra cierra la sesion.
 * No navega ni descarga nada. Deja el resultado en
 * ~/Documents/BVA-salidas/extractos/ingreso-bancos.json (o BVA_SALIDAS_PATH),
 * que la skill usa para decidir por donde se baja cada banco:
 *
 *   via "script"  → extractos-banco.js (login automatico).
 *   via "chrome"  → el banco frena al navegador automatizado (captcha, antibots,
 *                   corta la conexion) o pide token: se entra por Claude in Chrome,
 *                   la persona hace el login en su Chrome y Claude navega y baja.
 *   via "claves"  → clave vencida, rechazada, usuario bloqueado o sin datos:
 *                   primero hay que arreglar la clave (ninguna herramienta la cambia).
 *   via "revisar" → no se pudo confirmar (sesion abierta de antes, error de red):
 *                   volver a validar mas tarde.
 *
 *   node scripts/validar-ingreso.js                        todos los bancos, Galicia todas las sociedades
 *   node scripts/validar-ingreso.js --banco=COMAFI,MACRO
 *   node scripts/validar-ingreso.js --banco=GALICIA --sociedad="Sociedad Ejemplo"
 *   node scripts/validar-ingreso.js --banco=GALICIA --solo-fallidas   las que no entraron la vez anterior
 *
 * Se intenta UNA sola vez por usuario: si el banco rechaza la clave no se reintenta.
 * Salida: JSON por stdout; el progreso por stderr. NUNCA imprime credenciales.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { chromium } = require('playwright');

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const cb = require('../lib/claves-bancos');
const bancos = require('../lib/bancos');
const comun = require('../lib/bancos/comun');
const seccion = require('../lib/bancos/seccion-extractos');
const ex = require('../lib/extractos');

loadEnv();
const args = parseArgs();
if ('ver' in args) process.env.PLAYWRIGHT_HEADLESS = 'false';
const { launchOptions } = require('../lib/launch');

const log = (m) => process.stderr.write(`[validar-ingreso] ${m}\n`);
const TODOS = ['GALICIA', 'COMAFI', 'SANTANDER', 'HIPOTECARIO', 'MACRO', 'NACION', 'BBVA', 'PATAGONIA', 'PROVINCIA'];

const VIA = {
  ok: 'script',
  captcha: 'chrome', antibots: 'chrome', sitio_rechazo: 'chrome', segundo_factor: 'chrome',
  credenciales_invalidas: 'claves', clave_vencida: 'claves', bloqueado: 'claves', sin_credenciales: 'claves',
};
const QUE_HACER = {
  script: 'Bajar con extractos-banco.js (login automatico).',
  chrome: 'Bajar por Claude in Chrome: la persona entra al banco en su Chrome (usuario, clave y captcha/token a mano) y Claude navega hasta los resumenes y los descarga; despues archivar-descargas.js los renombra y archiva.',
  claves: 'Arreglar la clave antes (cambiarla entrando al banco a mano y actualizar Claves_Bancos.xlsx). No reintentar seguido: el banco bloquea el usuario.',
  revisar: 'No se pudo confirmar (sitio_no_carga = la pagina de login no cargo): volver a validar ese banco en unos minutos (si quedo una sesion abierta, Macro la libera a los 10 min).',
};

const archivoResultado = () => path.join(ex.salidaBase(args['salida']), 'ingreso-bancos.json');
function leerAnterior() {
  try { return JSON.parse(fs.readFileSync(archivoResultado(), 'utf8')); } catch { return null; }
}

const codigoDe = (e) => e.code || (/net::ERR_/.test(e.message) ? 'sitio_rechazo'
  : /page\.goto: Timeout/.test(e.message) ? 'sitio_no_carga' : 'error');

async function abrirNavegador() {
  try { return await chromium.launch(launchOptions({ channel: 'chrome' })); }
  catch { return chromium.launch(launchOptions()); }
}

// Captura de la pantalla en que quedo (queda en la PC, en ~/bva-inspeccion; no va a la unidad).
async function capturar(page, banco, code) {
  try {
    const dir = path.join(os.homedir(), 'bva-inspeccion');
    fs.mkdirSync(dir, { recursive: true });
    const f = path.join(dir, `validar-${banco.toLowerCase()}-${code}-${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.png`);
    await page.screenshot({ path: f, timeout: 10000 });
    return f;
  } catch { return null; }
}

// Un intento de login. Devuelve { code, mensaje }.
async function probar(browser, adaptador, cred) {
  const context = await browser.newContext({ locale: 'es-AR' });
  const page = await context.newPage();
  let adentro = false;
  try {
    await adaptador.login(page, cred, { pedirAccion: null });
    // Confirmacion extra: que no siga en el formulario ni en la pantalla de espera.
    await page.waitForTimeout(4000);
    const spec = adaptador.SPEC || {};
    const texto = comun.sinAcentos(await page.innerText('body').catch(() => '')).toLowerCase();
    if (comun.RX.sesionActiva.test(texto))
      throw Object.assign(new Error('El banco informa que hay otra sesion activa con ese usuario (alguien lo esta usando o quedo abierta).'), { code: 'sesion_activa' });
    if (spec.esperando && spec.esperando.test(texto))
      throw Object.assign(new Error('Quedo en la pantalla de verificacion del banco.'), { code: 'antibots' });
    if (spec.marcadorLogin && await comun.visible(page, spec.marcadorLogin))
      throw Object.assign(new Error('Despues de ingresar sigue en el formulario de login.'), { code: 'login_sin_confirmar' });
    if (spec.url && page.url().split(/[?#]/)[0] === spec.url.split(/[?#]/)[0])
      throw Object.assign(new Error(`Sigue en la pagina de login (${page.url()}).`), { code: 'login_sin_confirmar' });
    adentro = true;
    return { code: 'ok', mensaje: `Entra (quedo en ${page.url().split('?')[0]}).`, captura: await capturar(page, adaptador.banco, 'ok') };
  } catch (e) {
    let code = codigoDe(e);
    // BBVA: el antibots (Akamai) deja la pagina en el login sin mensaje de error.
    if (code === 'login_sin_confirmar' && adaptador.SPEC && adaptador.SPEC.esperando) code = 'antibots';
    return { code, mensaje: String(e.message || e).split('\n')[0], captura: await capturar(page, adaptador.banco, code) };
  } finally {
    if (adentro) {
      const pg = context.pages()[context.pages().length - 1] || page;
      const cerro = await Promise.race([seccion.salir(adaptador.banco, pg).catch(() => false),
        new Promise(r => setTimeout(() => r(false), 20000))]);
      log(`  ${cerro ? 'sesion cerrada' : 'sin boton de salida conocido (se cierra el navegador)'}.`);
    }
    await context.close().catch(() => {});
  }
}

(async () => {
  const pedidos = args['banco'] && args['banco'] !== 'todos'
    ? String(args['banco']).split(',').map(b => bancos.canonico(b)).filter(Boolean)
    : TODOS;
  for (const b of pedidos) bancos.adaptador(b);   // valida el nombre antes de abrir nada

  const { galicia } = await cb.load();
  const anterior = leerAnterior();
  const fallidasGalicia = new Set(((anterior && anterior.galicia) || []).filter(g => g.via !== 'script').map(g => g.sociedad));

  const resultado = {
    fecha: new Date().toISOString(),
    bancos: (anterior && anterior.bancos) || {},
    galicia: (anterior && anterior.galicia) || [],
  };
  const browser = await abrirNavegador();
  try {
    for (const banco of pedidos) {
      const adaptador = bancos.adaptador(banco);
      if (banco === 'GALICIA') {
        let filas = galicia.filter(g => g.usuario && g.clave);
        if (args['sociedad']) filas = filas.filter(g => comun.sinAcentos(g.sociedad).toLowerCase()
          .includes(comun.sinAcentos(args['sociedad']).toLowerCase()));
        if ('solo-fallidas' in args) filas = filas.filter(g => fallidasGalicia.has(g.sociedad));
        const sinDatos = galicia.filter(g => !g.usuario || !g.clave).map(g => g.sociedad);
        log(`GALICIA: ${filas.length} sociedad(es) a probar${sinDatos.length ? `, ${sinDatos.length} sin datos en el Excel` : ''}.`);
        for (const g of filas) {
          let cred;
          try { cred = await cb.buscar({ banco: 'GALICIA', sociedad: g.sociedad }); }
          catch (e) { cred = null; }
          const r = cred ? await probar(browser, adaptador, cred) : { code: 'sin_credenciales', mensaje: 'Sin usuario/clave.' };
          cred = null;
          const via = VIA[r.code] || 'revisar';
          log(`GALICIA · ${g.sociedad}: ${r.code} → ${via}`);
          resultado.galicia = resultado.galicia.filter(x => x.sociedad !== g.sociedad)
            .concat({ sociedad: g.sociedad, code: r.code, via, mensaje: r.mensaje, ...(r.captura ? { captura: r.captura } : {}), fecha: new Date().toISOString() });
          await new Promise(res => setTimeout(res, 3000));
        }
        for (const s of sinDatos) {
          resultado.galicia = resultado.galicia.filter(x => x.sociedad !== s)
            .concat({ sociedad: s, code: 'sin_credenciales', via: 'claves', mensaje: 'Sin usuario/clave en el bloque Galicia.' });
        }
        const g = resultado.galicia;
        const cuenta = (v) => g.filter(x => x.via === v).length;
        resultado.bancos.GALICIA = {
          code: cuenta('script') ? 'ok' : (g[0] && g[0].code) || 'sin_credenciales',
          via: cuenta('script') ? 'script' : 'claves',
          mensaje: `Por sociedad: ${cuenta('script')} entran, ${cuenta('chrome')} por Chrome, ${cuenta('claves')} con clave a arreglar, ${cuenta('revisar')} a revisar.`,
          fecha: new Date().toISOString(),
        };
        continue;
      }
      let cred;
      try { cred = await cb.buscar({ banco: bancos.nombreEnClaves(banco) }); }
      catch (e) { cred = null; }
      log(`${banco}: probando…`);
      const r = cred ? await probar(browser, adaptador, cred) : { code: 'sin_credenciales', mensaje: 'Sin usuario/clave en Claves_Bancos.xlsx.' };
      cred = null;
      const via = VIA[r.code] || 'revisar';
      log(`${banco}: ${r.code} → ${via}`);
      resultado.bancos[banco] = { code: r.code, via, mensaje: r.mensaje, ...(r.captura ? { captura: r.captura } : {}), fecha: new Date().toISOString() };
    }
  } finally {
    await browser.close().catch(() => {});
  }

  for (const v of Object.values(resultado.bancos)) v.que_hacer = QUE_HACER[v.via];
  const destino = archivoResultado();
  fs.mkdirSync(path.dirname(destino), { recursive: true });
  fs.writeFileSync(destino, JSON.stringify(resultado, null, 2));
  log(`resultado guardado en ${destino}`);
  console.log(JSON.stringify({ ok: true, archivo: destino, ...resultado }, null, 2));
})().catch((e) => fail(e.message, { code: e.code || 'error' }));
