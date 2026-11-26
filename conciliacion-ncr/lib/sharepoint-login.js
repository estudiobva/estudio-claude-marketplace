// sharepoint-login.js — ingreso al SharePoint del estudio (Microsoft 365) sin tipear nada.
//
// Mismo patrón que arca-login.js: las credenciales salen de Claves_Organismos.xlsx
// (grupo de columnas SHAREPOINT: Usuario | Clave) o de
// SHAREPOINT_USUARIO / SHAREPOINT_CLAVE en el entorno. Nunca se imprime la clave.
//
// La sesión se guarda en ~/.bva-playwright/sharepoint-state.json (permisos 600) y se
// reutiliza en las corridas siguientes: Microsoft no pide la clave de nuevo mientras
// la sesión siga viva, así que se evita un login por corrida (y el riesgo de bloqueo).
//
//   const { abrirSharepoint } = require('./lib/sharepoint-login');
//   const { browser, context, page } = await abrirSharepoint();   // headless por defecto
//   const r = await page.evaluate(() => fetch('/sites/EstudioContable/_api/web').then(r => r.status));

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const { chromium } = require('playwright');
const { launchOptions } = require('./launch');
const { loadEnv } = require('./env');
const claves = require('./claves');

// URL del sitio: SHAREPOINT_SITIO en el entorno o el .env (en el plugin no hay valor por defecto).
function sitioDefault() {
  loadEnv();
  return (process.env.SHAREPOINT_SITIO || '').replace(/\/+$/, '');
}
const ESTADO_DIR  = path.join(os.homedir(), '.bva-playwright');
const ESTADO_FILE = path.join(ESTADO_DIR, 'sharepoint-state.json');

function err(msg, code) { const e = new Error(msg); e.code = code; return e; }

// ── Credenciales ────────────────────────────────────────────────────────────
function credencialesSharepoint() {
  loadEnv();
  const u = (process.env.SHAREPOINT_USUARIO || '').trim();
  const c = process.env.SHAREPOINT_CLAVE || '';
  if (u && c) return { usuario: u, clave: c, origen: 'env' };

  if (!claves.filePath()) {
    throw err('No encontre Claves_Organismos.xlsx ni SHAREPOINT_USUARIO/SHAREPOINT_CLAVE en el .env.', 'sin_credenciales');
  }
  const { rows, file } = claves.loadSync();
  const conSp = rows.filter((r) => r.organismos.SHAREPOINT && r.organismos.SHAREPOINT.usuario && r.organismos.SHAREPOINT.clave);
  const pares = new Map(conSp.map((r) => [`${r.organismos.SHAREPOINT.usuario}\u0000${r.organismos.SHAREPOINT.clave}`, r]));
  if (pares.size === 1) {
    const s = [...pares.values()][0].organismos.SHAREPOINT;
    return { usuario: s.usuario, clave: s.clave, origen: 'archivo' };
  }
  if (pares.size === 0) {
    throw err(`No hay usuario y clave de SharePoint en ${file}: completa el grupo SHAREPOINT (Usuario | Clave).`, 'sin_credenciales');
  }
  throw err(`Hay ${pares.size} usuarios de SharePoint distintos en ${file}: deja uno solo.`, 'sin_credenciales');
}

// ── Sesión guardada ─────────────────────────────────────────────────────────
function guardarEstado(storageState) {
  fs.mkdirSync(ESTADO_DIR, { recursive: true, mode: 0o700 });
  fs.writeFileSync(ESTADO_FILE, JSON.stringify(storageState), { mode: 0o600 });
  try { fs.chmodSync(ESTADO_FILE, 0o600); } catch (_) {}
}
function hayEstado() { return fs.existsSync(ESTADO_FILE); }
function borrarEstado() { try { fs.unlinkSync(ESTADO_FILE); } catch (_) {} }

// ¿La sesión sirve? Un GET a la API REST del sitio: 200 = adentro, 401/403/redirect = no.
async function sesionViva(page, sitio) {
  try {
    const r = await page.request.get(`${sitio}/_api/web?$select=Title`, {
      headers: { Accept: 'application/json;odata=nometadata' },
      maxRedirects: 0, timeout: 20000,
    });
    return r.status() === 200;
  } catch (_) { return false; }
}

// ── Login (formulario de Microsoft) ─────────────────────────────────────────
// Pasos: correo → Siguiente → clave → Iniciar sesión → "¿Mantener la sesión?" → Sí.
// Cortes con código (no se reintenta nunca):
//   credenciales_invalidas — Microsoft rechazó usuario o clave
//   segundo_factor         — pide verificación (MFA); un script oculto no la puede resolver
//   captcha                — pide captcha
async function loginMicrosoft(page, { usuario, clave, sitio = sitioDefault(), timeout = 30000 } = {}) {
  await page.goto(sitio, { waitUntil: 'domcontentloaded', timeout });

  // Si ya estaba adentro (sesión vieja del contexto) no hay formulario.
  const aDestino = () => /\.sharepoint\.com/i.test(page.url()) && !/login\.microsoftonline/i.test(page.url());
  if (aDestino()) return page;

  await page.waitForSelector('input[name="loginfmt"]', { timeout: 20000 });
  await page.fill('input[name="loginfmt"]', usuario);
  await page.click('#idSIButton9');

  // Tras el correo puede aparecer: error de cuenta inexistente, o la clave.
  const clavePromesa = page.waitForSelector('input[name="passwd"]:visible', { timeout: 20000 }).then(() => 'clave');
  const errorUsuario = page.waitForSelector('#usernameError', { timeout: 20000 }).then(() => 'usuario_mal');
  const paso = await Promise.race([clavePromesa, errorUsuario]).catch(() => null);
  if (paso === 'usuario_mal') {
    throw err('Microsoft no reconoce el usuario de SharePoint (revisa la fila "Sharepoint - Microsoft" de Claves_Organismos.xlsx).', 'credenciales_invalidas');
  }
  if (paso !== 'clave') {
    const t = await page.innerText('body').catch(() => '');
    if (/captcha|robot/i.test(t)) throw err('Microsoft pidio un captcha.', 'captcha');
    throw err('No aparecio el campo de clave de Microsoft.', 'login_inesperado');
  }

  await page.fill('input[name="passwd"]', clave);
  await page.click('#idSIButton9');

  // Resultado: error de clave, MFA, "¿mantener sesión?" o directo al sitio.
  const limite = Date.now() + 40000;
  while (Date.now() < limite) {
    if (aDestino()) break;
    if (await page.locator('#passwordError').isVisible().catch(() => false)) {
      throw err('Microsoft rechazo la clave de SharePoint. NO reintentar: puede bloquear la cuenta.', 'credenciales_invalidas');
    }
    const t = await page.innerText('body').catch(() => '');
    if (/captcha/i.test(t)) throw err('Microsoft pidio un captcha.', 'captcha');
    if (/verifique su identidad|verify your identity|aprobar solicitud|approve sign.?in|ingrese el c[oó]digo|enter code|m[aá]s informaci[oó]n es necesaria|more information required/i.test(t) ||
        /\/SAS\/|ProcessAuth.*mfa|login\.microsoftonline\.com\/.*\/mfa/i.test(page.url())) {
      throw err('Microsoft pide verificacion en dos pasos (MFA). Un script oculto no la resuelve.', 'segundo_factor');
    }
    // "¿Quiere mantener la sesión iniciada?" → Sí: deja cookies persistentes.
    const kmsi = page.locator('#KmsiCheckboxField');
    if (await kmsi.isVisible().catch(() => false)) { await kmsi.check().catch(() => {}); }
    if (await page.locator('#idSIButton9:visible').isVisible().catch(() => false) &&
        /mantener la sesi|stay signed in/i.test(t)) {
      await page.click('#idSIButton9').catch(() => {});
    }
    await page.waitForTimeout(500);
  }
  if (!aDestino()) {
    throw err(`El login de Microsoft no llego al sitio (quedo en ${new URL(page.url()).host}).`, 'login_inesperado');
  }
  await page.waitForLoadState('domcontentloaded', { timeout }).catch(() => {});
  return page;
}

// ── Punto de entrada ────────────────────────────────────────────────────────
// Devuelve { browser, context, page, entrada } donde entrada = 'sesion_guardada' | 'login'.
// Con forzar=true ignora la sesión guardada. El que llama cierra el browser.
async function abrirSharepoint({ sitio = sitioDefault(), forzar = false, log = () => {} } = {}) {
  if (!sitio) throw err('Falta SHAREPOINT_SITIO (URL del sitio de SharePoint) en el .env.', 'sin_sitio');
  const browser = await chromium.launch(launchOptions());
  try {
    if (!forzar && hayEstado()) {
      const ctx = await browser.newContext({ locale: 'es-AR', acceptDownloads: true, storageState: ESTADO_FILE });
      const pg = await ctx.newPage();
      if (await sesionViva(pg, sitio)) {
        await pg.goto(sitio, { waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => {});
        log('sesion guardada vigente: no hizo falta la clave.');
        return { browser, context: ctx, page: pg, entrada: 'sesion_guardada' };
      }
      log('la sesion guardada vencio: vuelvo a entrar con usuario y clave.');
      await ctx.close();
    }

    const { usuario, clave, origen } = credencialesSharepoint();
    log(`credenciales desde: ${origen === 'archivo' ? 'Claves_Organismos.xlsx' : origen}`);
    const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
    const page = await context.newPage();
    await loginMicrosoft(page, { usuario, clave, sitio });

    if (!(await sesionViva(page, sitio))) {
      throw err('Entro a Microsoft pero el sitio no da acceso (la cuenta puede no tener permisos en EstudioContable).', 'sin_acceso');
    }
    guardarEstado(await context.storageState());
    return { browser, context, page, entrada: 'login' };
  } catch (e) {
    await browser.close().catch(() => {});
    throw e;
  }
}

module.exports = {
  abrirSharepoint, loginMicrosoft, credencialesSharepoint, sesionViva,
  guardarEstado, borrarEstado, hayEstado, sitioDefault, ESTADO_FILE,
};
