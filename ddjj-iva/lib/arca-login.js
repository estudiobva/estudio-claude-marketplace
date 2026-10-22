// ─────────────────────────────────────────────────────────────────────────────
// Vendorizado desde el plugin fisco-ar (MIT) — https://github.com/javiergradiche/fisco-ar-claude-plugin
// Commit de origen: 2ad06df · Copiado el 21/09/2026 al repo del Estudio BVA.
//
// Se copia en vez de depender del plugin para que una actualizacion del
// marketplace no pueda romper un cierre. Si ARCA cambia el login, el arreglo va
// ACA (y conviene mirar upstream por si ya lo resolvieron).
// ─────────────────────────────────────────────────────────────────────────────

// arca-login.js — login con clave fiscal en ARCA (ex AFIP).
//
// El formulario es en dos pasos (CUIT → clave) y hecho con JSF, así que los
// valores se setean por evaluate + eventos input/change: un `fill()` normal no
// dispara los listeners que JSF necesita para registrar el valor.
//
//   const { login } = require('./lib/arca-login');
//   await login(page, { cuitLogin, password });

const LOGIN_URL = 'https://auth.afip.gob.ar/contribuyente_/login.xhtml';

// Escribe en un input de JSF disparando los eventos que la página escucha.
async function setValue(page, selector, value) {
  await page.evaluate(({ sel, val }) => {
    const el = document.querySelector(sel);
    el.value = val;
    el.dispatchEvent(new Event('input',  { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }, { sel: selector, val: value });
}

// Loguea y deja la página en el portal. Lanza un Error con un `code` legible
// cuando ARCA rechaza el acceso, para que quien llama lo reporte como tal.
async function login(page, { cuitLogin, password, timeout = 30000 } = {}) {
  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout });

  await page.waitForSelector('#F1\\:username', { timeout: 15000 });
  await setValue(page, '#F1\\:username', cuitLogin);
  await page.click('#F1\\:btnSiguiente');

  await page.waitForSelector('#F1\\:password', { timeout: 15000 });
  await setValue(page, '#F1\\:password', password);
  await page.click('#F1\\:btnIngresar');

  await page.waitForURL('**/portal/app/**', { timeout: 20000 });

  const bodyText = await page.innerText('body').catch(() => '');
  if (/captcha/i.test(bodyText)) {
    const e = new Error('ARCA pidió un captcha en el login'); e.code = 'captcha'; throw e;
  }
  if (/clave incorrecta|cuit inv/i.test(bodyText)) {
    const e = new Error('ARCA rechazó las credenciales'); e.code = 'credenciales_invalidas'; throw e;
  }
  return page;
}

// Abre un servicio del portal en su pestaña nueva.
//
// El portal sólo lista los servicios "más utilizados"; el resto aparece recién
// al desplegar "Ver todos". Por eso buscar el link directo falla para cualquier
// servicio que el contribuyente no use seguido, aunque esté adherido.
async function abrirServicio(context, page, matcher, { timeout = 30000 } = {}) {
  const buscar = () => page.locator('a, li, h3', { hasText: matcher }).first();

  // Si no está a la vista, desplegar el listado completo.
  if (!(await buscar().isVisible({ timeout: 3000 }).catch(() => false))) {
    await page.locator('text=Ver todos').first().click({ timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(2000);
  }

  try {
    const [servicio] = await Promise.all([
      context.waitForEvent('page', { timeout }),
      buscar().click({ timeout: 10000 }),
    ]);
    await servicio.waitForLoadState('domcontentloaded', { timeout });
    return servicio;
  } catch (e) {
    const abierta = context.pages().find(p => /rcel|wsfev|fe\.afip|serviciosjava|mcmp/i.test(p.url()));
    if (abierta) return abierta;
    const err = new Error(
      `No pude abrir el servicio ${matcher}. Puede no estar adherido a este CUIT, ` +
      `o el portal cambió. Detalle: ${e.message.split('\n')[0]}`
    );
    err.code = 'servicio_no_disponible';
    throw err;
  }
}

module.exports = { login, setValue, abrirServicio, LOGIN_URL };
