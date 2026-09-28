// comun.js — motor de login compartido por todos los bancos.
//
// Cada banco se describe con un SPEC (ver lib/bancos/<banco>.js):
//
//   {
//     banco: 'SANTANDER',
//     url: 'https://...',
//     preparar: async (page) => {...},       // opcional: cerrar avisos, etc.
//     pasos: [                               // 1 paso por pantalla del login
//       { campos: [{ sel, dato, tipo }],     // dato: 'usuario' | 'acceso' | 'clave' | fn(cred)
//         humano: 'texto',                   // opcional: algo que tiene que hacer una persona
//         enviar: 'selector del boton' },    //   antes de enviar (ej. tildar un captcha)
//     ],
//     marcadorLogin: 'selector',             // visible mientras se siga en el login
//   }
//
// Los datos salen de Claves_Bancos.xlsx: usuario = USUARIO, acceso = ACCESO/DNI,
// clave = CLAVE ID (en el bloque Galicia: USUARIO y CONTRASEÑA).
//
// Reglas:
//   - Se intenta UNA sola vez. Nunca se reintenta una clave rechazada.
//   - Los mensajes del banco se buscan solo en el texto que APARECIO despues de
//     enviar: varias portadas ya dicen "token" o "clave" de entrada (Santander,
//     Hipotecario, Comafi, Macro) y compararlas enteras daria falsos positivos.
//   - Captchas y tokens los resuelve una persona: si hay `pedirAccion` (ventana
//     visible) se le pide y se espera; si no, se corta con el codigo que toca.

const RX = {
  bloqueado:     /bloquead|suspendid|inhabilitad|revocad/,
  credenciales:  /incorrect|invalid|no coincid|no es valid|verifica (los|tus) datos|datos ingresados|usuario o (la )?clave|error de autenticacion|no pudimos validar/,
  segundoFactor: /token|codigo de (seguridad|verificacion|acceso)|segundo factor|clave dinamica|ingresa el codigo|te enviamos un codigo|sms/,
};

const OTP_SEL =
  'input[autocomplete="one-time-code"]:visible, input[name*="token" i]:visible, ' +
  'input[id*="token" i]:visible, input[name*="otp" i]:visible, input[id*="otp" i]:visible';

// Captchas conocidos: Cloudflare Turnstile (Comafi), reCAPTCHA, hCaptcha. Se
// considera resuelto cuando el banco ya tiene la respuesta en su campo oculto.
const CAPTCHA_SEL = '.cf-turnstile, iframe[src*="challenges.cloudflare.com"], .g-recaptcha, ' +
  'iframe[src*="recaptcha"], .h-captcha, iframe[src*="hcaptcha"]';
const RESPUESTA_SEL = 'input[name="cf-turnstile-response"], textarea[name="g-recaptcha-response"], ' +
  'textarea[name="h-captcha-response"]';

async function captchaPendiente(page) {
  if (!(await page.locator(CAPTCHA_SEL).count().catch(() => 0))) return false;
  const respuestas = await page.locator(RESPUESTA_SEL).evaluateAll(els => els.map(e => e.value)).catch(() => []);
  return !respuestas.some(v => v && v.trim());
}

// Si hay captcha sin resolver: se esperan unos segundos (Turnstile a veces pasa
// solo) y si sigue, lo resuelve la persona en la ventana. Nunca se resuelve solo.
async function resolverCaptcha(page, banco, pedirAccion) {
  for (let i = 0; i < 10 && await captchaPendiente(page); i++) await page.waitForTimeout(800);
  if (!(await captchaPendiente(page))) return;
  if (!pedirAccion) throw error('captcha', `${banco} pide un captcha (corre con --ver).`);
  await pedirAccion(`${banco} pide verificar que sos humano: tilda el captcha en la ventana y toca "Listo" (no toques "Ingresar").`);
  if (await captchaPendiente(page)) throw error('captcha', `${banco}: el captcha sigue sin resolver.`);
}

const sinAcentos = (s) => String(s || '').toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

function error(code, mensaje) {
  const e = new Error(mensaje);
  e.code = code;
  return e;
}

const visible = (page, sel) => page.locator(sel).first().isVisible().catch(() => false);

async function lineas(page) {
  const t = await page.innerText('body').catch(() => '');
  return new Set(sinAcentos(t).split('\n').map(s => s.trim()).filter(Boolean));
}

const nuevo = (antes, ahora) => [...ahora].filter(l => !antes.has(l)).join('\n');

async function completar(page, campo, cred) {
  const valor = typeof campo.dato === 'function' ? campo.dato(cred) : cred[campo.dato];
  if (!valor) {
    throw error('sin_credenciales',
      `Falta el dato "${campo.nombre || campo.dato}" en Claves_Bancos.xlsx para este banco.`);
  }
  const loc = page.locator(campo.sel).first();
  await loc.waitFor({ state: 'visible', timeout: 20000 });
  if (campo.tipo === 'select') {
    await loc.selectOption({ label: valor });
  } else {
    await loc.fill('');
    await loc.fill(String(valor));
  }
}

// Espera el resultado de enviar un paso. Devuelve cuando aparece el paso
// siguiente (`listo`) o, si era el ultimo, cuando se sale del login.
async function esperarResultado(page, { antes, listo, marcadorLogin, timeout, pedirAccion, banco }) {
  const limite = Date.now() + timeout;
  let afuera = 0;
  while (Date.now() < limite) {
    await page.waitForTimeout(700);

    if (listo && await listo()) return;

    const enLogin = await visible(page, marcadorLogin);
    const texto = nuevo(antes, await lineas(page));
    const hayOtp = await page.locator(OTP_SEL).count().catch(() => 0);

    if (hayOtp || /token|otp|verific|desafio|challenge|2fa|mfa/i.test(page.url()) ||
        (enLogin && RX.segundoFactor.test(texto))) {
      if (!pedirAccion) throw error('segundo_factor', `${banco} pidio un token o codigo de verificacion.`);
      await pedirAccion(`${banco} pide un token o codigo. Ingresalo en la ventana y, cuando ya estes adentro, toca "Listo" (o Enter en la terminal).`);
      if (await visible(page, marcadorLogin)) {
        throw error('login_sin_confirmar', 'Despues del codigo la pagina sigue en el login.');
      }
      return;
    }

    if (enLogin) {
      afuera = 0;
      if (RX.bloqueado.test(texto))    throw error('bloqueado', `${banco} informa el usuario bloqueado o suspendido.`);
      if (RX.credenciales.test(texto)) throw error('credenciales_invalidas', `${banco} rechazo los datos de acceso.`);
      continue;
    }
    // Fuera del login dos sondeos seguidos (evita confundir un re-render con el ingreso).
    if (!listo && ++afuera >= 2) {
      await page.waitForLoadState('domcontentloaded').catch(() => {});
      return;
    }
  }
  throw error('login_sin_confirmar',
    `No pude confirmar el ingreso en ${Math.round(timeout / 1000)} s (la pagina quedo en ${page.url()}).`);
}

// Loguea segun el SPEC. Lanza Errores con `code`:
//   sin_credenciales | credenciales_invalidas | bloqueado | segundo_factor |
//   captcha | login_sin_confirmar
// Abre la pagina de login. Algunas (BNA Digital) se recargan solas al cargar:
// si esa segunda navegacion corta la primera, goto da net::ERR_ABORTED aunque la
// pagina termine bien. Se tolera y se reintenta abrir (nunca se envia nada aca).
async function abrir(page, url, timeout) {
  for (let i = 1; ; i++) {
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout });
      return;
    } catch (e) {
      if (!/ERR_ABORTED|interrupted by another navigation|frame was detached/i.test(e.message) || i >= 3) throw e;
      await page.waitForTimeout(2500);
      if (page.url() !== 'about:blank') {
        await page.waitForLoadState('domcontentloaded').catch(() => {});
        return;
      }
    }
  }
}

// URL de ingreso: la del link de Claves_Bancos.xlsx si es del mismo sitio del
// banco (BBVA la trae con parametro: ...SolicitarCredenciales.html?260828); si
// no es una URL o es de otro dominio, la del adaptador. Nunca se va a otro host.
function urlIngreso(spec, cred) {
  // Diagnostico: BVA_IGNORAR_LINK_EXCEL=1 fuerza la URL del adaptador.
  if (process.env.BVA_IGNORAR_LINK_EXCEL === '1') return spec.url;
  try {
    const propia = new URL(spec.url);
    const link = new URL(String((cred && cred.link) || '').trim());
    if (/^https:$/.test(link.protocol) && link.hostname === propia.hostname) return link.href;
  } catch { /* el link es texto ("Banco Comafi") o esta vacio */ }
  return spec.url;
}

async function login(page, spec, cred, { timeout = 45000, pedirAccion } = {}) {
  await abrir(page, urlIngreso(spec, cred), timeout);
  if (spec.preparar) await spec.preparar(page);

  for (let i = 0; i < spec.pasos.length; i++) {
    const paso = spec.pasos[i];
    const siguiente = spec.pasos[i + 1];

    for (const campo of paso.campos) await completar(page, campo, cred);
    // Algunos bancos (BBVA) vacian un campo cuando terminan de inicializar la
    // pagina: se revisa y se vuelve a completar el que haya quedado vacio.
    await page.waitForTimeout(1200);
    for (const campo of paso.campos) {
      if (campo.tipo === 'select') continue;
      const v = await page.locator(campo.sel).first().inputValue().catch(() => 'x');
      if (!v) await completar(page, campo, cred);
    }
    await resolverCaptcha(page, spec.banco, pedirAccion);

    if (paso.humano) {
      if (!pedirAccion) throw error('captcha', `${spec.banco}: ${paso.humano} (corre con --ver).`);
      await pedirAccion(`${spec.banco}: ${paso.humano}`);
    }

    const antes = await lineas(page);
    await page.locator(paso.enviar).first().click({ timeout: 15000 });

    await esperarResultado(page, {
      antes,
      timeout,
      pedirAccion,
      banco: spec.banco,
      marcadorLogin: siguiente ? paso.campos[paso.campos.length - 1].sel : spec.marcadorLogin,
      listo: siguiente ? () => visible(page, siguiente.campos[0].sel) : null,
    });
  }
  return page;
}

// Boton por texto, visible, sin importar mayusculas.
// Se prueban tambien las variantes con voseo y acento que usan los bancos
// ("INGRESÁ" en Macro): has-text no distingue mayusculas pero si acentos.
const VARIANTES = { Ingresar: ['Ingresar', 'Ingresá', 'Ingresa', 'Iniciar sesión', 'Entrar'],
                    Continuar: ['Continuar', 'Continuá', 'Siguiente'], Aceptar: ['Aceptar', 'Aceptá'] };
const boton = (...textos) => [...new Set(textos.flatMap(t => VARIANTES[t] || [t]))]
  .map(t => `button:visible:has-text("${t}"), [role="button"]:visible:has-text("${t}"), ` +
            `input[type="submit"][value*="${t}" i]:visible, a:visible:has-text("${t}")`)
  .join(', ');

module.exports = { login, urlIngreso, error, visible, boton, sinAcentos, RX };
