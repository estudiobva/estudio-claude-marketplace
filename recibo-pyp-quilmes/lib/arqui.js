// arqui.js — navegacion de ARQui (Agencia de Recaudacion del Municipio de
// Quilmes, https://arquivirtual.quilmes.gov.ar). Solo lee y descarga: nunca
// toca "Botón de pago", "Mis pagos online", el "+" de agregar cuentas ni las
// declaraciones juradas.
//
// Lo aprendido (22/10/2026):
//   - Login: Login_arqui.aspx, campos #ContentPlaceHolder1_txtMail / _txtPwd y
//     el boton #btnAceptarJS. Si entra, va a App/DDJJ/Sshh/Inicio.aspx.
//   - Publicidad y Propaganda > Recibo Actual = DDJJAddSeleccionActualPyp.aspx:
//     un <select> #ContentPlaceHolder1_cmbCuentas con las cuentas vinculadas al
//     usuario (value = numero SIN digito verificador, ej. 500123456; hay una
//     opcion vacia en el medio) y el boton #ContentPlaceHolder1_Button1
//     "Descargar", que hace un postback y devuelve el PDF como descarga
//     ("Recibo.pdf"). La pagina queda igual, lista para la siguiente cuenta.
//     Si la cuenta no tiene recibo, no hay descarga y aparece el motivo en
//     #ContentPlaceHolder1_msg (ej. "La cuenta ingresado no existe.").
//   - El sitio tiene un firewall que, con muchos pedidos seguidos, contesta 403
//     "Forbidden" a todo (login incluido) durante unos minutos. Por eso no se
//     recarga la pantalla entre cuentas, no se bajan imagenes ni fuentes, hay
//     una pausa entre recibos (ARQUI_PAUSA_MS, default 2500) y ante un 403 se
//     avisa con code 'bloqueo' para que el script espere y reintente.

const BASE = 'https://arquivirtual.quilmes.gov.ar/municipal';
const URL_LOGIN = `${BASE}/Login_arqui.aspx`;
const URL_PYP = `${BASE}/App/DDJJ/Sshh/DDJJAddSeleccionActualPyp.aspx`;
const SEL_CUENTAS = '#ContentPlaceHolder1_cmbCuentas';

class ArquiError extends Error {
  constructor(code, message) { super(message); this.code = code; }
}

const esBloqueo = (resp) => Boolean(resp && resp.status() === 403);
const errBloqueo = () => new ArquiError('bloqueo',
  'ARQui contesto 403 (Forbidden): el firewall del municipio bloqueo temporalmente los pedidos por ser muchos seguidos. Hay que esperar unos minutos.');

// Sin imagenes, fuentes ni videos: menos pedidos al servidor (ver arriba).
async function prepararPagina(page) {
  await page.route('**/*', (route) =>
    (['image', 'font', 'media'].includes(route.request().resourceType()) ? route.abort() : route.continue()));
}

const textoVisible = async (page) =>
  (await page.innerText('body').catch(() => '')).replace(/\s+/g, ' ').trim();

async function login(page, { usuario, clave }) {
  if (esBloqueo(await page.goto(URL_LOGIN, { waitUntil: 'domcontentloaded', timeout: 60000 }))) throw errBloqueo();
  await page.fill('#ContentPlaceHolder1_txtMail', usuario);
  await page.fill('#ContentPlaceHolder1_txtPwd', clave);
  await page.click('#btnAceptarJS');
  try {
    await page.waitForURL(/\/App\//, { timeout: 30000 });
  } catch {
    if (/Forbidden/i.test(await textoVisible(page))) throw errBloqueo();
    // Sigue en el login. El cartel "AVISO IMPORTANTE!" esta siempre en la
    // pagina: no es el motivo. Se toma solo lo visible del formulario.
    const msg = (await page.$$eval('.modal-body, [id*=lbl], [id*=msg], .alert', (es) =>
      es.filter((e) => e.offsetParent !== null)
        .map((e) => e.innerText.trim().replace(/\s+/g, ' '))
        .filter((t) => t && !/^AVISO IMPORTANTE/i.test(t)).join(' | ')).catch(() => '')) || 'sin mensaje';
    throw new ArquiError('login', `ARQui no dejo entrar con el usuario de la planilla (${msg.slice(0, 200)}).`);
  }
}

// Abre la pantalla de Recibo Actual de PyP. Si la sesion vencio, el sitio
// manda al login: se avisa con code 'sesion' para que el script reloguee.
async function abrirPyp(page) {
  if (esBloqueo(await page.goto(URL_PYP, { waitUntil: 'domcontentloaded', timeout: 60000 }))) throw errBloqueo();
  if (/Login_arqui/i.test(page.url())) throw new ArquiError('sesion', 'La sesion de ARQui vencio.');
  const sel = await page.waitForSelector(SEL_CUENTAS, { timeout: 30000 }).catch(() => null);
  if (!sel) {
    throw new ArquiError('pantalla', `No aparecio el desplegable de cuentas de Publicidad y Propaganda (${(await textoVisible(page)).slice(0, 200)}).`);
  }
}

// Cuentas del desplegable (sin digito verificador), en el orden del sitio.
async function cuentasPyp(page) {
  if (!page.url().startsWith(URL_PYP)) await abrirPyp(page);
  return page.$$eval(`${SEL_CUENTAS} option`, (os) => os.map((o) => o.value.trim()).filter(Boolean));
}

// Baja el recibo actual de una cuenta y lo guarda en `destino`.
async function descargarRecibo(page, cuenta, destino) {
  // Despues de una descarga la pantalla queda lista: no hace falta recargarla.
  if (!page.url().startsWith(URL_PYP) || !(await page.$(SEL_CUENTAS))) await abrirPyp(page);
  const disponibles = await cuentasPyp(page);
  if (!disponibles.includes(cuenta)) {
    throw new ArquiError('no_vinculada',
      `La cuenta ${cuenta} no esta en el desplegable de Publicidad y Propaganda de este usuario de ARQui ` +
      '(no esta vinculada). No se agrega sola: hay que vincularla a mano con el "+" si corresponde.');
  }
  await page.selectOption(SEL_CUENTAS, cuenta);
  const elegida = await page.$eval(SEL_CUENTAS, (s) => s.value);
  if (elegida !== cuenta) throw new ArquiError('pantalla', `El desplegable quedo en ${elegida}, no en ${cuenta}.`);

  let download;
  try {
    [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 45000 }),
      page.click('#ContentPlaceHolder1_Button1'),
    ]);
  } catch {
    if (/Login_arqui/i.test(page.url())) throw new ArquiError('sesion', 'La sesion de ARQui vencio.');
    const texto = await textoVisible(page);
    if (/Forbidden/i.test(texto)) throw errBloqueo();
    const msg = await page.$eval('#ContentPlaceHolder1_msg', (e) => e.innerText.trim()).catch(() => '');
    throw new ArquiError('sin_recibo', `ARQui no devolvio recibo para la cuenta ${cuenta}: ${msg || texto.replace(/.*Motores y calderas/, '').slice(0, 200)}`);
  }
  await download.saveAs(destino);
  return destino;
}

module.exports = { prepararPagina, login, abrirPyp, cuentasPyp, descargarRecibo, ArquiError, URL_LOGIN, URL_PYP };
