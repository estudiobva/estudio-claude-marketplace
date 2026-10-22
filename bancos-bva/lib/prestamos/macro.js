// prestamos/macro.js — Banca Internet Empresas. Un login muestra el selector de
// las ~55 empresas. SALIR adentro de una empresa vuelve al selector (no cierra
// la sesion); SALIR en el selector cierra la sesion. Un solo login para todas.
//
// Relevado el 08/10/2026:
//   - selector: botones #section0_repeatN_actionButtonVerify (el texto queda tapado).
//   - "Financiaciones" (#financiaciones) -> "Préstamos vigentes, cuotas y
//     comprobantes" (#consultaDetallePrestamo) -> "Listado de Préstamos".
//     Sin prestamos: "No hay préstamos para mostrar".
//   - SALIR (#widgetLogoutBtn): desde una empresa -> selector; desde el selector
//     -> fin de la sesion. Si una sesion queda abierta (por ejemplo, se cerro el
//     navegador en el selector), el banco no deja volver a entrar por 10 minutos
//     ("ya contás con una sesión activa").

const sinAcentos = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

async function botones(page) {
  await page.locator('[id^="section0_repeat"][id$="_actionButtonVerify"]').first()
    .waitFor({ state: 'attached', timeout: 30000 });
  return page.evaluate(() => [...document.querySelectorAll('[id^="section0_repeat"][id$="_actionButtonVerify"]')].map(b => {
    let el = b, nombre = '';
    for (let i = 0; i < 6 && el && !nombre; i++) {
      el = el.parentElement;
      const t = (el && el.innerText || '').trim();
      if (t && t.split('\n').filter(Boolean).length <= 3) nombre = t.split('\n').map(s => s.trim()).filter(Boolean)[0];
    }
    return { id: b.id, nombre };
  }));
}

async function empresas(page) {
  return (await botones(page)).map(b => b.nombre).filter(Boolean);
}

const SELECTOR = '[id^="section0_repeat"][id$="_actionButtonVerify"]';
const enSelector = async (page) => (await page.locator(SELECTOR).count().catch(() => 0)) > 0;

// SALIR una vez y esperar a que cambie de pantalla.
async function tocarSalir(page) {
  const antes = await enSelector(page);
  // En el selector hay mas de un #widgetLogoutBtn (uno oculto): se toca el visible por JS.
  const ok = await page.evaluate(() => {
    const bs = [...document.querySelectorAll('#widgetLogoutBtn')];
    const b = bs.find(x => x.offsetParent) || bs[0];
    if (b) b.click();
    return !!b;
  });
  if (!ok) throw new Error('MACRO: no encuentro el boton SALIR.');
  for (let i = 0; i < 30; i++) {
    await page.waitForTimeout(500);
    if (antes !== await enSelector(page)) break;
  }
  await page.waitForTimeout(1500);
}

async function entrar(page, nombre) {
  // Si quedo adentro de otra empresa, SALIR la devuelve al selector.
  if (!(await enSelector(page))) await tocarSalir(page);
  const b = (await botones(page)).find(x => x.nombre === nombre);
  if (!b) throw new Error(`MACRO: "${nombre}" no esta en el selector de empresas.`);
  await page.locator(`#${b.id}`).click();
  await page.locator('#financiaciones').waitFor({ state: 'visible', timeout: 30000 });
  await page.waitForTimeout(1500);
}

async function irAPrestamos(page) {
  await page.locator('#financiaciones').click();
  await page.waitForTimeout(1200);
  await page.locator('#consultaDetallePrestamo').click();
  for (let i = 0; i < 30; i++) {
    await page.waitForTimeout(700);
    const t = sinAcentos(await page.innerText('body').catch(() => ''));
    if ((/listado de prestamos/.test(t) && /no hay prestamos|capital adeudado/.test(t)) || SIN_SERVICIO.test(t)) break;
  }
  await page.waitForTimeout(1500);
}

// "Listado de Préstamos": Número | Préstamo | Capital Adeudado | Próxima cuota | Próximo vencimiento,
// una fila por prestamo, cada una con su menu (menuTableWidget<id>): Ver detalle /
// Plan de Cuotas (#menu_cuotas) / Comprobantes de pago.
const SIN_SERVICIO = /no se puede establecer conexion con el servidor/;

async function listar(page) {
  await irAPrestamos(page);
  let texto = await page.innerText('body').catch(() => '');
  // De noche el banco a veces no responde ("No se puede establecer conexión con el
  // Servidor"): se reintenta dos veces antes de darlo por caido.
  for (let k = 0; k < 2 && SIN_SERVICIO.test(sinAcentos(texto)); k++) {
    await page.waitForTimeout(8000);
    await irAPrestamos(page);
    texto = await page.innerText('body').catch(() => '');
  }
  if (SIN_SERVICIO.test(sinAcentos(texto))) {
    throw Object.assign(new Error('MACRO: "No se puede establecer conexión con el Servidor" al abrir Préstamos.'), { code: 'banco_sin_servicio' });
  }
  if (/no hay prestamos para mostrar/.test(sinAcentos(texto))) return { texto, prestamos: [] };
  const prestamos = [];
  for (const l of texto.split('\n')) {
    const c = l.split('\t').map(x => x.trim());
    if (/^\d{6,}$/.test(c[0]) && c.length >= 5) {
      prestamos.push({ id: c[0], descripcion: c[1], i: prestamos.length,
        datos: { producto: c[1], capital_adeudado: c[2], proxima_cuota: c[3], proximo_pago: c[4].replace(/^Vence\s*/i, '') } });
    }
  }
  return { texto, prestamos: prestamos.length ? prestamos : null };
}

// Plan de cuotas = el cuadro de marcha. El banco no ofrece descarga: se guarda la
// pantalla en PDF y la tabla como planilla (CSV que abre Excel).
async function cuadro(page, prestamo, { guardar, tabla }) {
  // Cada plan se abre desde el listado (VOLVER no siempre deja el menu listo).
  if (!/listado de prestamos/.test(sinAcentos(await page.innerText('body').catch(() => '')))) await irAPrestamos(page);
  const ok = await page.evaluate(({ id, i }) => {
    const filas = [...document.querySelectorAll('[id^="menuTableWidget"]')];
    const m = filas.find(x => (x.closest('tr') || x.parentElement.parentElement).innerText.includes(id)) || filas[i];
    const li = m && m.querySelector('#menu_cuotas');
    if (li) li.click();
    return !!li;
  }, { id: prestamo.id, i: prestamo.i });
  if (!ok) throw new Error(`MACRO: no encuentro "Plan de Cuotas" del prestamo ${prestamo.id}.`);
  for (let k = 0; k < 40; k++) {
    await page.waitForTimeout(700);
    if (/num\. de cuota/.test(sinAcentos(await page.innerText('body').catch(() => '')))) break;
  }
  await page.waitForTimeout(1500);
  const texto = await page.innerText('body');
  const lineas = texto.split('\n');
  const iEnc = lineas.findIndex(l => /^Num\. de cuota\t/i.test(l.trim()));
  if (iEnc < 0) throw new Error(`MACRO: el plan de cuotas del prestamo ${prestamo.id} no cargo.`);
  const enc = lineas[iEnc].split('\t').map(x => x.trim()).filter(Boolean);
  const filas = [];
  for (const l of lineas.slice(iEnc + 1)) {
    const c = l.split('\t').map(x => x.trim());
    if (!/^\d+$/.test(c[0])) break;
    filas.push(c.slice(0, enc.length));
  }
  const pagas = (texto.match(/Cuotas\s*pagas\s*\n\s*(\d+\/\d+)/i) || [])[1];
  if (pagas) prestamo.datos = { ...(prestamo.datos || {}), cuotas_pagas: pagas };
  prestamo.datos = { ...(prestamo.datos || {}), cuotas: filas.length, vencimiento: (filas[filas.length - 1] || [])[1] || '' };
  tabla(enc, filas);
  try { await guardar(await page.pdf({ format: 'A4', landscape: true, printBackground: true }), 'plan.pdf', 'cuadro-marcha'); }
  catch { await guardar(await page.screenshot({ fullPage: true }), 'plan.png', 'cuadro-marcha'); }
  await irAPrestamos(page);
}

// Cierra la sesion de verdad: de la empresa al selector y del selector afuera.
async function salir(page) {
  for (let i = 0; i < 2; i++) {
    if (!(await page.locator('#widgetLogoutBtn').count().catch(() => 0))) return;
    const eraSelector = await enSelector(page);
    await tocarSalir(page).catch(() => {});
    if (eraSelector) return;
  }
}

module.exports = { banco: 'MACRO', sesion: 'unica', empresas, entrar, listar, cuadro, salir };
