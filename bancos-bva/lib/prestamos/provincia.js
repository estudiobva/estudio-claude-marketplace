// prestamos/provincia.js — BIP Empresas (Banca Internet Provincia). El usuario
// del estudio tiene una sola empresa: no hay selector.
//
// Relevado el 08/10/2026:
//   - "Mi perfil" despliega "<RAZON SOCIAL> - " y "Cerrar Sesión".
//   - Menu Préstamos (#prestamos_prstamos_Button_0) -> "Mis Préstamos"
//     (#consultaPrestamos_misPrstamos_Button_1) -> /spa/prestamos: una tarjeta por
//     prestamo (div.box.button-container-child.cursor-pointer) con "Importe
//     original", "Vencimiento" y "Deuda a la fecha".
//   - Tarjeta -> /spa/prestamos/<uuid>: "Detalle Préstamo ..." con numero
//     (NNNN-NNNNNNN/N), fecha de otorgamiento, importe, TNA, deuda a la fecha,
//     proximo pago, vencimiento, forma de pago, ultima cuota paga.
//   - "Ver histórico" -> cuotas Todas / Paga / Pendientes ("Pendientes" vino vacio).
//   - "Más opciones" (#contextMenuButton) -> "Descargar contrato mutuo"
//     (#descargarContratoMutuo_descargarContratoMutuo_Button_0): el 08/10/2026 el
//     servicio devolvio 500 ("Error al invocar servicio core. DIR0004E").
//   El banco NO muestra un cronograma de cuotas a vencer: como cuadro se guarda
//   el detalle + historico en PDF y, si el banco lo entrega, el contrato mutuo.
//   - Todo por clics: es una SPA.

const sinAcentos = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
const TARJETA = 'div.box.button-container-child.cursor-pointer';

async function cerrarPopup(page) {
  const b = page.locator("#modal-popup_content-modal-header button[aria-label='Cerrar.']");
  if (await b.isVisible().catch(() => false)) await b.click().catch(() => {});
}

async function empresas(page) {
  // Al entrar puede tardar el inicio y aparecer un aviso: se reintenta unos segundos.
  for (let i = 0; i < 8; i++) {
    await page.waitForTimeout(2000);
    await cerrarPopup(page);
    const t = await page.innerText('body').catch(() => '');
    const m = t.match(/Mi perfil\s*\n\s*([^\n]+?)\s+-\s*\n/);
    if (m) {
      if (/Mi perfil\s*\n\s*[^\n]+? -\s*\n/.test(t)) {
        await page.getByText('Mi perfil', { exact: true }).first().click().catch(() => {});   // cierra el desplegable
      }
      return [m[1].trim()];
    }
    await page.getByText('Mi perfil', { exact: true }).first().click().catch(() => {});
  }
  throw new Error('PROVINCIA: no pude leer la razon social en "Mi perfil".');
}

async function entrar() { /* una sola empresa */ }

async function irAMisPrestamos(page) {
  await cerrarPopup(page);
  const mis = page.locator('#consultaPrestamos_misPrstamos_Button_1');
  if (!(await mis.isVisible().catch(() => false))) {
    await page.click('#prestamos_prstamos_Button_0');
    await page.waitForTimeout(1000);
  }
  await mis.click();
  for (let i = 0; i < 30; i++) {
    await page.waitForTimeout(700);
    const t = sinAcentos(await page.innerText('body').catch(() => ''));
    if (/mis prestamos/.test(t) && ((await page.locator(TARJETA).count()) || /no (tenes|posee|hay|registra)/.test(t))) break;
  }
  await page.waitForTimeout(1500);
}

const montoDe = (s) => (s || '').replace(/[^\d,.-]/g, '');

async function listar(page) {
  await irAMisPrestamos(page);
  const texto = await page.innerText('body').catch(() => '');
  const n = await page.locator(TARJETA).count();
  const prestamos = [];
  for (let i = 0; i < n; i++) {
    const t = (await page.locator(TARJETA).nth(i).innerText()).replace(/\n\s*\n/g, '\n');
    prestamos.push({ i, id: `tarjeta-${i + 1}`, descripcion: t.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 160) });
  }
  if (!n && !/no (tenes|posee|hay|registra)/.test(sinAcentos(texto))) return { texto, prestamos: null };
  return { texto, prestamos };
}

// Datos del detalle: "Etiqueta: valor" en una linea (la version accesible del banco).
function datosDetalle(texto) {
  const d = {};
  const linea = (rx) => ((texto.match(rx) || [])[1] || '').trim();
  d.numero = linea(/Detalle Pr[eé]stamo[^\n]*\n\s*([\d\-\/]+)/i);
  d.producto = linea(/Producto: ([^\n]+)/);
  d.fecha_otorgamiento = linea(/Fecha de otorgamiento: ([^\n]+)/);
  d.importe = linea(/Importe: ([^\n]+)/);
  d.tna = linea(/Tasa Nominal Anual: ([^\n]+)/);
  d.deuda_a_la_fecha = linea(/Deuda a la fecha: ([^\n]+)/);
  d.proximo_pago = linea(/Próximo pago: ([^\n]+)/);
  d.vencimiento = linea(/Vencimiento: ([^\n]+)/);
  d.forma_de_pago = linea(/Forma de pago: ([^\n]+)/);
  d.ultima_cuota_paga = montoDe(linea(/Última cuota paga:\s*\$?\s*Pesos\s*([\d.,]+)/));
  return d;
}

// PDF de la pantalla (solo anda oculto); con la ventana visible, captura PNG.
async function foto(page, guardar, nombre, tipo) {
  try { await guardar(await page.pdf({ format: 'A4', printBackground: true }), `${nombre}.pdf`, tipo); }
  catch { await guardar(await page.screenshot({ fullPage: true }), `${nombre}.png`, tipo); }
}

async function cuadro(page, prestamo, { guardar, aviso }) {
  await irAMisPrestamos(page);
  await page.locator(TARJETA).nth(prestamo.i).click();
  await page.waitForURL(/\/spa\/prestamos\/[0-9a-f-]{20,}/, { timeout: 30000 });
  await page.waitForTimeout(4000);
  const texto = await page.innerText('body');
  const datos = datosDetalle(texto);
  if (datos.numero) prestamo.id = datos.numero;
  Object.assign(prestamo, { datos });
  await foto(page, guardar, 'detalle', 'detalle');

  // Historico de cuotas (todas).
  await page.getByRole('button', { name: 'Ver histórico' }).first().click();
  await page.waitForURL(/\/historico/, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(4000);
  await foto(page, guardar, 'historico', 'cuotas');

  // Contrato mutuo (si el banco lo entrega).
  if (!(await page.locator('#descargarContratoMutuo_descargarContratoMutuo_Button_0').isVisible().catch(() => false))) {
    await page.click('#contextMenuButton').catch(() => {});
    await page.waitForTimeout(1200);
  }
  const descarga = page.waitForEvent('download', { timeout: 30000 }).catch(() => null);
  const pdfResp = page.context().waitForEvent('response', { timeout: 30000,
    predicate: (r) => /descargar\/mutuo/.test(r.url()) }).catch(() => null);
  await page.click('#descargarContratoMutuo_descargarContratoMutuo_Button_0', { timeout: 10000 }).catch(() => {});
  const [d, r] = await Promise.all([descarga, pdfResp]);
  if (d) {
    const ruta = await d.path();
    await guardar(require('fs').readFileSync(ruta), d.suggestedFilename(), 'contrato');
  } else {
    const estado = r ? r.status() : 'sin respuesta';
    aviso(`contrato mutuo no disponible (el banco respondio ${estado})`);
    if (/\/spa\/error/.test(page.url())) await page.getByText('Ir al inicio').first().click().catch(() => {});
  }
}

async function salir(page) {
  try {
    await page.getByText('Mi perfil', { exact: true }).first().click({ timeout: 5000 });
    await page.waitForTimeout(1000);
    await page.getByText('Cerrar Sesión', { exact: true }).first().click({ timeout: 5000 });
    await page.waitForTimeout(3000);
  } catch {
    await page.getByText('Cerrar sesión', { exact: true }).first().click({ timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(3000);
  }
}

module.exports = { banco: 'PROVINCIA', sesion: 'unica', empresas, entrar, listar, cuadro, salir };
