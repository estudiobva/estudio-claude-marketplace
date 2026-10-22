// prestamos/nacion.js — BNA+ (digital.bna.com.ar). Un solo login para todas las
// sociedades: se cambia de "Ambiente" desde el menu del usuario.
//
// Relevado el 08/10/2026:
//   - #loggedUserMenu -> "Ambientes" -> lista de razones sociales (texto exacto).
//   - Menu superior "Préstamos" -> /loans. Sin prestamos muestra solo "Conocé
//     nuestras diferentes líneas de Créditos" + "Simular un préstamo".
//   - NUNCA usar page.goto adentro: BNA+ guarda la sesion en memoria y una
//     recarga la corta ("Tu sesión ha expirado"). Todo por clics.

const SIN_PRESTAMOS = /conoce nuestras diferentes lineas de creditos|no (tenes|posee|registra|hay) prestamos/;
const sinAcentos = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

async function abrirAmbientes(page) {
  const amb = page.getByText('Ambientes', { exact: true }).first();
  if (!(await amb.isVisible().catch(() => false))) {
    await page.locator('#loggedUserMenu').click();
    await page.waitForTimeout(1200);
  }
  await amb.click();
  await page.waitForTimeout(1500);
}

async function empresas(page) {
  await page.waitForTimeout(3000);
  await abrirAmbientes(page);
  const t = await page.innerText('body');
  const m = t.split(/\nAmbientes\n/)[1];
  if (!m) throw new Error('NACION: no pude leer la lista de Ambientes.');
  const lista = m.split(/\nSalir\b/)[0].split('\n').map(s => s.trim()).filter(Boolean);
  await page.keyboard.press('Escape').catch(() => {});
  return lista;
}

async function entrar(page, nombre) {
  await abrirAmbientes(page);
  await page.getByText(nombre, { exact: true }).first().click();
  // Al cambiar de ambiente vuelve al inicio con la razon social en el saludo.
  for (let i = 0; i < 30; i++) {
    await page.waitForTimeout(700);
    const t = await page.innerText('body').catch(() => '');
    if (/Mis cuentas/i.test(t) && t.includes(nombre)) break;
  }
  await page.waitForTimeout(1500);
}

async function irAPrestamos(page) {
  const link = page.locator('a[href$="/loans"], a[href*="/loans?"]').first();
  if (await link.isVisible().catch(() => false)) await link.click();
  else await page.getByText('Préstamos', { exact: true }).first().click();
  await page.waitForURL(/\/loans/, { timeout: 20000 }).catch(() => {});
  await page.waitForTimeout(4000);
}

async function listar(page) {
  await irAPrestamos(page);
  const texto = await page.innerText('body').catch(() => '');
  const sin = SIN_PRESTAMOS.test(sinAcentos(texto)) && !/cuota|capital|saldo|vencimiento/i.test(sinAcentos(texto.split('Préstamos').slice(2).join(' ')));
  // Con prestamos: falta relevar la pantalla (correr con --relevar).
  return { texto, prestamos: sin ? [] : null };
}

async function salir(page) {
  try {
    if (!(await page.getByText('Salir', { exact: true }).first().isVisible().catch(() => false))) {
      await page.locator('#loggedUserMenu').click();
      await page.waitForTimeout(1000);
    }
    await page.getByText('Salir', { exact: true }).first().click({ timeout: 5000 });
    await page.waitForTimeout(2500);
    const confirmar = page.locator('button:visible:has-text("Salir"), button:visible:has-text("Aceptar")').first();
    if (await confirmar.isVisible().catch(() => false)) await confirmar.click().catch(() => {});
    await page.waitForTimeout(1500);
  } catch { /* la sesion igual vence sola en minutos */ }
}

module.exports = { banco: 'NACION', sesion: 'unica', empresas, entrar, listar, salir };
