// prestamos/patagonia.js — Patagonia eBank Empresas. Un login para las 5
// sociedades del usuario; el selector muestra razon social y CUIT.
//
// Relevado el 08/10/2026:
//   - selector: radios name=encodedcollection_collectionTablePerfiles (una fila
//     por empresa: "RAZON SOCIAL  CUIT 30000000007  Operador") · #actionButtonContinuar.
//   - Desde adentro: "Cambiar de empresa" (#Cambiardeempresa) vuelve al selector.
//   - Préstamos (#Prestamos) -> "Consulta de préstamos" (#Consultaprestamos):
//     tabla Fecha de alta | Número | Capital original | Moneda | Capital adeudado |
//     Cuotas vencidas/impagas | Cuotas a vencer. Sin prestamos: "No existe
//     información disponible para la consulta realizada."
//   - Los links del menu estan ocultos hasta el hover: se disparan con click() por JS.
//   - Salir: #widgetLogoutBtn (closeSesion()).

const RADIO = 'input[type=radio][name="encodedcollection_collectionTablePerfiles"]';
const sinAcentos = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
const clicJS = (page, id) => page.evaluate((i) => document.getElementById(i).click(), id);

async function filas(page) {
  await page.locator(RADIO).first().waitFor({ state: 'attached', timeout: 30000 });
  return page.evaluate((sel) => [...document.querySelectorAll(sel)].map((r, i) => {
    const celdas = (r.closest('tr').innerText || '').split('\t').map(s => s.trim()).filter(Boolean);
    const cuit = ((r.closest('tr').innerText || '').match(/CUIT\s*(\d{11})/) || [])[1] || null;
    return { i, nombre: celdas[0], cuit };
  }), RADIO);
}

async function empresas(page) {
  return (await filas(page)).map(f => ({ nombre: f.nombre, cuit: f.cuit }));
}

async function entrar(page, nombre) {
  if (!(await page.locator(RADIO).count())) {
    await clicJS(page, 'Cambiardeempresa');
    await page.waitForTimeout(3000);
  }
  const f = (await filas(page)).find(x => x.nombre === nombre);
  if (!f) throw new Error(`PATAGONIA: "${nombre}" no esta en el selector de empresas.`);
  await page.locator(RADIO).nth(f.i).check();
  await page.click('#actionButtonContinuar');
  await page.locator('#Prestamos').waitFor({ state: 'attached', timeout: 30000 });
  await page.waitForTimeout(2000);
}

async function listar(page) {
  await clicJS(page, 'Consultaprestamos');
  for (let i = 0; i < 30; i++) {
    await page.waitForTimeout(700);
    const t = sinAcentos(await page.innerText('body').catch(() => ''));
    if (/capital adeudado|no tiene contratada esta funcionalidad/.test(t)) break;
  }
  await page.waitForTimeout(1500);
  const texto = await page.innerText('body').catch(() => '');
  // Visto el 08/10/2026 en una sociedad: "Su empresa no tiene contratada esta funcionalidad".
  if (/no tiene contratada esta funcionalidad/.test(sinAcentos(texto))) {
    throw Object.assign(new Error('PATAGONIA: la empresa no tiene contratada la consulta de prestamos en eBank.'), { code: 'sin_funcionalidad' });
  }
  const sin = /no existe informacion disponible/.test(sinAcentos(texto));
  return { texto, prestamos: sin ? [] : null };
}

async function salir(page) {
  try {
    await page.evaluate(() => { const b = document.getElementById('widgetLogoutBtn') || document.getElementById('actionButtonSalir'); if (b) b.click(); });
    await page.waitForTimeout(3000);
  } catch { /* ya afuera */ }
}

module.exports = { banco: 'PATAGONIA', sesion: 'unica', empresas, entrar, listar, salir };
