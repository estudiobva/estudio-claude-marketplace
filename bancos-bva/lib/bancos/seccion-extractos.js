// seccion-extractos.js — despues del login deja cada banco SIEMPRE en el mismo
// lugar: la empresa de la sociedad pedida y la seccion de resumenes/extractos.
// Asi quien sigue (Claude con scripts/control.js, o la persona) arranca siempre
// de la misma pantalla y no busca los extractos por otra parte del banco
// (Movimientos, Ultimos movimientos) ni en otra empresa del mismo usuario.
//
// La eleccion de empresa reusa la navegacion de lib/prestamos/<banco>.js y el
// emparejamiento nombre del banco -> CUIT de lib/prestamos/sociedades.js: si la
// sociedad no aparece (o aparece mas de una vez) corta, nunca entra a otra.
//
// Relevado en las corridas reales del 28-29/09/2026 y 28/10/2026:
//   MACRO      empresa -> menu de la cuenta -> "Resúmenes" (#menu_resumenes)
//   NACION     Ambiente -> cuenta -> "Más opciones" -> "Extracto de cuenta"
//   PATAGONIA  empresa -> "Resúmenes electrónicos" (#resumenesElectronicosCuentas)
//   PROVINCIA  (una sola empresa) -> "Extractos electrónicos"
// Sin ruta fija todavia: GALICIA, HIPOTECARIO, COMAFI, SANTANDER, BBVA.

const prest = require('../prestamos');
const socs = require('../prestamos/sociedades');

const sinAcentos = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
const error = (code, msg) => Object.assign(new Error(msg), { code });

async function esperarTexto(page, rx, ms = 25000) {
  for (let t = 0; t < ms; t += 700) {
    if (rx.test(sinAcentos(await page.innerText('body').catch(() => '')))) return true;
    await page.waitForTimeout(700);
  }
  return false;
}

// Entra a la empresa del banco que corresponde a la sociedad (por CUIT).
async function elegirEmpresa(banco, page, soc) {
  const mod = prest.MODULOS[banco];
  const lista = (await mod.empresas(page)).map(e => (typeof e === 'string' ? { nombre: e } : e));
  const hits = lista.filter(e => { const s = socs.emparejar(e.nombre, { cuit: e.cuit }); return s && s.cuit === soc.cuit; });
  if (hits.length !== 1) {
    throw error('empresa_no_encontrada', hits.length
      ? `${banco}: ${hits.length} empresas del banco corresponden a ${soc.nombre}: ${hits.map(h => h.nombre).join(' | ')}.`
      : `${banco}: ${soc.nombre} (${soc.cuit}) no esta entre las ${lista.length} empresas de este usuario.`);
  }
  if (lista.length > 1) await mod.entrar(page, hits[0].nombre);
  return hits[0].nombre;
}

const RUTAS = {
  MACRO: {
    seccion: 'Resúmenes de la cuenta (menú de la cuenta → Resúmenes)',
    async ir(page) {
      // El menu de cada cuenta esta oculto hasta el hover: se dispara por JS.
      const ok = await page.evaluate(() => {
        const li = document.querySelector('[id^="menuTableWidget"] #menu_resumenes');
        if (li) li.click();
        return !!li;
      });
      if (!ok) throw error('seccion_no_encontrada', 'MACRO: no encuentro "Resúmenes" en el menú de las cuentas.');
      await esperarTexto(page, /resumen/);
    },
  },
  NACION: {
    seccion: 'Extracto de cuenta (cuenta → Más opciones → Extracto de cuenta)',
    async ir(page) {
      // BNA+ guarda la sesion en memoria: todo por clics, nunca page.goto.
      const cuenta = page.getByText(/cuenta corriente/i).first();
      await cuenta.waitFor({ state: 'visible', timeout: 25000 });
      await cuenta.click();
      await page.getByText('Más opciones', { exact: true }).first().click({ timeout: 20000 });
      await page.locator('[id="account.last.summary.title.label"]').first().click({ timeout: 15000 });
      await esperarTexto(page, /extracto/);
    },
  },
  PATAGONIA: {
    seccion: 'Resúmenes electrónicos (Cuentas → Resúmenes electrónicos)',
    async ir(page) {
      const ok = await page.evaluate(() => { const a = document.getElementById('resumenesElectronicosCuentas'); if (a) a.click(); return !!a; });
      if (!ok) throw error('seccion_no_encontrada', 'PATAGONIA: no encuentro "Resúmenes electrónicos" en el menú Cuentas.');
      await page.locator('#collectionTableCuentas_DETAIL, [id^="collectionTableCuentas"]').first()
        .waitFor({ state: 'attached', timeout: 25000 }).catch(() => {});
    },
  },
  PROVINCIA: {
    // Relevado el 28/10/2026: Extractos Electronicos esta adentro de cada cuenta
    // (Mis Cuentas -> Ver movimientos -> boton "Extractos Electrónicos"). Se entra
    // a la cuenta corriente: las cajas de ahorro pueden tener extractos diarios.
    seccion: 'Extractos electrónicos de la cuenta corriente (Cuentas → Mis Cuentas → Ver movimientos → Extractos Electrónicos)',
    async ir(page) {
      const b = page.locator("#modal-popup_content-modal-header button[aria-label='Cerrar.']");
      if (await b.isVisible().catch(() => false)) await b.click().catch(() => {});
      await page.evaluate(() => { const x = document.getElementById('consultaCuentas_misCuentas_Button_1'); if (x) x.click(); });
      await page.waitForURL(/\/spa\/cuentas$/, { timeout: 20000 }).catch(() => {});
      await page.getByText('Ver movimientos', { exact: true }).first().waitFor({ state: 'visible', timeout: 20000 });
      const ok = await page.evaluate(() => {
        const bs = [...document.querySelectorAll('a, button, [role="button"]')].filter(x => x.innerText.trim() === 'Ver movimientos');
        const tarjeta = (x) => { let el = x; for (let i = 0; i < 8 && el; i++) { el = el.parentElement; if (el && /\b(CC|CA) \$/.test(el.innerText)) return el.innerText; } return ''; };
        const b = bs.find(x => /\bCC \$/.test(tarjeta(x).split(/\bCA \$/)[0])) || bs[0];
        if (b) b.click();
        return !!b;
      });
      if (!ok) throw error('seccion_no_encontrada', 'PROVINCIA: no encuentro "Ver movimientos" en Mis Cuentas.');
      const ext = page.locator('#extractosElectronicos_extractosElectrnicos_Button_0');
      await ext.waitFor({ state: 'attached', timeout: 20000 })
        .catch(() => { throw error('seccion_no_encontrada', 'PROVINCIA: no aparece "Extractos Electrónicos" en la cuenta.'); });
      await page.evaluate(() => document.getElementById('extractosElectronicos_extractosElectrnicos_Button_0').click());
      await page.locator('[id^="pdf-"]').first().waitFor({ state: 'attached', timeout: 25000 }).catch(() => {});
    },
  },
};

const tieneRuta = (banco) => !!RUTAS[banco];

// Devuelve { empresa, seccion, aviso }. Si no puede entrar a la empresa de la
// sociedad lanza (code empresa_no_encontrada u otro): ese banco no sigue. Si ya
// esta en la empresa pero no llega a la seccion, devuelve `aviso` y se sigue a
// mano desde ahi.
async function irAExtractos(banco, page, soc) {
  const ruta = RUTAS[banco];
  if (!ruta) return null;
  const empresa = await elegirEmpresa(banco, page, soc);
  try { await ruta.ir(page); }
  catch (e) { return { empresa, seccion: ruta.seccion, aviso: e.message.split('\n')[0] }; }
  return { empresa, seccion: ruta.seccion };
}

module.exports = { irAExtractos, tieneRuta, RUTAS };
