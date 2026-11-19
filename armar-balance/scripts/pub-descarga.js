#!/usr/bin/env node
/**
 * pub-descarga.js — Baja de la Presentación Única de Balances (PUB) de ARCA el PDF de los
 * estados contables presentados (adjunto) de una o varias sociedades. Solo lectura:
 * abre el menú de la fila del período y usa "Descargar adjunto"; nunca "Rectificar" ni "Nueva presentación".
 *
 *   node scripts/pub-descarga.js --cuit=<CUIT> --periodo=2025
 *   node scripts/pub-descarga.js --cuit=<CUIT>,<CUIT> --periodo=2025 [--constancia]
 *
 * --login-alt=<cuit de sociedad>: si el apoderado de la planilla no la representa en PUB, reintenta con el login de esa
 *   sociedad (por defecto PUB_LOGIN_ALT del .env: el CUIT de una sociedad que entra con el apoderado que
 *   representa a más sociedades en PUB). --login-alt= vacío lo desactiva.
 * Guarda en [Sociedad]/04-Balances y EECC/<periodo>/EECC Legalizados/EECC_<cuit>_<periodo>_0.pdf (nombre que da ARCA)
 * (si ya hay uno idéntico no hace nada; si hay otro con el mismo nombre, guarda el de PUB como _PUB.pdf). Salida: JSON por stdout.
 */
const path = require('path'); const fs = require('fs');
const { chromium } = require('playwright');
const { launchOptions } = require('../lib/launch'); const { parseArgs, fail } = require('../lib/args');
const { resolveCredentials } = require('../lib/env'); const { login, abrirServicio } = require('../lib/arca-login');
const { carpetaSociedad } = require('../lib/rutas-bva');
const args = parseArgs();
const CUITS = String(args['cuit'] || '').split(',').map(s => s.replace(/\D/g, '')).filter(Boolean);
const PERIODO = String(args['periodo'] || new Date().getFullYear() - 1);
if (!CUITS.length) fail('Falta --cuit');
const log = (m) => process.stderr.write(`[pub] ${m}\n`);
require('../lib/env').loadEnv();
const ALT = String(args['login-alt'] === undefined ? (process.env.PUB_LOGIN_ALT || '') : args['login-alt']).replace(/\D/g, '');
const fmt = (c) => `${c.slice(0, 2)}-${c.slice(2, 10)}-${c.slice(10)}`;

async function unaSociedad(cuit, cred) {
  let cuitLogin, password;
  try { ({ cuitLogin, password } = cred || resolveCredentials({ cuit })); }
  catch (e) { return { cuit, estado: 'error', detalle: e.message.split('\n')[0] }; }
  let browser;
  try { browser = await chromium.launch(launchOptions()); }
  catch (e) { return { cuit, estado: 'error', detalle: 'No pude abrir el navegador (¿falta `npx playwright install chromium`?): ' + e.message.split('\n')[0] }; }
  const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
  const page = await context.newPage();
  try {
    await login(page, { cuitLogin, password });
    const pub = await abrirServicio(context, page, /[uú]nica de balances/i);
    await pub.waitForTimeout(4000);
    if (/seleccion-representado/.test(pub.url())) {
      await pub.getByText('Elegí la persona para operar').first().waitFor({ timeout: 20000 }).catch(() => {});
      await pub.waitForTimeout(2000);
      const op = pub.getByText(fmt(cuit), { exact: true }).first();
      if (!(await op.count())) {
        const lista = await pub.innerText('body').catch(() => '');
        return { cuit, estado: 'sin_representacion', detalle: `login ${cuitLogin}; ${(lista.match(/\d{2}-\d{8}-\d/g) || []).length} representados en la lista` };
      }
      await op.scrollIntoViewIfNeeded().catch(() => {}); await op.click();
    }
    await pub.waitForSelector('a[href="#/presentaciones"]', { timeout: 30000 });
    const cab = await pub.innerText('body');
    if (!cab.includes(fmt(cuit))) return { cuit, estado: 'otra_sociedad_en_pantalla' };
    await pub.locator('a[href="#/presentaciones"]').first().click();
    await pub.waitForTimeout(6000);
    // Filas: la del período con Tipo 'Estados contables'; si hay más de una (rectificativas), la de fecha de presentación más nueva.
    // AG Grid: datos en el contenedor central, botón de acciones en el anclado a la derecha (mismo row-index).
    await pub.locator('.ag-center-cols-container [role=row]').first().or(pub.getByText('NO SE ENCONTRARON REGISTROS')).waitFor({ timeout: 30000 });
    if (await pub.locator('text=NO SE ENCONTRARON REGISTROS').isVisible().catch(() => false)) return { cuit, estado: 'sin_presentacion', periodo: PERIODO, detalle: 'PUB sin ninguna presentación' };
    const filas = pub.locator('.ag-center-cols-container [role=row]');
    const n = await filas.count();
    let mejor = null;
    for (let i = 0; i < n; i++) {
      const t = (await filas.nth(i).innerText()).split(/\n|\t/).map(s => s.trim()).filter(Boolean);
      if (t[0] !== PERIODO || !/Estados contables/i.test(t.join(' '))) continue;
      const fechas = t.filter(x => /^\d{2}\/\d{2}\/\d{4}$/.test(x));
      const key = (fechas[fechas.length - 1] || '').split('/').reverse().join('');
      const idx = await filas.nth(i).getAttribute('row-index');
      if (!mejor || key > mejor.key) mejor = { idx, key, t };
    }
    if (!mejor) return { cuit, estado: 'sin_presentacion', periodo: PERIODO };
    const dir = path.join(carpetaSociedad(cuit), '04-Balances y EECC', PERIODO, 'EECC Legalizados');
    fs.mkdirSync(dir, { recursive: true });
    const res = { cuit, estado: 'ok', periodo: PERIODO, fila: mejor.t.join(' | '), archivos: [] };
    // Siempre baja de PUB (el nombre de la carpeta no garantiza el cierre: hay legalizados del año anterior guardados en
    // la carpeta del año de legalización). Si ya hay un archivo igual, no hace nada; si hay uno distinto con ese nombre,
    // guarda el de PUB como <nombre>_PUB.pdf y no pisa el existente.
    const bajar = async (texto, nombreDef) => {
      await pub.locator(`.ag-pinned-right-cols-container [row-index="${mejor.idx}"] button`).first().click();
      await pub.waitForTimeout(800);
      const dl = pub.waitForEvent('download', { timeout: 60000 });
      await pub.locator('a', { hasText: texto }).filter({ visible: true }).first().click();
      const d = await dl; const tmp = await d.path(); const nuevo = fs.readFileSync(tmp);
      const nombre = d.suggestedFilename() || nombreDef;
      const iguales = fs.readdirSync(dir).filter(f => f.toLowerCase().endsWith('.pdf')).find(f => { try { return fs.readFileSync(path.join(dir, f)).equals(nuevo); } catch { return false; } });
      if (iguales) { res.archivos.push({ archivo: path.join(dir, iguales), ya_estaba: true }); }
      else {
        let destino = path.join(dir, nombre);
        if (fs.existsSync(destino)) destino = destino.replace(/\.pdf$/i, '_PUB.pdf');
        fs.writeFileSync(destino, nuevo); res.archivos.push({ archivo: destino });
      }
      await pub.keyboard.press('Escape').catch(() => {});
    };
    await bajar('Descargar adjunto', `EECC_${cuit}_${PERIODO}_0.pdf`);
    if ('constancia' in args) await bajar('Descargar constancia', `PUB-CONSTANCIA EECC ${PERIODO}.pdf`);
    return res;
  } catch (e) {
    return { cuit, estado: 'error', detalle: e.message.split('\n')[0] };
  } finally { await browser.close(); }
}

(async () => {
  const out = [];
  for (const c of CUITS) {
    log(`${c}…`); let r = await unaSociedad(c);
    // Si el apoderado de la planilla no la representa en PUB, segundo intento con el login de otra sociedad (PUB_LOGIN_ALT)
    if (r.estado === 'sin_representacion' && ALT) {
      const cred = resolveCredentials({ cuit: ALT });
      if (cred.cuitLogin !== (() => { try { return resolveCredentials({ cuit: c }).cuitLogin; } catch { return null; } })()) { const r2 = await unaSociedad(c, cred); r2.intento = `login alternativo ${cred.cuitLogin} (antes: ${r.detalle})`; r = r2; }
    }
    log(`${c}: ${r.estado}`); out.push(r);
  }
  console.log(JSON.stringify(out, null, 1));
})();
