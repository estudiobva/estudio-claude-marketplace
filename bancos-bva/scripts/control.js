#!/usr/bin/env node
/**
 * control.js — maneja la ventana de un banco YA LOGUEADO por extractos-banco.js --control.
 *
 * Lo usa Claude para navegar hasta los resumenes y descargarlos; el que renombra y
 * archiva cada archivo sigue siendo extractos-banco.js. Se conecta al puerto local
 * (127.0.0.1) que abre --control, hace UNA accion y se desconecta sin cerrar nada.
 *
 *   node scripts/control.js paginas
 *   node scripts/control.js captura                      -> ~/bva-inspeccion/control-<fecha>.png
 *   node scripts/control.js elementos                    botones, links, campos y selects visibles
 *   node scripts/control.js texto
 *   node scripts/control.js clic --texto="Resúmenes"     (o --sel="css", --n=2 para el tercero)
 *   node scripts/control.js elegir --sel="#mes" --opcion="Octubre 2025"
 *   node scripts/control.js completar --sel="#desde" --valor="01/10/2025"
 *   node scripts/control.js esperar --ms=3000
 *   node scripts/control.js descargar --sel="#pdf0"      baja el archivo a ~/bva-inspeccion/bandeja/ (para revisarlo)
 *   node scripts/control.js descargar --sel="#pdf0" --banco=NACION --cuit=30000000007 --periodo=202608
 *                                                        ...y lo archiva como extractos-banco.js: mismo nombre,
 *                                                        misma carpeta, sin pisar, rechaza "PDF" que no son PDF
 *                                                        (--salida=..., o --unidad para la unidad compartida)
 *   node scripts/control.js listo                        toca "Listo" en el recuadro Estudio BVA
 *   node scripts/control.js responder --valor=202510     contesta el periodo que pide el recuadro
 *
 * Limites a proposito:
 *   - nunca escribe en campos de clave, token, PIN o codigo;
 *   - nunca navega fuera del sitio del banco que esta abierto;
 *   - no cierra la ventana ni la sesion.
 * Opciones comunes: --puerto=9333, --pagina=N (por defecto, la ultima abierta).
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { chromium } = require('playwright');
const { parseArgs } = require('../lib/args');
require('../lib/env').loadEnv();   // BVA_SALIDAS_PATH / BVA_UNIDAD_PATH pueden venir del .env

const [cmd] = process.argv.slice(2).filter(a => !a.startsWith('--'));
const args = parseArgs(process.argv.slice(2).filter(a => a.startsWith('--')));
const PUERTO = Number(args['puerto'] || 9333);
const out = (o) => { console.log(JSON.stringify(o, null, 2)); process.exit(o.error ? 1 : 0); };

const SENSIBLE = /pass|clave|contrase|token|pin\b|otp|codigo|c[oó]digo|secret|cvv/i;

async function paginaActual(browser) {
  // Las descargas de algunos bancos (Macro) dejan una pestana sin URL: no cuenta.
  const pages = browser.contexts().flatMap(c => c.pages()).filter(p => p.url() && p.url() !== 'about:blank');
  if (!pages.length) throw new Error('No hay ninguna pagina abierta en la ventana del banco.');
  const i = args['pagina'] != null ? Number(args['pagina']) : pages.length - 1;
  if (!pages[i]) throw new Error(`No existe la pagina ${i} (hay ${pages.length}).`);
  return { page: pages[i], pages };
}

// Archivo de descarga mas nuevo (creado despues de t0) en las carpetas temporales
// de Playwright, esperando a que deje de crecer.
async function buscarArtefacto(t0) {
  const tmp = os.tmpdir();
  const espera = (ms) => new Promise(r => setTimeout(r, ms));
  for (let i = 0; i < 40; i++) {
    const cands = [];
    for (const dd of fs.readdirSync(tmp).filter(n => n.startsWith('playwright-artifacts-'))) {
      for (const f of fs.readdirSync(path.join(tmp, dd))) {
        const fp = path.join(tmp, dd, f);
        try { const st = fs.statSync(fp); if (st.isFile() && st.mtimeMs >= t0 - 1000) cands.push({ fp, t: st.mtimeMs, size: st.size }); } catch { /* en uso */ }
      }
    }
    cands.sort((a, b) => b.t - a.t);
    if (cands.length && cands[0].size > 0) {
      await espera(1200);
      if (fs.statSync(cands[0].fp).size === cands[0].size) return cands[0].fp;
    } else await espera(500);
  }
  return null;
}

async function objetivo(page) {
  if (args['sel']) return page.locator(args['sel']).nth(Number(args['n'] || 0));
  if (args['texto']) {
    // Primero botones y links con ese texto; si no hay, el texto mas especifico
    // (getByText apunta al elemento mas chico que lo contiene, no a un contenedor).
    const t = args['texto'];
    const n = Number(args['n'] || 0);
    const clickeables = page.locator('a:visible, button:visible, [role="button"]:visible, ' +
      '[role="menuitem"]:visible, [role="tab"]:visible, [role="link"]:visible').filter({ hasText: t });
    if (await clickeables.count()) return clickeables.nth(n);
    return page.getByText(t, { exact: false }).filter({ visible: true }).nth(n);
  }
  throw new Error('Falta --sel o --texto.');
}

(async () => {
  let browser;
  try { browser = await chromium.connectOverCDP(`http://127.0.0.1:${PUERTO}`); }
  catch { out({ error: `No hay ninguna ventana en control en el puerto ${PUERTO}. ¿Corriste extractos-banco.js con --control?` }); }

  try {
    const { page, pages } = await paginaActual(browser);
    const origen = new URL(page.url()).origin;

    switch (cmd) {
      case 'paginas':
        return out({ paginas: await Promise.all(pages.map(async (p, i) => ({ i, url: p.url(), titulo: await p.title().catch(() => '') }))) });

      case 'captura': {
        const dir = path.join(os.homedir(), 'bva-inspeccion');
        fs.mkdirSync(dir, { recursive: true });
        const f = args['archivo'] || path.join(dir, `control-${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.png`);
        await page.screenshot({ path: f, fullPage: 'completa' in args });
        return out({ captura: f, url: page.url() });
      }

      case 'texto':
        return out({ url: page.url(), texto: (await page.innerText('body')).slice(0, Number(args['max'] || 6000)) });

      case 'elementos': {
        const els = await page.evaluate(() => {
          const vis = (e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(e).visibility !== 'hidden'; };
          const txt = (e) => (e.innerText || e.getAttribute('aria-label') || e.title || '').trim().replace(/\s+/g, ' ').slice(0, 70);
          const q = (s) => [...document.querySelectorAll(s)].filter(vis).filter(e => !e.closest('#bva-barra'));
          return {
            clickeables: q('a, button, [role=button], [role=menuitem], [role=tab], [role=link]').map(e => ({
              tag: e.tagName, texto: txt(e), id: e.id || undefined, href: e.getAttribute('href') || undefined,
              deshabilitado: e.disabled || e.getAttribute('aria-disabled') === 'true' || undefined })).filter(x => x.texto || x.id),
            campos: q('input, textarea').filter(e => !['hidden'].includes(e.type)).map(e => ({
              tipo: e.type, id: e.id || undefined, name: e.name || undefined, placeholder: e.placeholder || undefined,
              label: (e.labels && e.labels[0] && e.labels[0].innerText.trim()) || undefined, completo: !!e.value })),
            selects: q('select').map(e => ({ id: e.id || undefined, name: e.name || undefined,
              opciones: [...e.options].map(o => o.text.trim()).slice(0, 40), elegida: e.options[e.selectedIndex] && e.options[e.selectedIndex].text })),
          };
        });
        return out({ url: page.url(), ...els });
      }

      case 'clic': {
        // El recuadro "Estudio BVA" puede tapar menus del banco: mientras se hace
        // el clic se lo vuelve transparente a los clics (pointer-events: none).
        await page.evaluate(() => { const b = document.getElementById('bva-barra'); if (b) b.style.pointerEvents = 'none'; }).catch(() => {});
        const el = await objetivo(page);   // async: decide entre clickeables y texto
        await el.first().scrollIntoViewIfNeeded().catch(() => {});
        await el.first().click({ timeout: Number(args['timeout'] || 10000) });
        await page.waitForTimeout(Number(args['espera'] || 2000));
        const actual = (() => { try { return new URL(page.url()).origin; } catch { return origen; } })();
        return out({ ok: true, url: page.url(), ...(actual !== origen ? { aviso: `la pagina salio a ${actual}` } : {}) });
      }

      case 'elegir': {
        if (!args['sel'] || args['opcion'] == null) throw new Error('Falta --sel y --opcion.');
        await page.locator(args['sel']).first().selectOption({ label: args['opcion'] });
        return out({ ok: true });
      }

      case 'completar': {
        if (!args['sel'] || args['valor'] == null) throw new Error('Falta --sel y --valor.');
        const loc = page.locator(args['sel']).first();
        const info = await loc.evaluate(e => ({ type: e.type, id: e.id, name: e.name, ph: e.placeholder,
          label: e.labels && e.labels[0] ? e.labels[0].innerText : '' }), null, { timeout: 5000 });
        if (info.type === 'password' || SENSIBLE.test([info.id, info.name, info.ph, info.label].join(' '))) {
          throw new Error('Ese campo parece de clave/token/codigo: control.js no escribe ahi. Lo completa la persona.');
        }
        await loc.fill(String(args['valor']));
        return out({ ok: true });
      }

      case 'esperar':
        await page.waitForTimeout(Number(args['ms'] || 2000));
        return out({ ok: true, url: page.url() });

      case 'descargar': {
        // Al conectarse por CDP el navegador entrega las descargas a ESTA conexion,
        // no a extractos-banco.js: por eso se reciben y se archivan aca.
        await page.evaluate(() => { const b = document.getElementById('bva-barra'); if (b) b.style.pointerEvents = 'none'; }).catch(() => {});
        const el = await objetivo(page);
        const t0 = Date.now();
        const [d] = await Promise.all([
          page.waitForEvent('download', { timeout: Number(args['timeout'] || 30000) }),
          el.first().click({ timeout: 10000 }),
        ]);
        const bandeja = path.join(os.homedir(), 'bva-inspeccion', 'bandeja');
        fs.mkdirSync(bandeja, { recursive: true });
        const sugerido = d.suggestedFilename();
        const local = path.join(bandeja, `${Date.now()}__${sugerido}`);
        try { await d.saveAs(local); }
        catch {
          // El navegador guarda el archivo en la carpeta temporal de extractos-banco.js
          // (playwright-artifacts-*), no en la de esta conexion: se toma de ahi.
          const encontrado = await buscarArtefacto(t0);
          if (!encontrado) throw new Error('El banco inicio la descarga pero no encontre el archivo.');
          fs.copyFileSync(encontrado, local);
        }
        if (!args['periodo']) return out({ ok: true, archivo: sugerido, bandeja: local });

        const ex = require('../lib/extractos');
        const p = ex.periodo(args['periodo']);
        if (!p) throw new Error('--periodo invalido (AAAAMM).');
        if (!args['banco'] || !args['cuit']) throw new Error('Para archivar hacen falta --banco y --cuit.');
        const soc = ex.sociedad({ cuit: args['cuit'] });
        const cb = 'unidad' in args
          ? ex.carpetaBanco(soc, args['banco'], args['carpeta'] ? { elegida: args['carpeta'] } : {})
          : ex.carpetaSalida(ex.salidaBase(args['salida']), soc, args['banco']);
        if (cb.candidatas) throw new Error(`Hay varias carpetas para ${args['banco']}: ${cb.candidatas.join(' | ')}. Pasa --carpeta=.`);
        const ext = (path.extname(sugerido) || '.pdf').toLowerCase();
        const semana = args['semana'] ? Number(args['semana']) : null;
        const nombre = ex.nombreArchivo({ soc, banco: args['banco'], periodo: p, semana, ext });
        const res = ex.archivar({ origen: local, dirBanco: cb.dir, periodo: p, nombre });
        fs.rmSync(local, { force: true });   // copia propia en la bandeja
        return out({ ok: true, archivo: sugerido, estado: res.estado, destino: res.destino });
      }

      case 'archivar': {
        // Archiva un archivo ya bajado a la bandeja (despues de revisar su periodo).
        const ex = require('../lib/extractos');
        const local = args['archivo'];
        if (!local || !fs.existsSync(local)) throw new Error('Falta --archivo (ruta en la bandeja).');
        const p = ex.periodo(args['periodo']);
        if (!p || !args['banco'] || !args['cuit']) throw new Error('Hacen falta --banco, --cuit y --periodo=AAAAMM.');
        const soc = ex.sociedad({ cuit: args['cuit'] });
        const cb = 'unidad' in args
          ? ex.carpetaBanco(soc, args['banco'], args['carpeta'] ? { elegida: args['carpeta'] } : {})
          : ex.carpetaSalida(ex.salidaBase(args['salida']), soc, args['banco']);
        if (cb.candidatas) throw new Error(`Hay varias carpetas para ${args['banco']}: ${cb.candidatas.join(' | ')}. Pasa --carpeta=.`);
        const ext = (path.extname(local.split('__').pop()) || '.pdf').toLowerCase();
        const nombre = ex.nombreArchivo({ soc, banco: args['banco'], periodo: p, semana: args['semana'] ? Number(args['semana']) : null, ext });
        const res = ex.archivar({ origen: local, dirBanco: cb.dir, periodo: p, nombre });
        fs.rmSync(local, { force: true });
        return out({ ok: true, estado: res.estado, destino: res.destino });
      }

      case 'listo': {
        // clic/descargar dejan el recuadro sin eventos para no tapar el banco: se restaura.
        await page.evaluate(() => { const b = document.getElementById('bva-barra'); if (b) b.style.pointerEvents = 'auto'; }).catch(() => {});
        const btn = page.locator('#bva-barra button');
        if (!(await btn.isVisible().catch(() => false))) throw new Error('No hay recuadro Estudio BVA esperando respuesta.');
        await btn.click();
        return out({ ok: true });
      }

      case 'responder': {
        await page.evaluate(() => { const b = document.getElementById('bva-barra'); if (b) b.style.pointerEvents = 'auto'; }).catch(() => {});
        const inp = page.locator('#bva-barra input');
        if (!(await inp.isVisible().catch(() => false))) throw new Error('El recuadro no esta pidiendo un valor.');
        await inp.fill(String(args['valor'] || ''));
        await page.locator('#bva-barra button').click();
        return out({ ok: true });
      }

      default:
        throw new Error(`Comando desconocido "${cmd || ''}". Ver la ayuda al principio de scripts/control.js.`);
    }
  } catch (e) {
    out({ error: e.message.split('\n')[0], ...('debug' in args ? { detalle: e.message.split('\n').slice(1, 14) } : {}) });
  }
})();
