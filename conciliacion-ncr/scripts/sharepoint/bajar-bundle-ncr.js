#!/usr/bin/env node
/**
 * bajar-bundle-ncr.js — arma el paquete SP-NCR-AAAAMM.bundle del SharePoint por
 * Playwright con el navegador OCULTO, sin pasar por Chrome ni por Claude in Chrome.
 *
 * Reemplaza a bundle-ncr.js (que hay que pegar en la pestaña del SharePoint): hace lo
 * mismo, solo lectura, y deja el paquete en el MISMO formato, asi que descargar-ncr.js y
 * ncr-lote.js lo abren igual. Entra con lib/sharepoint-login.js (usuario y clave de
 * Claves_Organismos.xlsx, o la sesion guardada).
 *
 *   node scripts/sharepoint/bajar-bundle-ncr.js --periodo=08/2026
 *   node scripts/sharepoint/bajar-bundle-ncr.js --periodo=202608 --salida="/ruta/SP-NCR-202608.bundle"
 *
 * Argumentos:
 *   --periodo   MM/AAAA o AAAAMM.
 *   --salida    Ruta del paquete (default: ~/Downloads/SP-NCR-AAAAMM.bundle, pisa el anterior).
 *   --forzar    No usa la sesion guardada del SharePoint.
 *   --ver       Navegador visible. Solo si lo pedis.
 *
 * Baja SUIZO/ y MONROE/ de "NC por NRF/<AAAA>/<MES>/". Formato del paquete: 4 bytes
 * (largo del indice, big endian) + indice JSON + los archivos concatenados.
 * Salida: JSON por stdout; progreso por stderr. Solo lectura (GET).
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');

const { parseArgs, fail } = require('../../lib/args');
const sp = require('../../lib/sharepoint-login');

const args = parseArgs();
if ('ver' in args && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';
const log = (m) => process.stderr.write(`[bajar-bundle-ncr] ${m}\n`);

const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
               'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
const norm = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
const q = (p) => p.replace(/'/g, "''");

function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)))          return s;
  return null;
}

(async () => {
  const periodo = normalizarPeriodo(args['periodo']);
  if (!periodo || Number(periodo.slice(4)) < 1 || Number(periodo.slice(4)) > 12) fail('Falta --periodo (MM/AAAA o AAAAMM).');
  const salida = path.resolve(args['salida'] || path.join(os.homedir(), 'Downloads', `SP-NCR-${periodo}.bundle`));

  let sesion;
  try {
    sesion = await sp.abrirSharepoint({ forzar: 'forzar' in args, log });
  } catch (e) {
    fail(e.message, { code: e.code || 'error' });
  }
  const { browser, page } = sesion;

  try {
    const base = sp.sitioDefault();
    const sitio = new URL(base).pathname;                       // /sites/<sitio>
    const api = async (ruta) => {
      const r = await page.request.get(`${base}/_api/web/${ruta}`, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!r.ok()) throw new Error(`SharePoint ${r.status()} en ${ruta}`);
      return (await r.json()).value || [];
    };
    const carpetas = (p) => api(`GetFolderByServerRelativePath(decodedurl='${q(p)}')/Folders?$select=Name`);
    const archivos = (p) => api(`GetFolderByServerRelativePath(decodedurl='${q(p)}')/Files?$select=Name,Length,TimeLastModified`);

    const anio = periodo.slice(0, 4), mes = MESES[Number(periodo.slice(4)) - 1];
    const baseAnio = `${sitio}/Shared Documents/NC por NRF/${anio}`;
    const meses = (await carpetas(baseAnio)).map((x) => x.Name);
    const carpetaMes = meses.find((m) => norm(m) === mes || (mes === 'SEPTIEMBRE' && norm(m) === 'SETIEMBRE'));
    if (!carpetaMes) fail(`No hay carpeta ${mes} en NC por NRF/${anio}`, { code: 'sin_carpeta', hay: meses });
    const baseMes = `${baseAnio}/${carpetaMes}`;
    log(`bajando ${baseMes.replace(`${sitio}/Shared Documents/`, '')}…`);

    const indice = { periodo, origen: baseMes, generado: new Date().toISOString(), archivos: [] };
    const partes = [];
    for (const sub of (await carpetas(baseMes)).map((x) => x.Name)) {
      const drog = norm(sub);
      if (drog !== 'SUIZO' && drog !== 'MONROE') continue;
      for (const a of await archivos(`${baseMes}/${sub}`)) {
        const r = await page.request.get(
          `${base}/_api/web/GetFileByServerRelativePath(decodedurl='${q(`${baseMes}/${sub}/${a.Name}`)}')/$value`);
        if (!r.ok()) fail(`No pude bajar ${sub}/${a.Name}: ${r.status()}`, { code: 'descarga' });
        const buf = await r.body();
        indice.archivos.push({ drogueria: drog, nombre: a.Name, bytes: buf.length, modificado: a.TimeLastModified });
        partes.push(buf);
      }
    }
    if (!indice.archivos.length) fail(`No encontre archivos de SUIZO ni MONROE en ${carpetaMes}.`, { code: 'vacio' });

    const cab = Buffer.from(JSON.stringify(indice), 'utf-8');
    const largo = Buffer.alloc(4); largo.writeUInt32BE(cab.length);
    fs.mkdirSync(path.dirname(salida), { recursive: true });
    fs.writeFileSync(salida, Buffer.concat([largo, cab, ...partes]));

    const cuenta = (d) => indice.archivos.filter((x) => x.drogueria === d).length;
    console.log(JSON.stringify({
      ok: true, periodo, archivo: salida, bytes: fs.statSync(salida).size,
      suizo: cuenta('SUIZO'), monroe: cuenta('MONROE'), entrada: sesion.entrada,
    }, null, 2));
  } catch (e) {
    fail(e.message, { code: 'error' });
  } finally {
    await browser.close().catch(() => {});
  }
})();
