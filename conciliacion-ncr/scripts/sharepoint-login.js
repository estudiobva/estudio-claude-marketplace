#!/usr/bin/env node
/**
 * sharepoint-login.js — Ingreso al SharePoint del estudio (EstudioContable), sin tipear nada.
 *
 * Es el equivalente de login.js (ARCA) para Microsoft 365: lee el usuario y la clave de
 * Claves_Organismos.xlsx (fila "Sharepoint - Microsoft"), entra con Playwright y el
 * navegador OCULTO, y guarda la sesion en ~/.bva-playwright/sharepoint-state.json para
 * no volver a pedir la clave en las corridas siguientes.
 *
 *   node scripts/sharepoint-login.js
 *   node scripts/sharepoint-login.js --carpeta="NC por NRF/2026"
 *   node scripts/sharepoint-login.js --forzar        (ignora la sesion guardada)
 *   node scripts/sharepoint-login.js --olvidar       (borra la sesion guardada y sale)
 *
 * Argumentos:
 *   --sitio      URL del sitio (default: SHAREPOINT_SITIO del .env).
 *   --carpeta    Ademas de entrar, lista las subcarpetas de esa ruta de "Documentos" (solo lectura).
 *   --forzar     No usa la sesion guardada: entra con usuario y clave.
 *   --olvidar    Borra la sesion guardada.
 *   --ver        Navegador visible. Solo si lo pedis: por defecto corre oculto.
 *
 * Salida: JSON por stdout; el progreso va por stderr. NUNCA imprime la clave.
 * Solo lectura: no sube, no mueve ni borra nada en el SharePoint.
 */

const { parseArgs, fail } = require('../lib/args');
const sp = require('../lib/sharepoint-login');

const args = parseArgs();
if ('ver' in args && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';

const log = (m) => process.stderr.write(`[sharepoint] ${m}\n`);
const SITIO = (args['sitio'] || sp.sitioDefault()).replace(/\/+$/, '');

(async () => {
  if ('olvidar' in args) {
    sp.borrarEstado();
    console.log(JSON.stringify({ ok: true, sesionGuardada: 'borrada' }));
    return;
  }

  let sesion;
  try {
    log('entrando al SharePoint…');
    sesion = await sp.abrirSharepoint({ sitio: SITIO, forzar: 'forzar' in args, log });
  } catch (e) {
    if (e.code === 'credenciales_invalidas') {
      fail(e.message, { code: e.code, que_hacer: 'Verifica la fila "Sharepoint - Microsoft" de Claves_Organismos.xlsx. No reintentar de mas.' });
    }
    if (e.code === 'segundo_factor') {
      fail(e.message, { code: e.code, que_hacer: 'La cuenta tiene verificacion en dos pasos. Hay que excluirla (o usar una cuenta de servicio sin MFA) o entrar una vez a mano.' });
    }
    if (e.code === 'captcha') {
      fail(e.message, { code: e.code, que_hacer: 'Estos scripts no resuelven captchas.' });
    }
    fail(e.message, { code: e.code || 'error' });
  }

  const { browser, page, entrada } = sesion;
  try {
    const web = await page.request.get(`${SITIO}/_api/web?$select=Title,Url`, {
      headers: { Accept: 'application/json;odata=nometadata' },
    }).then((r) => r.json());
    log(`adentro: ${web.Title}`);

    const salida = { ok: true, sitio: web.Title, url: web.Url, entrada, sesionGuardada: sp.ESTADO_FILE };

    if (args['carpeta']) {
      const ruta = `${new URL(SITIO).pathname}/Shared Documents/${args['carpeta']}`.replace(/'/g, "''");
      const r = await page.request.get(
        `${SITIO}/_api/web/GetFolderByServerRelativePath(decodedurl='${ruta}')/Folders?$select=Name`,
        { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!r.ok()) fail(`No pude listar "${args['carpeta']}": SharePoint ${r.status()}.`, { code: 'carpeta' });
      salida.carpeta = args['carpeta'];
      salida.subcarpetas = ((await r.json()).value || []).map((x) => x.Name);
    }
    console.log(JSON.stringify(salida, null, 2));
  } finally {
    await browser.close().catch(() => {});
  }
})();
