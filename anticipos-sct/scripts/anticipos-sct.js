#!/usr/bin/env node
/**
 * anticipos-sct.js — releva los anticipos de Ganancias y Bienes Personales
 * de una sociedad en el Sistema de Cuentas Tributarias, y los vuelca a Excel.
 *
 *   node scripts/anticipos-sct.js --nombre="<sociedad>"
 *   node scripts/anticipos-sct.js --cuit=<cuit> --ver
 *   node scripts/anticipos-sct.js --nombre="<sociedad>" --crudo
 *
 * Argumentos:
 *   --nombre / --cuit  La sociedad.
 *   --salida           Carpeta destino (default: ~/Documents/BVA-salidas/anticipos,
 *                      o $BVA_SALIDAS_PATH/anticipos).
 *   --crudo            Ademas del Excel, guarda el JSON con TODAS las filas de
 *                      la solapa Deudas (util para entender un caso raro).
 *   --ver              Navegador visible.
 *
 * Se trabaja de a una sociedad. No recorre el portfolio entero de una.
 *
 * Esto solo LEE el SCT. Nunca modifica, elimina, presenta ni paga nada.
 */

const fs   = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { resolveCredentials, envFile } = require('../lib/env');
const { login } = require('../lib/arca-login');
const sctLib = require('../lib/sct');

const args = parseArgs();
const VER   = 'ver' in args;
const CRUDO = 'crudo' in args;
if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';

const log = (m) => process.stderr.write(`[anticipos] ${m}\n`);
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');

let CUIT, CUIT_LOGIN, PASSWORD, TITULAR;
try {
  ({ cuit: CUIT, cuitLogin: CUIT_LOGIN, password: PASSWORD, titular: TITULAR } =
    resolveCredentials({
      cuit: args['cuit'], nombre: args['nombre'],
      cuitLogin: args['cuit-login'], password: args['password'],
    }));
} catch (e) {
  fail(e.message, { sugerencia: `Revisa el .env (${envFile() || 'no encontrado'}) o el Excel de claves.` });
}

const SALIDA = args['salida'] || path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'anticipos');
const HOY = new Date();
// Fecha LOCAL: toISOString() es UTC y despues de las 21 h da el dia siguiente.
const iso = [HOY.getFullYear(), String(HOY.getMonth() + 1).padStart(2, '0'),
             String(HOY.getDate()).padStart(2, '0')].join('-');
const slug = (TITULAR || CUIT).toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');

(async () => {
  // launch.js lee PLAYWRIGHT_HEADLESS al cargarse, antes de que --ver lo toque.
  const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 250 } : {}));
  const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1500, height: 950 } });
  const page = await context.newPage();

  try {
    log(`${TITULAR || CUIT} · CUIT ${CUIT}`);
    log('entrando a ARCA…');
    await login(page, { cuitLogin: CUIT_LOGIN, password: PASSWORD });

    const { sct, marco } = await sctLib.abrirSct(context, page, CUIT);
    const { encabezados, filas, resumen } = await sctLib.leerDeudas(sct, marco);
    const { anticipos, inesperados, otrasDeudas } = sctLib.clasificar(filas, HOY);

    log(`anticipos de Ganancias/Bienes Personales: ${anticipos.length}`);
    if (inesperados.length) {
      log(`NOVEDAD: ${inesperados.length} anticipo(s) de otro impuesto — se reportan aparte:`);
      for (const a of inesperados) log(`   ${a.impuesto} · periodo ${a.periodo} · ${a.importe}`);
    }
    log(`otras deudas que no son anticipos: ${otrasDeudas}`);

    const vencidos = anticipos.filter(a => a.estado === 'Vencido').length;

    fs.mkdirSync(SALIDA, { recursive: true });
    const destino = path.join(SALIDA, `Anticipos_SCT_${slug}_${iso}.xlsx`);

    // Si la sociedad no tiene anticipos, igual se genera el Excel con una fila
    // "Sin anticipos": saltearla en silencio hace que despues nadie sepa si se
    // reviso o no.
    const payload = {
      sociedad: TITULAR || CUIT,
      cuit: CUIT,
      fecha: iso,
      anticipos,
      inesperados,
      resumen,
    };
    const jsonTmp = path.join(SALIDA, `.anticipos-${CUIT}.json`);
    fs.writeFileSync(jsonTmp, JSON.stringify(payload), 'utf-8');

    // Si falla el Excel, el JSON queda: lo leido del SCT no se pierde y se
    // puede regenerar con anticipos_a_excel.py sin volver a entrar a ARCA.
    const { execFileSync } = require('child_process');
    let res;
    try {
      const salida = execFileSync(PYTHON,
        [path.join(__dirname, 'anticipos_a_excel.py'), jsonTmp, destino],
        { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
      res = JSON.parse(salida.trim().split(/\r?\n/).pop());
    } catch (err) {
      const e = new Error(`No pude armar el Excel (${PYTHON}): ${(err.stderr || err.message || '').toString().trim().split(/\r?\n/).pop()}. ` +
        `Los datos leidos quedaron en ${jsonTmp}; regeneralo con: ${PYTHON} scripts/anticipos_a_excel.py "${jsonTmp}" "${destino}"`);
      e.code = 'excel_fallo';
      throw e;
    }
    fs.unlinkSync(jsonTmp);
    log(`Excel: ${res.filas} fila(s) -> ${destino}`);

    if (CRUDO) {
      const crudo = path.join(SALIDA, `deudas-crudo-${CUIT}-${iso}.json`);
      fs.writeFileSync(crudo, JSON.stringify({ encabezados, filas }, null, 1), 'utf-8');
      log(`filas crudas de la solapa Deudas: ${crudo}`);
    }

    console.log(JSON.stringify({
      ok: true,
      sociedad: TITULAR || null, cuit: CUIT, fecha: iso,
      anticipos: anticipos.length,
      vencidos, noVencidos: anticipos.length - vencidos,
      anticiposInesperados: inesperados,
      otrasDeudasNoAnticipo: otrasDeudas,
      resumenSct: resumen,
      archivo: destino,
      detalle: anticipos,
    }, null, 2));
  } catch (e) {
    if (e.code === 'captcha') {
      fail('ARCA pidio un captcha.', { code: 'captcha',
        que_hacer: `Corre 'node scripts/login.js --cuit=${CUIT} --quedate', resolvelo a mano, y reintenta.` });
    }
    if (e.code === 'credenciales_invalidas') {
      fail(`ARCA rechazo las credenciales del CUIT ${CUIT}.`, { code: 'credenciales_invalidas',
        que_hacer: 'Verifica la fila en Claves_Organismos.xlsx. Ojo con reintentar: ARCA bloquea la clave.' });
    }
    fail(e.message, { code: e.code || 'error' });
  } finally {
    await browser.close().catch(() => {});
  }
})();
