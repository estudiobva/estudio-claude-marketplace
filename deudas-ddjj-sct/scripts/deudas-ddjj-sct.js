#!/usr/bin/env node
/**
 * deudas-ddjj-sct.js — releva en el Sistema de Cuentas Tributarias de una
 * sociedad las deudas que NO son anticipos (solapa Deudas) y las DDJJ
 * pendientes de presentacion (solapa "DDJJ pendientes de presentación"), y
 * las vuelca a Excel.
 *
 *   node scripts/deudas-ddjj-sct.js --nombre="<sociedad>"
 *   node scripts/deudas-ddjj-sct.js --cuit=<cuit> --ver
 *   node scripts/deudas-ddjj-sct.js --nombre="<sociedad>" --crudo
 *
 * Argumentos:
 *   --nombre / --cuit  La sociedad.
 *   --salida           Carpeta destino (default: ~/Documents/BVA-salidas/deudas-ddjj,
 *                      o $BVA_SALIDAS_PATH/deudas-ddjj).
 *   --crudo            Ademas del Excel, guarda el JSON con TODAS las filas de
 *                      las dos solapas (util para entender un caso raro).
 *   --ver              Navegador visible.
 *
 * Los anticipos se excluyen a proposito: los releva anticipos-sct.js. Los dos
 * usan el mismo criterio de "anticipo" (lib/sct.js), asi que entre ambos se
 * cubre toda la solapa Deudas sin duplicar nada.
 *
 * Esto solo LEE el SCT. Nunca modifica, elimina, presenta ni paga nada.
 */

const fs   = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
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

const log = (m) => process.stderr.write(`[deudas-ddjj] ${m}\n`);
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

const SALIDA = args['salida'] || path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'deudas-ddjj');
const HOY = new Date();
// Fecha LOCAL: toISOString() es UTC y despues de las 21 h da el dia siguiente.
const iso = [HOY.getFullYear(), String(HOY.getMonth() + 1).padStart(2, '0'),
             String(HOY.getDate()).padStart(2, '0')].join('-');
const slug = (TITULAR || CUIT).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');

(async () => {
  const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 250 } : {}));
  const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1500, height: 950 } });
  const page = await context.newPage();

  try {
    log(`${TITULAR || CUIT} · CUIT ${CUIT}`);
    log('entrando a ARCA…');
    await login(page, { cuitLogin: CUIT_LOGIN, password: PASSWORD });

    const { sct, marco } = await sctLib.abrirSct(context, page, CUIT);
    const d = await sctLib.leerDeudas(sct, marco);
    const p = await sctLib.leerPendientes(sct, marco);

    const deudas = sctLib.otrasDeudas(d.filas, HOY);
    const pendientes = sctLib.pendientes(p.filas, HOY);
    const anticiposExcluidos = d.filas.length - deudas.length;
    const vencidas = deudas.filter(x => x.estado === 'Vencido').length;

    log(`deudas que no son anticipos: ${deudas.length} (${vencidas} vencidas) · ` +
        `anticipos excluidos: ${anticiposExcluidos}`);
    log(`DDJJ pendientes de presentacion: ${pendientes.length}`);

    fs.mkdirSync(SALIDA, { recursive: true });
    const destino = path.join(SALIDA, `Deudas_DDJJ_SCT_${slug}_${iso}.xlsx`);
    const payload = {
      sociedad: TITULAR || CUIT, cuit: CUIT, fecha: iso,
      deudas, pendientes, anticiposExcluidos, resumen: d.resumen,
    };
    const jsonTmp = path.join(SALIDA, `.deudas-ddjj-${CUIT}.json`);
    fs.writeFileSync(jsonTmp, JSON.stringify(payload), 'utf-8');

    // Si falla el Excel, el JSON queda: lo leido del SCT no se pierde y se
    // puede regenerar con deudas_ddjj_a_excel.py sin volver a entrar a ARCA.
    let res;
    try {
      const salida = execFileSync(PYTHON,
        [path.join(__dirname, 'deudas_ddjj_a_excel.py'), jsonTmp, destino],
        { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
      res = JSON.parse(salida.trim().split(/\r?\n/).pop());
    } catch (err) {
      const e = new Error(`No pude armar el Excel (${PYTHON}): ${(err.stderr || err.message || '').toString().trim().split(/\r?\n/).pop()}. ` +
        `Los datos leidos quedaron en ${jsonTmp}; regeneralo con: ${PYTHON} scripts/deudas_ddjj_a_excel.py "${jsonTmp}" "${destino}"`);
      e.code = 'excel_fallo';
      throw e;
    }
    fs.unlinkSync(jsonTmp);
    log(`Excel: ${res.deudas} deuda(s), ${res.pendientes} DDJJ pendiente(s) -> ${destino}`);

    if (CRUDO) {
      const crudo = path.join(SALIDA, `sct-crudo-${CUIT}-${iso}.json`);
      fs.writeFileSync(crudo, JSON.stringify({ deudas: d, pendientes: p }, null, 1), 'utf-8');
      log(`filas crudas de las dos solapas: ${crudo}`);
    }

    console.log(JSON.stringify({
      ok: true,
      sociedad: TITULAR || null, cuit: CUIT, fecha: iso,
      deudas: deudas.length, deudasVencidas: vencidas,
      ddjjPendientes: pendientes.length,
      anticiposExcluidos,
      archivo: destino,
      detalleDeudas: deudas,
      detallePendientes: pendientes,
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
