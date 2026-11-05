#!/usr/bin/env node
/**
 * bajar-compras-zetti.js — Subdiario de IVA Compras de Zetti (reporte 5.6.4, CSV) a Excel.
 *
 * Mismo procedimiento que bajar-ventas-zetti.js: sociedad completa (todos los locales),
 * emision del 1 al ultimo dia del mes, formato CSV, pasado a Excel y guardado en:
 *
 *   [Sociedad]/01-Impuestos Mensuales/AAAAMM/AAAAMM - ZETTI - COMPRAS.xlsx
 *
 *   node scripts/bajar-compras-zetti.js --cuits="30699346019;30595259548" --periodo=202609
 *   node scripts/bajar-compras-zetti.js --cuit=30699346019 --periodo=09/2026 --si-existe=reemplazar
 *
 * Un solo login para todas. Login automatico con el usuario de Zetti de la planilla de claves. Solo lee Zetti.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

// lib/ vendorizada de zetti-subdiario-iva-ventas (login de Zetti, planilla de claves, rutas de la unidad).
const { chromium } = require('playwright');
const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const { credencialesZetti } = require('../lib/zetti-claves');
const zetti = require('../lib/zetti');
const rutas = require('../lib/rutas-bva');

const args = parseArgs();
loadEnv();
const log = (m) => process.stderr.write(`[compras] ${m}\n`);

const s = String(args.periodo || '').trim();
let m, PERIODO;
if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) PERIODO = m[2] + m[1].padStart(2, '0');
else if (/^\d{6}$/.test(s)) PERIODO = s;
else fail('--periodo tiene que ser MM/AAAA o AAAAMM');
const CUITS = String(args.cuits || args.cuit || '').split(/[;,\s]+/).map((c) => c.replace(/\D/g, '')).filter(Boolean);
if (!CUITS.length || CUITS.some((c) => c.length !== 11)) fail('Falta --cuit / --cuits (11 digitos cada uno)');
const SI_EXISTE = String(args['si-existe'] || 'abortar').toLowerCase();

const ANIO = Number(PERIODO.slice(0, 4)), MES = PERIODO.slice(4);
const DESDE = `01/${MES}/${ANIO}`;
const HASTA = `${String(new Date(ANIO, Number(MES), 0).getDate()).padStart(2, '0')}/${MES}/${ANIO}`;
const PY = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');

// 5.6.4 Reporte IVA Compra en CSV, de desde a hasta (DD/MM/AAAA).
async function subdiarioIvaCompras(ctx, auth, entidad, destino) {
  const f = await zetti.abrirReporte(ctx, auth, 'rpt_subdiario_iva_compra', entidad);
  try {
    await f.fill('input[name=fecEDesde]', DESDE);
    await f.fill('input[name=fecEHasta]', HASTA);
    await f.fill('input[name=periodo]', '');
    await f.selectOption('select[name=tipo_codif]', { label: 'Mostrar todos' });
    for (const n of ['no_importados', 'resumen', 'hideParamUser', 'hideHeader']) {
      await f.locator(`input[name=${n}]`).setChecked(false);
    }
    await f.selectOption('select[name=tipo]', 'csv');
    const descarga = ctx.waitForEvent('download', { timeout: 300000 }).catch(() => null);
    await f.getByText('Aceptar', { exact: true }).click();
    const dl = await descarga;
    if (!dl) {
      const texto = (await f.locator('body').innerText().catch(() => '')).replace(/\s+/g, ' ').slice(0, 300);
      throw Object.assign(new Error(`El reporte no bajo ningun archivo en 5 minutos. Pantalla: ${texto}`), { code: 'sin_descarga' });
    }
    await dl.saveAs(destino);
    return destino;
  } finally {
    await f.close().catch(() => {});
  }
}

function destinoDe(cuit) {
  const socDir = rutas.carpetaSociedad(cuit);
  const mensuales = path.join(socDir, '01-Impuestos Mensuales');
  const mesDir = rutas.carpetaMes(mensuales, PERIODO);
  if (!fs.existsSync(mesDir)) throw Object.assign(new Error(`No existe la carpeta ${mesDir}`), { code: 'sin_carpeta' });
  return { socDir, archivo: path.join(mesDir, `${PERIODO} - ZETTI - COMPRAS.xlsx`) };
}

(async () => {
  const browser = await chromium.launch(launchOptions());
  const ctx = await browser.newContext({ acceptDownloads: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'bva-compras-'));
  const salida = [];
  try {
    const { page, auth } = await zetti.login(ctx, credencialesZetti(), log);
    log('leyendo entidades y CUITs');
    const ents = await zetti.entidades(page);
    for (const cuit of CUITS) {
      const r = { cuit };
      try {
        const { socDir, archivo } = destinoDe(cuit);
        r.carpeta = path.basename(socDir);
        if (fs.existsSync(archivo) && SI_EXISTE !== 'reemplazar') {
          Object.assign(r, { ok: false, code: 'ya_existe', archivo });
          salida.push(r);
          continue;
        }
        const { entidad, locales } = zetti.entidadDeCuit(ents, cuit);
        r.entidad = entidad.nombre;
        r.locales = locales.length;
        log(`${entidad.nombre} (${locales.length} locales): 5.6.4 del ${DESDE} al ${HASTA}`);
        const csv = await subdiarioIvaCompras(ctx, auth, entidad, path.join(tmp, `${cuit}.csv`));
        const xlsx = path.join(tmp, `${cuit}.xlsx`);
        let res;
        try {
          res = JSON.parse(execFileSync(PY, [path.join(__dirname, 'zetti_compras_excel.py'), csv, xlsx, `--periodo=${PERIODO}`], { encoding: 'utf-8' }).trim().split(/\r?\n/).pop());
        } catch (e) {
          const out = String(e.stdout || '').trim().split(/\r?\n/).pop();
          let j; try { j = JSON.parse(out); } catch { j = { error: out || e.message }; }
          throw Object.assign(new Error(j.error), { code: j.sin_datos ? 'sin_datos' : 'control', detalle: j });
        }
        if (res.cuit !== cuit) throw Object.assign(new Error(`El reporte salio con el CUIT ${res.cuit}`), { code: 'control' });
        if (res.desde !== DESDE || res.hasta !== HASTA) throw Object.assign(new Error(`El reporte salio del ${res.desde} al ${res.hasta}`), { code: 'control' });
        fs.copyFileSync(xlsx, archivo);
        delete res.xlsx;
        Object.assign(r, { ok: true, ...res, archivo });
      } catch (e) {
        Object.assign(r, { ok: false, code: e.code || 'error', error: e.message, ...(e.detalle ? { detalle: e.detalle } : {}) });
        log(`${cuit}: ${e.message}`);
      }
      salida.push(r);
    }
    console.log(JSON.stringify(salida, null, 1));
  } finally {
    await browser.close().catch(() => {});
    fs.rmSync(tmp, { recursive: true, force: true });
  }
})().catch((e) => fail(e.message, { code: e.code || 'error' }));
