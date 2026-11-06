#!/usr/bin/env node
/**
 * bajar-ventas-zetti.js — paso 1 del proceso TSH / IIBB.
 *
 * Baja de Zetti (T&S Web) el Subdiario de IVA Ventas (reporte 5.6.5, CSV) de una
 * sociedad para un mes, lo pasa a Excel con formato de papel de trabajo (con la
 * columna de jurisdiccion y totales de I a M) y lo guarda en la unidad compartida:
 *
 *   [Sociedad]/01-Impuestos Mensuales/AAAAMM/AAAAMM- papeles de trabajo- <Farmacia>.xlsx
 *
 *   node scripts/bajar-ventas-zetti.js --cuit=<CUIT> --periodo=202609
 *   node scripts/bajar-ventas-zetti.js --cuit=<CUIT> --periodo=09/2026 --si-existe=reemplazar
 *   node scripts/bajar-ventas-zetti.js --cuit=<CUIT> --periodo=202510 --subcarpeta=controles
 *     (excepcion: guarda en 01-Impuestos Mensuales/controles/AAAAMM, que tiene que existir)
 *
 * Login y descarga con lib/ (copia de zetti-subdiario-iva-ventas). Solo lee Zetti.
 * Normalmente lo llama papel-ventas.js, que arma el papel completo.
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
const log = (m) => process.stderr.write(`[tsh-iibb] ${m}\n`);

const s = String(args.periodo || '').trim();
let m, PERIODO;
if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) PERIODO = m[2] + m[1].padStart(2, '0');
else if (/^\d{6}$/.test(s)) PERIODO = s;
else fail('--periodo tiene que ser MM/AAAA o AAAAMM');
const CUIT = String(args.cuit || '').replace(/\D/g, '');
if (CUIT.length !== 11) fail('Falta --cuit (11 digitos)');
const SI_EXISTE = String(args['si-existe'] || 'abortar').toLowerCase();

const ANIO = Number(PERIODO.slice(0, 4)), MES = PERIODO.slice(4);
const DESDE = `01/${MES}/${ANIO}`;
const HASTA = `${String(new Date(ANIO, Number(MES), 0).getDate()).padStart(2, '0')}/${MES}/${ANIO}`;

// Carpeta de la sociedad: "<N> - <Farmacia> - <CUIT>" -> nombre "<Farmacia>".
const socDir = rutas.carpetaSociedad(CUIT);
const FARMACIA = path.basename(socDir).replace(/^\s*\d+\s*-\s*/, '').replace(new RegExp(`\\s*-\\s*${CUIT}\\s*$`), '').trim();
// --dir=<carpeta>: destino explicito (lo usa papel-ventas.js).
// --subcarpeta=controles: guarda en 01-Impuestos Mensuales/controles/AAAAMM, que tiene que existir.
const SUB = String(args.subcarpeta || '').trim();
const mensuales = path.join(socDir, '01-Impuestos Mensuales');
const mesDir = args.dir ? String(args.dir) : SUB ? path.join(mensuales, SUB, PERIODO) : rutas.carpetaMes(mensuales, PERIODO);
if (!fs.existsSync(mesDir)) fail(`No existe la carpeta ${mesDir}`);
const destino = path.join(mesDir, `${PERIODO}- papeles de trabajo- ${FARMACIA}.xlsx`);
if (fs.existsSync(destino) && SI_EXISTE !== 'reemplazar') fail(`Ya existe ${destino}. Usar --si-existe=reemplazar.`, { code: 'ya_existe' });

(async () => {
  const browser = await chromium.launch(launchOptions());
  const ctx = await browser.newContext({ acceptDownloads: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'bva-tsh-'));
  try {
    const { page, auth } = await zetti.login(ctx, credencialesZetti(), log);
    log('leyendo entidades y CUITs');
    const ents = await zetti.entidades(page);
    const { entidad, locales } = zetti.entidadDeCuit(ents, CUIT);
    log(`${entidad.nombre} (${locales.length} locales): reporte 5.6.5 del ${DESDE} al ${HASTA}`);
    const csv = await zetti.subdiarioIvaVentas(ctx, auth, entidad, { desde: DESDE, hasta: HASTA }, path.join(tmp, 'reporte.csv'));
    const xlsx = path.join(tmp, 'salida.xlsx');
    const py = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
    let res;
    try {
      res = JSON.parse(execFileSync(py, [path.join(__dirname, 'zetti_papel_trabajo.py'), csv, xlsx, `--periodo=${PERIODO}`], { encoding: 'utf-8' }).trim().split(/\r?\n/).pop());
    } catch (e) {
      fail(String(e.stdout || e.message).trim(), { code: 'control' });
    }
    if (res.cuit !== CUIT) fail(`El reporte salio con el CUIT ${res.cuit}, se pidio ${CUIT}`, { code: 'control' });
    if (res.desde !== DESDE || res.hasta !== HASTA) fail(`El reporte salio del ${res.desde} al ${res.hasta}`, { code: 'control' });
    fs.copyFileSync(xlsx, destino);
    console.log(JSON.stringify({ ok: true, farmacia: FARMACIA, entidad: entidad.nombre, locales: locales.map((l) => l.nombre), ...res, archivo: destino }, null, 2));
  } finally {
    await browser.close().catch(() => {});
    fs.rmSync(tmp, { recursive: true, force: true });
  }
})().catch((e) => fail(e.message, { code: e.code || 'error' }));
