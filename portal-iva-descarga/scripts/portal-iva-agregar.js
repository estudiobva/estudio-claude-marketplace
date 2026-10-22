#!/usr/bin/env node
/**
 * portal-iva-agregar.js — carga en el borrador del Libro Compras (Portal IVA)
 * las facturas de servicios (Edesur, Metrogas, AySA, ...) que estan en las
 * compras de Zetti y no vinieron con la importacion desde ARCA. Las carga una
 * por una con el boton "AGREGAR COMPROBANTE..." del libro, como a mano.
 *
 *   node scripts/portal-iva-agregar.js --cuit=<CUIT> --periodo=09/2026              (simulacion)
 *   node scripts/portal-iva-agregar.js --cuit=<CUIT> --periodo=09/2026 --ejecutar --limite=1
 *   node scripts/portal-iva-agregar.js --nombre="<sociedad>" --periodo=09/2026 --archivo=servicios.xlsx --ejecutar
 *
 * Argumentos:
 *   --nombre / --cuit     Sociedad (titular de Claves_Organismos.xlsx, o CUIT).
 *   --periodo             MM/AAAA o AAAAMM. Default: el mes anterior.
 *   --proveedores         Nombres a buscar en la columna Proveedor, separados por ";".
 *                         Default "EDESUR;METROGAS;AYSA;AGUA Y SANEAMIENTOS". "*" = todos.
 *   --archivo             Excel con las columnas de la solapa "Compras Zetti" (Fecha, TC,
 *                         M, Nro. Comp., Proveedor, CUIT, Exen, Grav, IVA, P/R IIBB,
 *                         P/R IVA, Total). Default: [IVA]/AAAAMM - ZETTI - COMPRAS.xlsx.
 *                         Sirve para cargar facturas que no estan en Zetti (Metrogas,
 *                         AySA): se arma un Excel con esas columnas desde el PDF.
 *   --ejecutar            Guarda de verdad. SIN este flag solo simula: abre el formulario,
 *                         lo completa, saca una captura y lo cancela.
 *   --limite=N            Carga solo los primeros N (para pilotos).
 *   --cred-fiscal=arca    Credito Fiscal Computable: "arca" (default, lo deja como lo
 *                         calcula ARCA, igual que en la importacion), "iva" o "0".
 *   --aunque-presentada   Sigue aunque haya acuse de la DDJJ IVA en la carpeta. ARCA
 *                         igual no deja abrir un periodo presentado.
 *   --ver                 Navegador visible.
 *
 * Como carga: AGREGAR COMPROBANTE... abre #modalAltaComprobante; se completan los
 * campos cm_* (al cargar el neto gravado ARCA calcula el IVA; despues se pisa con
 * el IVA de la factura) y se clickea "Agregar" (#btnAgregarComprobanteManual), que
 * manda guardarComprobanteManual.do. La respuesta de ARCA (ok, errores u
 * observaciones) queda en el log.
 *
 * No carga un comprobante si:
 *   - ya esta en el libro (mismo CUIT, tipo, punto de venta y numero);
 *   - los importes no suman el total, o el IVA no corresponde a una sola alicuota;
 *   - el TC/letra no es FC/ND/NC A, B o C.
 * No hace nada si el periodo ya esta presentado, si el borrador no esta iniciado
 * (eso lo hace portal-iva-descarga) o si el libro no es de la sociedad/periodo.
 * Nunca presenta, ni elimina, ni toca los datos iniciales.
 *
 * Salida: JSON por stdout; progreso por stderr. Log CSV (y capturas en simulacion).
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { resolveCredentials, envFile } = require('../lib/env');
const { login } = require('../lib/arca-login');
const { readWorkbook } = require('../lib/xlsx-min');
const rutas = require('../lib/rutas-bva');
const pi    = require('../lib/portal-iva');

const args = parseArgs();
const EJECUTAR = 'ejecutar' in args;
const LIMITE = args['limite'] != null ? Number(args['limite']) : Infinity;
const CRED_FISCAL = String(args['cred-fiscal'] || 'arca').toLowerCase();
const AUNQUE_PRESENTADA = 'aunque-presentada' in args;
if ('ver' in args && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';
if (!(LIMITE > 0)) fail('--limite tiene que ser un numero mayor a 0');
if (!['arca', 'iva', '0'].includes(CRED_FISCAL)) fail('--cred-fiscal acepta arca, iva o 0');

const log = (m) => process.stderr.write(`[portal-iva-agregar] ${m}\n`);

function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)))          return s;
  return null;
}
function mesAnterior(hoy = new Date()) {
  const d = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
}
const PERIODO = args['periodo'] ? normalizarPeriodo(args['periodo']) : mesAnterior();
if (!PERIODO) fail('--periodo tiene que ser MM/AAAA o AAAAMM');

const sinAcentos = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
const PROVEEDORES = String(args['proveedores'] || 'EDESUR;METROGAS;AYSA;AGUA Y SANEAMIENTOS')
  .split(';').map((s) => sinAcentos(s).trim()).filter(Boolean);
const TODOS = PROVEEDORES.includes('*');

let CUIT, CUIT_LOGIN, PASSWORD, TITULAR;
try {
  ({ cuit: CUIT, cuitLogin: CUIT_LOGIN, password: PASSWORD, titular: TITULAR } =
    resolveCredentials({ cuit: args['cuit'], nombre: args['nombre'] }));
} catch (e) {
  fail(e.message, { sugerencia: `Revisa el .env (${envFile() || 'no encontrado'}) o el Excel de claves.` });
}

// ── 1. Comprobantes a cargar ───────────────────────────────────────────────
// Codigos de ARCA por TC y letra de Zetti.
const TIPOS = { FC: { A: 1, B: 6, C: 11 }, ND: { A: 2, B: 7, C: 12 }, NC: { A: 3, B: 8, C: 13 } };
// Alicuota -> sufijo de los campos cm_netoGravadoIVA* / cm_importeIVA*.
const ALICUOTAS = [[0.27, '27'], [0.21, '21'], [0.105, '105'], [0.05, '5'], [0.025, '25']];

const num = (v) => { const n = Number(String(v ?? '').trim() || 0); return Number.isFinite(n) ? n : NaN; };
const r2 = (n) => Math.round(n * 100) / 100;

// La fecha llega como serial de Excel (Zetti) o como texto dd/mm/aaaa.
function fechaDMY(v) {
  const s = String(v ?? '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return `${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[3]}`;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/)))      return `${m[3]}/${m[2]}/${m[1]}`;
  if (/^\d+(\.\d+)?$/.test(s)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(s)) * 86400000);
    return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
  }
  return null;
}

function ubicarArchivo() {
  if (args['archivo']) return path.resolve(String(args['archivo']));
  // En la carpeta IVA del mes; a veces queda suelto en la carpeta del mes.
  const nombre = `${PERIODO} - ZETTI - COMPRAS.xlsx`;
  const iva = rutas.carpetaIva(CUIT, PERIODO, { crear: false });
  const candidatos = [path.join(iva, nombre), path.join(path.dirname(iva), nombre)];
  const archivo = candidatos.find((a) => fs.existsSync(a));
  if (!archivo) throw new Error(`No esta ${candidatos[0]}. Bajalo primero con zetti-subdiario-iva-compras, o pasa --archivo.`);
  return archivo;
}

// Lee la hoja (la de "Compras Zetti" o la primera), ubica el encabezado por
// sus titulos y devuelve los comprobantes de los proveedores pedidos.
function leerComprobantes(archivo) {
  const wb = readWorkbook(fs.readFileSync(archivo));
  const hoja = wb.sheetNames.find((n) => /compras zetti/i.test(n)) || wb.sheetNames[0];
  const filas = wb.rows(hoja);
  const iCab = filas.findIndex((f) => f.some((c) => /^nro\.?\s*comp/i.test(String(c).trim())));
  if (iCab < 0) throw new Error(`No encontre el encabezado (Nro. Comp.) en la hoja "${hoja}" de ${archivo}.`);
  const cab = filas[iCab].map((c) => sinAcentos(c).trim());
  const col = (re, req = true) => {
    const i = cab.findIndex((c) => re.test(c));
    if (i < 0 && req) throw new Error(`Falta la columna ${re} en ${archivo}.`);
    return i;
  };
  const c = {
    fecha: col(/^FECHA$/), tc: col(/^TC$/), letra: col(/^M$/), nro: col(/^NRO\.?\s*COMP/),
    prov: col(/^PROVEEDOR$/), cuit: col(/^CUIT$/), exen: col(/^EXEN/), grav: col(/^GRAV/),
    iva: col(/^IVA$/), piibb: col(/IIBB/), piva: col(/^P\/R IVA$/), total: col(/^TOTAL$/),
    pig: col(/^P\/R IG$/, false),
  };
  // Formato viejo del papel: "Nro. Comp." trae el punto de venta y la columna
  // de al lado, sin titulo, el numero.
  const nroPartido = cab[c.nro + 1] === '';

  const out = [];
  for (const [i, f] of filas.slice(iCab + 1).entries()) {
    const proveedor = String(f[c.prov] || '').trim();
    if (!proveedor) continue;   // fila de total o vacia
    const nombre = sinAcentos(proveedor);
    if (!TODOS && !PROVEEDORES.some((p) => nombre.includes(p))) continue;

    const it = {
      fila: iCab + 2 + i, proveedor, cuit: String(f[c.cuit] || '').replace(/\D/g, ''),
      fecha: fechaDMY(f[c.fecha]), tc: String(f[c.tc] || '').trim().toUpperCase(),
      letra: String(f[c.letra] || '').trim().toUpperCase(),
      comprobante: nroPartido ? `${String(f[c.nro] || '').trim()}-${String(f[c.nro + 1] || '').trim()}` : String(f[c.nro] || '').trim(),
      exento: Math.abs(num(f[c.exen])), neto: Math.abs(num(f[c.grav])), iva: Math.abs(num(f[c.iva])),
      perc_iibb: Math.abs(num(f[c.piibb])), perc_iva: Math.abs(num(f[c.piva])),
      perc_ig: c.pig >= 0 ? Math.abs(num(f[c.pig])) : 0, total: Math.abs(num(f[c.total])),
      tipo: null, pv: null, nro: null, alicuota: null, errores: [],
    };
    it.tipo = (TIPOS[it.tc] || {})[it.letra] || null;
    if (!it.tipo) it.errores.push(`TC ${it.tc} ${it.letra} no se carga (solo FC/ND/NC A, B o C)`);
    const m = it.comprobante.match(/^(\d{1,5})-(\d{1,8})$/);
    if (m) { it.pv = Number(m[1]); it.nro = Number(m[2]); } else it.errores.push(`numero "${it.comprobante}" no es PPPP-NNNNNNNN`);
    if (!/^\d{11}$/.test(it.cuit)) it.errores.push(`CUIT "${it.cuit}" invalido`);
    if (!it.fecha) it.errores.push('fecha ilegible');
    else if (Number(it.fecha.slice(6) + it.fecha.slice(3, 5)) > Number(PERIODO)) it.errores.push(`fecha ${it.fecha} posterior al periodo`);
    for (const k of ['exento', 'neto', 'iva', 'perc_iibb', 'perc_iva', 'perc_ig', 'total']) if (Number.isNaN(it[k])) it.errores.push(`${k} no es numero`);
    if (it.neto > 0 || it.iva > 0) {
      const hit = ALICUOTAS.find(([a]) => Math.abs(it.neto * a - it.iva) <= 0.05);
      if (hit) it.alicuota = hit; else it.errores.push(`IVA ${it.iva} no corresponde a una sola alicuota sobre ${it.neto}: cargalo a mano`);
    }
    if (it.letra === 'C' && it.iva > 0) it.errores.push('comprobante C con IVA');
    const suma = r2(it.exento + it.neto + it.iva + it.perc_iibb + it.perc_iva + it.perc_ig);
    if (Math.abs(suma - it.total) > 0.02) it.errores.push(`los importes suman ${suma} y el total es ${it.total}`);
    out.push(it);
  }
  return { hoja, comprobantes: out };
}

// ── 2. La DDJJ del periodo ya esta presentada? ──────────────────────────────
// Acuse "AAAAMM-IVA-ACUSE-..." o la DDJJ "AAAAMM-iva-dj-..." / "AAAAMM-dj-iva-..." en la carpeta del mes.
function ddjjPresentada() {
  let mes;
  try { mes = rutas.carpetaMes(path.join(rutas.carpetaSociedad(CUIT), '01-Impuestos Mensuales'), PERIODO); } catch { return null; }
  if (!fs.existsSync(mes)) return null;
  const re = new RegExp(`^${PERIODO}.*(\\bIVA\\b.*\\b(ACUSE|DJ|DDJJ)\\b|\\b(DJ|DDJJ)\\b.*\\bIVA\\b)`, 'i');
  const dirs = [mes, ...fs.readdirSync(mes, { withFileTypes: true }).filter((e) => e.isDirectory()).map((e) => path.join(mes, e.name))];
  for (const d of dirs) {
    const hit = fs.readdirSync(d).find((f) => re.test(f.replace(/[_-]/g, ' ')) && !/\bvep\b/i.test(f.replace(/[_-]/g, ' ')));
    if (hit) return path.relative(mes, path.join(d, hit));
  }
  return null;
}

// ── 3. Navegacion ────────────────────────────────────────────────────────────
// Igual que pi.abrirBorrador, pero antes de elegir el periodo se fija que ARCA
// lo ofrezca: si ya se presento, el desplegable de "Nueva declaracion jurada" no
// lo trae (solo el siguiente) y no hay borrador donde cargar.
async function abrirBorrador(page) {
  log(`abriendo borrador del periodo ${PERIODO}`);
  await page.waitForSelector('button:has-text("INGRESAR"), a:has-text("INGRESAR")', { timeout: 30000 });
  await pi.clickBotonDeTarjeta(page, /Nueva declaraci/i, /INGRESAR/i, 'INGRESAR de Nueva declaracion jurada');
  const select = page.locator('select').first();
  await select.waitFor({ state: 'visible', timeout: 30000 });
  await page.waitForFunction(() => document.querySelectorAll('select option').length > 1, null, { timeout: 15000 }).catch(() => {});
  const opciones = await select.evaluate((s) => Array.from(s.options).map((o) => o.value).filter(Boolean));
  if (!opciones.includes(PERIODO)) {
    const e = new Error(`ARCA no ofrece el periodo ${PERIODO} (ofrece ${opciones.join(', ') || 'ninguno'}): ` +
      'la DDJJ de IVA de ese mes ya esta presentada. Para agregar comprobantes habria que rectificar, y eso se decide y se hace a mano.');
    e.code = 'periodo_presentado';
    throw e;
  }
  await select.selectOption(PERIODO);
  await pi.clickSeguro(page.locator('a, button').filter({ hasText: /CONTINUAR/i }).first(), 'CONTINUAR del periodo');
  await page.waitForTimeout(2000);
  const cuerpo = await page.innerText('body').catch(() => '');
  const otro = cuerpo.match(/borrador sin presentar, correspondiente a otro per.odo \((\d{2}\/\d{4})\)/i);
  if (otro) throw Object.assign(new Error(`ARCA tiene abierto un borrador de ${otro[1]}: no toco nada.`), { code: 'otro_borrador_abierto' });
  await page.waitForSelector('button:has-text("INGRESAR"), a:has-text("INGRESAR")', { timeout: 30000 });
  await pi.clickBotonDeTarjeta(page, /Registraci.n y declaraci/i, /INGRESAR/i, 'INGRESAR de Registracion y declaracion');
  await page.waitForTimeout(3000);
}

// Todas las filas del Libro Compras (misma lectura que eliminar-ncr.js de
// conciliacion-ncr). Indices de la fila en compras.js de ARCA:
//   0 fecha · 1 tipo · 3 PV · 4 numero · 11 CUIT emisor · 12 denominacion · 23 total · 25 idReg
async function leerLibro(page) {
  await page.waitForFunction(() => window.jQuery && jQuery.fn.dataTable &&
    jQuery.fn.dataTable.isDataTable('#tablaDataTables') &&
    (document.querySelector('#tablaDataTables tbody tr .botonEliminarComprobante') ||
     /Ning.n dato|No hay/i.test(document.body.innerText)),
  null, { timeout: 180000, polling: 1000 });

  return page.evaluate(async () => {
    const dt = jQuery('#tablaDataTables').DataTable();
    const get = (data) => new Promise((res, rej) => jQuery.ajax({ url: 'ajax.do', data, dataType: 'json',
      timeout: 120000, success: res, error: (x) => rej(new Error(`HTTP ${x.status} leyendo el libro`)) }));
    let rows, totalDeclarado;
    if (!window.gServerSide) {
      rows = dt.rows().data().toArray();
      totalDeclarado = rows.length;
    } else {
      rows = [];
      const TANDA = 500;
      for (let draw = 1; ; draw++) {
        const r = await get({ f: 'listaComprobantesIncluidosSS', c: window.gTipoOperacion, draw,
                              start: rows.length, length: TANDA,
                              'search[value]': '', 'search[regex]': false,
                              'order[0][column]': 0, 'order[0][dir]': 'asc' });
        const d = (r && r.datos) || {};
        const data = d.data || [];
        totalDeclarado = Number(d.recordsTotal ?? r.recordsTotal ?? d.recordsFiltered ?? r.recordsFiltered ?? NaN);
        rows.push(...data);
        if (!data.length || rows.length >= totalDeclarado || data.length < TANDA) break;
        if (draw > 200) throw new Error('Demasiadas paginas leyendo el libro.');
      }
    }
    const enPantalla = document.body.innerText.match(/de un total de ([\d.]+)/i);
    if (enPantalla) totalDeclarado = Number(enPantalla[1].replace(/\./g, ''));

    let validacion = rows.length ? 'sin filas visibles' : 'ok';
    const tr = document.querySelector('#tablaDataTables tbody tr');
    const btn = tr && tr.querySelector('.botonEliminarComprobante');
    if (btn) {
      const r = dt.row(tr).data();
      const nroVisible = tr.children[2].innerText.replace(/\D/g, '');
      const ok = String(r[25]) === String(btn.getAttribute('data-id-reg')) &&
                 Number(nroVisible) === Number(String(r[3]) + String(r[4]).padStart(8, '0')) &&
                 /^\d{11}$/.test(String(r[11])) && /^\d{2}\/\d{2}\/\d{4}$/.test(String(r[0]));
      validacion = ok ? 'ok' : `no coincide: fila=${JSON.stringify(r).slice(0, 300)} nroVisible=${nroVisible}`;
    }
    if (rows.length !== totalDeclarado) validacion = `lei ${rows.length} filas y ARCA declara ${totalDeclarado}`;
    return {
      validacion, estado: window.gEstadoPresentacion, cuit: window.gUserEmpresaRepresentadaCUIT,
      mesInicio: window.gConfigFechas && window.gConfigFechas.fechaMesInicio,
      filas: rows.map((r) => ({ fecha: r[0], tipo: Number(r[1]), pv: Number(r[3]), nro: Number(r[4]),
                                cuit: String(r[11]), nombre: r[12], total: Number(r[23]), idReg: String(r[25]) })),
    };
  });
}

const clave = (cuit, tipo, pv, nro) => `${cuit}|${Number(tipo)}|${Number(pv)}|${Number(nro)}`;

// ── 4. El formulario de AGREGAR COMPROBANTE ──────────────────────────────────
async function abrirFormulario(page) {
  await pi.cerrarModales(page);
  const boton = page.locator('a, button').filter({ hasText: /AGREGAR COMPROBANTE/i }).first();
  await boton.waitFor({ state: 'visible', timeout: 30000 });
  await pi.clickSeguro(boton, 'AGREGAR COMPROBANTE...');
  await page.locator('#modalAltaComprobante #cm_tipoComprobante').waitFor({ state: 'visible', timeout: 20000 });
  await page.waitForTimeout(800);
}

// Completa el formulario con jQuery (dispara los mismos change/keyup que al
// tipear, asi corren los calculos de ARCA) y devuelve lo que quedo cargado.
async function completarFormulario(page, it) {
  const fmt = (n) => r2(n).toFixed(2);
  const datos = {
    tipo: String(it.tipo), fecha: it.fecha, pv: String(it.pv), nro: String(it.nro), cuit: it.cuit,
    denominacion: it.proveedor.replace(/\s+/g, ' ').slice(0, 50),
    exento: it.exento ? fmt(it.exento) : '', perc_iva: it.perc_iva ? fmt(it.perc_iva) : '',
    perc_iibb: it.perc_iibb ? fmt(it.perc_iibb) : '', perc_ig: it.perc_ig ? fmt(it.perc_ig) : '', total: fmt(it.total),
    sufijo: it.alicuota ? it.alicuota[1] : null, neto: it.neto ? fmt(it.neto) : '', iva: it.iva ? fmt(it.iva) : '',
    credFiscal: CRED_FISCAL === 'iva' ? fmt(it.iva) : CRED_FISCAL === '0' ? '0.00' : null,
  };

  // Tipo primero: muestra u oculta la seccion de IVA segun el comprobante.
  await page.evaluate((d) => {
    const $ = window.jQuery;
    $('#modalAltaComprobante input.soloNumerosConDecimales').val('');
    $('#cm_tipoComprobante').val(d.tipo).trigger('change');
  }, datos);
  await page.waitForTimeout(600);

  await page.evaluate((d) => {
    const $ = window.jQuery;
    const set = (id, v) => $('#' + id).val(v).trigger('change').trigger('keyup');
    set('cm_fechaEmision', d.fecha);
    try { $('#cm_fechaEmision').datepicker('update', d.fecha); } catch (e) { /* sin datepicker */ }
    set('cm_ptoVta', d.pv);
    set('cm_numero', d.nro);
    $('#cm_denominacionEmisor').val('');
    set('cm_cuitEmisor', d.cuit);   // dispara la busqueda de la razon social
    set('cm_codMoneda', 'PES');
    set('cm_tipoCambio', '1');
  }, datos);

  // La razon social la trae ARCA por el CUIT; si no aparece, la de Zetti.
  const razon = await page.waitForFunction(() => (window.jQuery('#cm_denominacionEmisor').val() || '').trim(),
    null, { timeout: 10000 }).then((h) => h.jsonValue()).catch(() => '');
  if (!razon) await page.evaluate((v) => window.jQuery('#cm_denominacionEmisor').val(v).trigger('change'), datos.denominacion);

  await page.evaluate((d) => {
    const $ = window.jQuery;
    const set = (id, v) => $('#' + id).val(v).trigger('change').trigger('keyup');
    set('cm_importeExento', d.exento);
    set('cm_percepcionesIVA', d.perc_iva);
    set('cm_percepcionesIIBB', d.perc_iibb);
    set('cm_percepcionesImpNac', d.perc_ig);         // percepciones de Ganancias
    if (d.sufijo) {
      set('cm_netoGravadoIVA' + d.sufijo, d.neto);   // ARCA calcula neto * alicuota...
      set('cm_importeIVA' + d.sufijo, d.iva);        // ...y se deja el IVA de la factura
    }
    set('cm_importeTotal', d.total);
    if (d.credFiscal !== null && !$('#cm_credFiscalComp').prop('readonly')) set('cm_credFiscalComp', d.credFiscal);
  }, datos);
  await page.waitForTimeout(500);

  return page.evaluate(() => {
    const $ = window.jQuery;
    const v = (id) => ($('#' + id).val() || '').trim();
    return {
      tipo: v('cm_tipoComprobante'), fecha: v('cm_fechaEmision'), pv: v('cm_ptoVta'), nro: v('cm_numero'),
      cuit: v('cm_cuitEmisor'), denominacion: v('cm_denominacionEmisor'),
      receptor: v('cm_nroDocReceptor'), tipoDocReceptor: $('#cm_textoTipoDocumentoReceptor').text().trim(),
      moneda: v('cm_codMoneda'), exento: v('cm_importeExento'), perc_iva: v('cm_percepcionesIVA'),
      perc_iibb: v('cm_percepcionesIIBB'), perc_ig: v('cm_percepcionesImpNac'),
      total: v('cm_importeTotal'), credFiscal: v('cm_credFiscalComp'),
      credFiscalEditable: !$('#cm_credFiscalComp').prop('readonly'),
      ivas: ['25', '5', '105', '21', '27'].map((s) => ({ s, neto: v('cm_netoGravadoIVA' + s), iva: v('cm_importeIVA' + s) }))
        .filter((x) => Number(x.neto) || Number(x.iva)),
    };
  });
}

// Lo que quedo en el formulario tiene que ser lo que se quiso cargar.
function controlarFormulario(it, f) {
  const dif = [];
  const igualN = (a, b) => Math.abs(Number(a || 0) - Number(b || 0)) <= 0.001;
  if (Number(f.tipo) !== it.tipo) dif.push(`tipo ${f.tipo}`);
  if (f.fecha !== it.fecha) dif.push(`fecha ${f.fecha}`);
  if (Number(f.pv) !== it.pv || Number(f.nro) !== it.nro) dif.push(`numero ${f.pv}-${f.nro}`);
  if (f.cuit.replace(/\D/g, '') !== it.cuit) dif.push(`CUIT emisor ${f.cuit}`);
  if (f.receptor.replace(/\D/g, '') !== String(CUIT) || f.tipoDocReceptor !== 'CUIT') dif.push(`receptor ${f.tipoDocReceptor} ${f.receptor}`);
  if (f.moneda !== 'PES') dif.push(`moneda ${f.moneda}`);
  if (!igualN(f.exento, it.exento)) dif.push(`exento ${f.exento}`);
  if (!igualN(f.perc_iva, it.perc_iva)) dif.push(`perc. IVA ${f.perc_iva}`);
  if (!igualN(f.perc_iibb, it.perc_iibb)) dif.push(`perc. IIBB ${f.perc_iibb}`);
  if (!igualN(f.perc_ig, it.perc_ig)) dif.push(`perc. otros imp. nac. ${f.perc_ig}`);
  if (!igualN(f.total, it.total)) dif.push(`total ${f.total}`);
  const esperado = it.alicuota ? [{ s: it.alicuota[1], neto: it.neto, iva: it.iva }] : [];
  if (f.ivas.length !== esperado.length || esperado.some((e) => {
    const x = f.ivas.find((y) => y.s === e.s);
    return !x || !igualN(x.neto, e.neto) || !igualN(x.iva, e.iva);
  })) dif.push(`IVA ${JSON.stringify(f.ivas)}`);
  if (!f.denominacion) dif.push('sin denominacion del emisor');
  return dif;
}

async function cancelarFormulario(page) {
  const cancelar = page.locator('#btnAltaCancelar');
  if (await cancelar.isVisible().catch(() => false)) await cancelar.click().catch(() => {});
  await page.waitForTimeout(800);
  await pi.cerrarModales(page);
}

// Click en "Agregar" y la respuesta de guardarComprobanteManual.do.
async function guardarFormulario(page) {
  const boton = page.locator('#btnAgregarComprobanteManual');
  const respuesta = page.waitForResponse((r) => /guardarComprobanteManual\.do/.test(r.url()), { timeout: 60000 }).catch(() => null);
  await pi.clickSeguro(boton, 'Agregar (guardar comprobante)');
  const r = await respuesta;
  if (!r) {
    // Con NC A/ND A con IVA mayor al neto ARCA pregunta si es canje antes de mandar.
    const dialogo = await page.locator('.bootstrap-dialog .modal-body').first().innerText().catch(() => '');
    return { ok: false, detalle: dialogo ? `ARCA pidio confirmacion: ${dialogo.trim().slice(0, 150)}` : 'ARCA no respondio en 60 s' };
  }
  const json = await r.json().catch(() => null);
  if (!json) return { ok: false, detalle: `respuesta ilegible (HTTP ${r.status()})` };
  const campos = (o) => Object.entries(o || {}).map(([k, v]) => `${k.replace(/^cm_/, '')}: ${[].concat(v).join(' / ')}`).join('; ');
  if (json.estado === 'ok') {
    return { ok: true, detalle: json.mensajeObservacion ? `observacion: ${json.mensajeObservacion} ${campos(json.camposObservaciones)}`.trim() : '' };
  }
  return { ok: false, detalle: `${json.mensajeError || 'error'} ${campos(json.camposErrores)}`.trim() };
}

// ── principal ──────────────────────────────────────────────────────────────
(async () => {
  let ARCHIVO, ENTRADA;
  try {
    ARCHIVO = ubicarArchivo();
    ENTRADA = leerComprobantes(ARCHIVO);
  } catch (e) { fail(e.message, { code: 'entrada' }); }

  const resultados = ENTRADA.comprobantes.map((it) => ({
    ...it, accion: it.errores.length ? 'error_datos' : null, detalle: it.errores.join('; '), cred_fiscal: null,
  }));
  log(`${TITULAR || CUIT} · CUIT ${CUIT} · periodo ${PERIODO} · ${EJECUTAR ? 'MODO EJECUCION' : 'SIMULACION (no guarda nada)'}`);
  log(`${path.basename(ARCHIVO)}: ${resultados.length} comprobantes de ${TODOS ? 'todos los proveedores' : PROVEEDORES.join(' / ')}` +
      (resultados.some((r) => r.accion) ? ` (${resultados.filter((r) => r.accion).length} con errores de datos)` : ''));

  const salida = (extra = {}) => {
    const porAccion = {};
    for (const r of resultados) porAccion[r.accion || 'sin_procesar'] = (porAccion[r.accion || 'sin_procesar'] || 0) + 1;
    return { modo: EJECUTAR ? 'ejecucion' : 'simulacion', titular: TITULAR || null, cuit: CUIT, periodo: PERIODO,
      archivo: ARCHIVO, proveedores: TODOS ? '*' : PROVEEDORES, por_accion: porAccion,
      comprobantes: resultados.map((r) => ({ proveedor: r.proveedor, fecha: r.fecha, tipo: r.tipo, comprobante: r.comprobante,
        total: r.total, accion: r.accion, detalle: r.detalle || undefined, cred_fiscal: r.cred_fiscal ?? undefined })),
      ...extra };
  };

  if (!resultados.length) { console.log(JSON.stringify({ ok: true, ...salida(), aviso: 'No hay comprobantes de esos proveedores en el archivo.' }, null, 2)); return; }
  if (!resultados.some((r) => !r.accion)) { console.log(JSON.stringify({ ok: false, ...salida(), error: 'Ningun comprobante paso los controles de datos.' }, null, 2)); process.exit(1); }

  const presentada = ddjjPresentada();
  if (presentada && !AUNQUE_PRESENTADA) {
    fail(`La DDJJ de IVA ${PERIODO} ya esta presentada (${presentada}). Agregar comprobantes implica rectificar: eso se hace a mano.`,
      { code: 'ddjj_presentada', ...salida() });
  }

  const sello = (() => { const d = new Date(), z = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`; })();
  const dirSalidas = path.join(process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'portal-iva-agregar');
  const dirCapturas = path.join(dirSalidas, `${PERIODO}-${CUIT}-${sello}`);

  const browser = await chromium.launch(launchOptions());
  const context = await browser.newContext({ locale: 'es-AR' });
  const page = await context.newPage();
  let verificacion = null, repr = null;
  try {
    await login(page, { cuitLogin: CUIT_LOGIN, password: PASSWORD });
    const iva = await pi.abrirPortalIva(context, page);
    repr = await pi.asegurarRepresentacion(iva, CUIT);
    await abrirBorrador(iva);
    if (/verDatosInicialesPresentacion/i.test(iva.url())) {
      throw Object.assign(new Error('El borrador del periodo no esta iniciado (ARCA pide datos iniciales). Primero corre portal-iva-descarga, que lo inicia e importa las compras.'), { code: 'borrador_sin_iniciar' });
    }
    await pi.abrirLibro(iva, 'compras');
    await pi.cerrarModales(iva);

    // Controles antes de tocar nada.
    const antes = await leerLibro(iva);
    const mmAAAA = `${PERIODO.slice(4, 6)}/${PERIODO.slice(0, 4)}`;
    if (antes.estado !== 'BO') throw new Error(`El libro no esta en Borrador (estado ${antes.estado}). No se puede agregar.`);
    if (String(antes.cuit) !== String(CUIT)) throw new Error(`El libro abierto es del CUIT ${antes.cuit}, no de ${CUIT}.`);
    if (!String(antes.mesInicio || '').endsWith(mmAAAA)) throw new Error(`El libro abierto es de ${antes.mesInicio}, no de ${mmAAAA}.`);
    if (antes.validacion !== 'ok') throw new Error(`La grilla de ARCA cambio (${antes.validacion}). Hay que revisar los indices.`);
    log(`libro: ${antes.filas.length} comprobantes`);

    // Cruce: lo que ya esta en el libro no se carga.
    const indice = new Map();
    for (const f of antes.filas) indice.set(clave(f.cuit, f.tipo, f.pv, f.nro), f);
    for (const r of resultados.filter((x) => !x.accion)) {
      const f = indice.get(clave(r.cuit, r.tipo, r.pv, r.nro));
      if (!f) { r.accion = 'cargar'; continue; }
      r.accion = 'ya_esta';
      r.detalle = Math.abs(Math.abs(f.total) - r.total) > 0.01 ? `en el libro con total ${f.total} (Zetti ${r.total})` : 'importado desde ARCA';
    }
    const aCargar = resultados.filter((r) => r.accion === 'cargar').slice(0, LIMITE);
    log(`cruce: cargar=${resultados.filter((r) => r.accion === 'cargar').length} ya_esta=${resultados.filter((r) => r.accion === 'ya_esta').length}` +
        (LIMITE !== Infinity ? ` (limite ${aCargar.length})` : ''));

    if (!EJECUTAR && aCargar.length) fs.mkdirSync(dirCapturas, { recursive: true });
    for (const r of aCargar) {
      const etiqueta = `${r.proveedor} ${r.tc} ${r.letra} ${r.comprobante}`;
      await abrirFormulario(iva);
      const form = await completarFormulario(iva, r);
      r.cred_fiscal = form.credFiscal || '0';
      const dif = controlarFormulario(r, form);
      if (dif.length) {
        r.accion = 'error_formulario'; r.detalle = `el formulario quedo distinto: ${dif.join('; ')}`;
        log(`${etiqueta}: ${r.detalle}. No lo guardo.`);
        await cancelarFormulario(iva);
        continue;
      }
      if (!EJECUTAR) {
        const archivo = path.join(dirCapturas, `${r.pv}-${r.nro}.png`);
        // El modal es mas alto que la ventana: se agranda para que la captura lo muestre entero.
        const vista = iva.viewportSize();
        await iva.setViewportSize({ width: 1280, height: 2600 }).catch(() => {});
        await iva.waitForTimeout(500);
        await iva.locator('#modalAltaComprobante .modal-content').screenshot({ path: archivo }).catch(() => {});
        if (vista) await iva.setViewportSize(vista).catch(() => {});
        r.accion = 'cargaria'; r.detalle = `formulario completo y controlado; credito fiscal ${r.cred_fiscal}${form.credFiscalEditable ? ' (editable)' : ''}`;
        log(`${etiqueta}: cargaria (captura ${path.basename(archivo)})`);
        await cancelarFormulario(iva);
        continue;
      }
      const g = await guardarFormulario(iva);
      if (g.ok) {
        r.accion = 'agregado'; r.detalle = g.detalle;
        log(`${etiqueta}: agregado${g.detalle ? ' · ' + g.detalle : ''}`);
        await iva.waitForTimeout(1500);
        await pi.cerrarModales(iva);
      } else {
        r.accion = 'rechazado'; r.detalle = g.detalle;
        log(`${etiqueta}: ARCA lo rechazo · ${g.detalle}`);
        await cancelarFormulario(iva);
      }
    }

    // Verificacion: se relee el libro desde ARCA.
    if (EJECUTAR && resultados.some((r) => r.accion === 'agregado')) {
      await iva.reload({ waitUntil: 'domcontentloaded' });
      await iva.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
      await pi.cerrarModales(iva);
      const despues = await leerLibro(iva);
      const idx = new Map(despues.filas.map((f) => [clave(f.cuit, f.tipo, f.pv, f.nro), f]));
      const agregados = resultados.filter((r) => r.accion === 'agregado');
      for (const r of agregados) {
        const f = idx.get(clave(r.cuit, r.tipo, r.pv, r.nro));
        if (!f) { r.accion = 'error'; r.detalle = `ARCA dijo ok pero no esta en el libro. ${r.detalle}`.trim(); }
        else if (Math.abs(Math.abs(f.total) - r.total) > 0.01) { r.accion = 'error'; r.detalle = `quedo en el libro con total ${f.total}`; }
      }
      verificacion = {
        antes: antes.filas.length, despues: despues.filas.length, agregados: agregados.length,
        cuadra: antes.filas.length + agregados.length === despues.filas.length && !agregados.some((r) => r.accion === 'error'),
      };
      log(`verificacion: ${verificacion.antes} -> ${verificacion.despues} comprobantes · ${verificacion.cuadra ? 'OK' : 'NO CUADRA'}`);
    }
  } catch (e) {
    await browser.close().catch(() => {});
    if (e.code === 'captcha') fail('ARCA pidio un captcha.', { code: 'captcha' });
    if (e.code === 'credenciales_invalidas') fail(`ARCA rechazo las credenciales del CUIT ${CUIT}. No reintentes: bloquea la clave.`, { code: e.code });
    fail(e.message, { code: e.code || 'error', ...salida() });
  }
  await browser.close().catch(() => {});

  // Log: archivo nuevo, nunca se pisa. Al ejecutar, en la carpeta IVA del periodo.
  const dirLog = EJECUTAR ? rutas.carpetaIva(CUIT, PERIODO, { crear: false }) : dirSalidas;
  fs.mkdirSync(dirLog, { recursive: true });
  const rutaLog = path.join(dirLog, `${PERIODO}-CARGA-MANUAL-PORTAL-${sello}${EJECUTAR ? '' : '-SIMULACION'}.csv`);
  const cols = ['fila', 'proveedor', 'cuit', 'fecha', 'tc', 'letra', 'comprobante', 'tipo', 'pv', 'nro', 'exento', 'neto',
                'alicuota', 'iva', 'perc_iibb', 'perc_iva', 'perc_ig', 'total', 'cred_fiscal', 'accion', 'detalle'];
  const valor = (r, c) => c === 'alicuota' ? (r.alicuota ? r.alicuota[0] * 100 + '%' : '') : r[c];
  const csv = [cols.join(';'), ...resultados.map((r) => cols.map((c) => String(valor(r, c) ?? '').replace(/;/g, ',')).join(';'))].join('\r\n');
  if (fs.existsSync(rutaLog)) fail(`Ya existe ${rutaLog}; no lo piso.`);
  fs.writeFileSync(rutaLog, '﻿' + csv, 'utf-8');

  const ok = !resultados.some((r) => ['error', 'rechazado', 'error_formulario'].includes(r.accion)) && (!verificacion || verificacion.cuadra);
  console.log(JSON.stringify({ ok, ...salida({ representando: repr && repr.linea, verificacion, log: rutaLog,
    capturas: !EJECUTAR && fs.existsSync(dirCapturas) ? dirCapturas : undefined }) }, null, 2));
  process.exit(ok ? 0 : 1);
})();
