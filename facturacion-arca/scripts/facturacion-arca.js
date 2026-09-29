#!/usr/bin/env node
/**
 * facturacion-arca.js — emite facturas de servicios en Comprobantes en Línea
 * (RCEL) de ARCA para cualquier emisor: Monotributo (Factura C) o Responsable
 * Inscripto (Factura A / B). Es la skill facturacion-arca hecha script.
 *
 * ⚠️  ESCRIBE EN ARCA. Cada factura sale con CAE, es un comprobante fiscal real
 *     y sólo se anula con una nota de crédito. Por eso hay tres modos:
 *
 *   1. --validar   No entra a ARCA. Lee la planilla, controla y muestra el
 *                  resumen con el total y un CÓDIGO de aprobación.
 *   2. (default)   Dry-run: entra a ARCA, recorre el asistente de cada factura
 *                  hasta la pantalla de confirmación, la controla contra la
 *                  planilla, saca captura y NO confirma. Da el mismo código.
 *   3. --emitir=<código>  Emite. El código tiene que coincidir con el del
 *                  resumen aprobado: si la planilla cambió (otra fila, otro
 *                  importe, otra fecha) el código cambia y no emite nada.
 *
 * Desde planilla (modelo "Plantilla Facturacion ARCA.xlsx"):
 *   node scripts/facturacion-arca.js --planilla="/ruta/Facturacion X.xlsx" --hoja="Octubre 2026" --validar
 *   node scripts/facturacion-arca.js --planilla="/ruta/Facturacion X.xlsx" --hoja="Octubre 2026"
 *   node scripts/facturacion-arca.js --planilla="/ruta/Facturacion X.xlsx" --hoja="Octubre 2026" --emitir=1A2B3C4D
 *
 * Una sola factura por argumentos:
 *   node scripts/facturacion-arca.js --cuit=<emisor> --pv=1 --condicion=mono --actividad=692000 \
 *     --receptor=<cuit cliente> --cond-receptor=ri --desc="Honorarios septiembre 2026" --importe=150000 \
 *     --fecha=30/09/2026 --desde=01/09/2026 --hasta=30/09/2026 --vto=10/10/2026 --cond-venta="Transferencia Bancaria"
 *
 * Argumentos:
 *   --planilla, --hoja   Planilla mensual y hoja del mes (si hay una sola, --hoja se puede omitir).
 *   --filas=5,7          Sólo esas filas de la hoja.
 *   --cuit / --nombre    Emisor (en modo planilla sale de la hoja Emisor).
 *   --emisor-nombre      Nombre del emisor como figura en ARCA (para elegir la empresa en RCEL).
 *   --pv, --condicion (mono|ri), --actividad, --alicuota (21, 10.5, 27)
 *   --receptor, --cond-receptor (ri|mono|exento|cf), --cliente (razón social, opcional)
 *   --desc, --importe    Importe NETO (en Factura C es el total).
 *   --fecha, --desde, --hasta, --vto (dd/mm/aaaa), --cond-venta
 *   --salida             Carpeta de los PDF si el emisor no tiene carpeta propia
 *                        (default ~/Documents/BVA-salidas/facturas/<EMISOR>, o $BVA_SALIDAS_PATH).
 *   --ver                Navegador visible (sólo si se pide).
 *
 * Login: igual que anticipos-sct — credenciales del .env (ARCA_<cuit>_LOGIN /
 * _PASSWORD) o de Claves_Organismos.xlsx, y lib/arca-login.js. Nunca se imprime
 * ni se guarda una clave.
 *
 * Registro: BVA-salidas/facturas/.registro-<cuit>.json guarda cada factura
 * emitida. Una fila ya emitida no se vuelve a emitir aunque la planilla no haya
 * quedado marcada; una que quedó "confirmando" (se cortó justo al emitir) frena
 * todo hasta que alguien verifique a mano en RCEL → Consultas.
 */

const fs     = require('fs');
const path   = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { resolveCredentials, envFile } = require('../lib/env');
const { login, abrirServicio } = require('../lib/arca-login');
const rcelLib = require('../lib/rcel');

const args = parseArgs();
const VER     = 'ver' in args;
const VALIDAR = 'validar' in args;
const EMITIR  = 'emitir' in args;
const CODIGO_PEDIDO = String(args['emitir'] || '').trim().toUpperCase();
if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';

const log = (m) => process.stderr.write(`[facturacion] ${m}\n`);
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
// Fuera de la carpeta del plugin: una actualización la reemplaza entera, y con
// ella se perdería el registro de lo emitido (la protección contra duplicados).
const SALIDAS = process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas');
const RAIZ_SALIDA = path.join(SALIDAS, 'facturas');
const DEBUG_DIR = path.join(RAIZ_SALIDA, '_debug');

if (EMITIR && VALIDAR) fail('--emitir y --validar no van juntos.');
if (EMITIR && !/^[0-9A-F]{8}$/.test(CODIGO_PEDIDO)) {
  fail('--emitir necesita el código de 8 caracteres del resumen aprobado (corré antes con --validar o en dry-run).');
}

// ── Utilidades ────────────────────────────────────────────────────────────────

const soloDigitos = (s) => String(s || '').replace(/\D/g, '');
const pad = (n) => String(n).padStart(2, '0');
const redondear = (n) => Math.round(n * 100) / 100;
const pesos = (n) => '$ ' + Number(n).toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const escapeRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

function cuitValido(c) {
  if (!/^\d{11}$/.test(c)) return false;
  const m = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2];
  const s = m.reduce((a, k, i) => a + k * Number(c[i]), 0);
  let dv = 11 - (s % 11);
  if (dv === 11) dv = 0;
  if (dv === 10) dv = 9;
  return dv === Number(c[10]);
}

function parseFecha(s) {
  const m = String(s || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  if (d.getDate() !== Number(m[1]) || d.getMonth() !== Number(m[2]) - 1) return null;
  return d;
}
const fmtFecha = (d) => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;

function numeroArg(v) {
  if (v === undefined || v === '') return null;
  const s = String(v).replace(/\$|\s/g, '');
  return Number(s.includes(',') ? s.replace(/\./g, '').replace(',', '.') : s);
}

// Alícuota: acepta 0.21, 21, "21%", "10,5".
function parseAlicuota(v) {
  if (v === undefined || v === null || v === '') return null;
  let n = typeof v === 'number' ? v : Number(String(v).replace('%', '').replace(',', '.').trim());
  if (!Number.isFinite(n)) return null;
  if (n > 1) n = n / 100;
  return Math.round(n * 10000) / 10000;
}

function condicionEmisor(txt) {
  const t = String(txt || '').toLowerCase();
  if (/^mono$|monotrib/.test(t)) return 'mono';
  if (/^ri$|responsable inscripto/.test(t)) return 'ri';
  return null;
}

// Condición IVA del receptor → clave interna y regex del <select idivareceptor>.
const COND_RECEPTOR = {
  ri:     { re: /^iva responsable inscripto$/i, nombre: 'IVA Responsable Inscripto' },
  mono:   { re: /^responsable monotributo$/i,   nombre: 'Responsable Monotributo' },
  exento: { re: /^iva sujeto exento$/i,         nombre: 'IVA Sujeto Exento' },
  cf:     { re: /^consumidor final$/i,          nombre: 'Consumidor Final' },
};
function condicionReceptor(txt) {
  const t = String(txt || '').toLowerCase().trim();
  if (!t) return null;
  if (t === 'ri' || /responsable inscripto/.test(t)) return 'ri';
  if (t === 'mono' || /monotrib/.test(t)) return 'mono';
  if (t === 'exento' || /exento/.test(t)) return 'exento';
  if (t === 'cf' || /consumidor final/.test(t)) return 'cf';
  return null;
}

const TIPOS = {
  A: { re: /factura\s*a\b/i, nombre: 'Factura A', codigo: '001' },
  B: { re: /factura\s*b\b/i, nombre: 'Factura B', codigo: '006' },
  C: { re: /factura\s*c\b/i, nombre: 'Factura C', codigo: '011' },
};

// Condición de venta de la planilla → regex del checkbox de RCEL.
function condVentaRe(txt) {
  const t = String(txt || '').toLowerCase();
  if (/transferencia/.test(t)) return /transferencia/i;
  if (/^contado$/.test(t.trim())) return /^contado$/i;
  if (/cuenta corriente/.test(t)) return /cuenta corriente/i;
  if (/cheque/.test(t)) return /cheque/i;
  if (/d[eé]bito/.test(t)) return /tarjeta de d[eé]bito/i;
  if (/cr[eé]dito/.test(t)) return /tarjeta de cr[eé]dito/i;
  if (/^otra?s?$/.test(t.trim())) return /^otr[ao]s?$/i;
  return new RegExp(escapeRe(txt.trim()), 'i');
}

function alicuotaRe(a) {
  const n = String(Math.round(a * 10000) / 100).replace('.', '[.,]');
  return new RegExp(`^${n}\\s*%$`);
}

const nombreArchivo = (s) => String(s || '').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();
const slug = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
  .toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');

// ── Datos: planilla o argumentos ──────────────────────────────────────────────

function leerPlanilla() {
  const ruta = path.resolve(args['planilla']);
  if (!fs.existsSync(ruta)) fail(`No existe la planilla ${ruta}`);
  const py = [path.join(__dirname, 'facturacion_planilla.py'), 'leer', ruta];
  if (args['hoja']) py.push(args['hoja']);
  let out;
  try {
    out = execFileSync(PYTHON, py, { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
  } catch (e) {
    const txt = (e.stdout || e.stderr || e.message || '').toString().trim();
    try { const j = JSON.parse(txt.split(/\r?\n/).pop()); fail(j.error, { hojas: j.hojas }); } catch { /* no era JSON */ }
    fail(`No pude leer la planilla (${PYTHON}): ${txt.split(/\r?\n/).pop()}`);
  }
  return { ruta, ...JSON.parse(out.trim().split(/\r?\n/).pop()) };
}

// Arma { emisor, facturas, errores, avisos }. Todos los controles de la skill
// viven acá, así que --validar, el dry-run y --emitir controlan lo mismo.
function armarLote() {
  const errores = [], avisos = [];
  let emisor, crudas, clientes = [], planilla = null;

  if (args['planilla']) {
    planilla = leerPlanilla();
    const e = planilla.emisor;
    emisor = {
      nombre: args['emisor-nombre'] || e.nombre, cuit: soloDigitos(e.cuit),
      condicion: condicionEmisor(e.condicionIva), condicionTxt: e.condicionIva,
      pv: soloDigitos(args['pv'] || e.puntoVenta), actividad: String(args['actividad'] || e.actividad || '').trim(),
      alicuota: parseAlicuota(args['alicuota'] ?? e.alicuota), carpeta: e.carpeta || '',
    };
    clientes = planilla.clientes;
    const soloFilas = args['filas'] ? new Set(String(args['filas']).split(',').map(Number)) : null;
    crudas = planilla.filas.filter(f => !soloFilas || soloFilas.has(f.fila));
    if (soloFilas) for (const n of soloFilas) if (!crudas.some(f => f.fila === n)) errores.push(`La fila ${n} no tiene CUIT cargado en la hoja ${planilla.hoja}.`);
  } else {
    emisor = {
      nombre: args['emisor-nombre'] || '', cuit: soloDigitos(args['cuit']), nombreClaves: args['nombre'],
      condicion: condicionEmisor(args['condicion']), condicionTxt: args['condicion'] || '',
      pv: soloDigitos(args['pv']), actividad: String(args['actividad'] || '').trim(),
      alicuota: parseAlicuota(args['alicuota']), carpeta: '',
    };
    crudas = [{
      fila: null, cuit: args['receptor'], descripcion: args['desc'], neto: numeroArg(args['importe']),
      fecha: args['fecha'], desde: args['desde'], hasta: args['hasta'], vencimiento: args['vto'],
      condVenta: args['cond-venta'], condReceptor: args['cond-receptor'], cliente: args['cliente'],
      emitida: '', nroComprobante: '',
    }];
  }

  // Emisor
  if (!emisor.cuit && !emisor.nombreClaves) errores.push('Falta el CUIT del emisor (hoja Emisor o --cuit).');
  if (emisor.cuit && !cuitValido(emisor.cuit)) errores.push(`El CUIT del emisor ${emisor.cuit} no es válido.`);
  if (!emisor.condicion) errores.push(`Condición IVA del emisor desconocida ("${emisor.condicionTxt}"): Responsable Monotributo o IVA Responsable Inscripto (--condicion=mono|ri).`);
  if (!emisor.pv) errores.push('Falta el punto de venta del emisor.');
  if (emisor.condicion === 'ri' && !emisor.alicuota) errores.push('Emisor Responsable Inscripto: falta la alícuota de IVA (hoja Emisor o --alicuota).');

  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const facturas = [];
  let yaEmitidas = 0;

  for (const r of crudas) {
    const donde = r.fila ? `Fila ${r.fila}` : 'Factura';
    if (/^s[ií]$/i.test(String(r.emitida || '').trim())) { yaEmitidas++; continue; }

    const cuit = soloDigitos(r.cuit);
    const cli = clientes.find(c => soloDigitos(c.cuit) === cuit);
    if (!/^\d{11}$/.test(cuit)) { errores.push(`${donde}: el CUIT del cliente "${r.cuit || ''}" no tiene 11 dígitos.`); continue; }
    if (!cuitValido(cuit)) errores.push(`${donde}: el CUIT ${cuit} no es válido (dígito verificador).`);
    if (planilla && !cli) errores.push(`${donde}: el CUIT ${cuit} no está cargado en la hoja Clientes.`);

    const condR = condicionReceptor(cli ? cli.condicionIva : r.condReceptor);
    if (!condR) errores.push(`${donde}: falta o no reconozco la condición IVA del cliente ("${(cli && cli.condicionIva) || r.condReceptor || ''}").`);

    let tipoK = null;
    if (emisor.condicion === 'mono') tipoK = 'C';
    else if (emisor.condicion === 'ri' && condR) tipoK = (condR === 'ri' || condR === 'mono') ? 'A' : 'B';

    let descripcion = String(r.descripcion || '').trim();
    if (!descripcion && cli && cli.descripcion) { descripcion = cli.descripcion; avisos.push(`${donde}: sin descripción, uso la habitual del cliente ("${descripcion}").`); }
    if (!descripcion) errores.push(`${donde}: falta la descripción del servicio.`);

    let condVenta = String(r.condVenta || '').trim();
    if (!condVenta && cli && cli.condVenta) { condVenta = cli.condVenta; avisos.push(`${donde}: sin condición de venta, uso la habitual del cliente (${condVenta}).`); }
    if (!condVenta) errores.push(`${donde}: falta la condición de venta.`);

    const neto = typeof r.neto === 'number' ? r.neto : Number(r.neto);
    if (!(neto > 0)) errores.push(`${donde}: el importe "${r.neto ?? ''}" tiene que ser mayor a 0.`);

    const fechas = {};
    for (const [k, etiqueta] of [['fecha', 'fecha de la factura'], ['desde', 'período desde'], ['hasta', 'período hasta'], ['vencimiento', 'vencimiento de pago']]) {
      fechas[k] = parseFecha(r[k]);
      if (!fechas[k]) errores.push(`${donde}: falta o es inválida la ${etiqueta} ("${r[k] || ''}"). No se inventa: completala.`);
    }
    if (fechas.desde && fechas.hasta && fechas.hasta < fechas.desde) errores.push(`${donde}: el período hasta es anterior al período desde.`);
    if (fechas.fecha && fechas.vencimiento && fechas.vencimiento < fechas.fecha) errores.push(`${donde}: el vencimiento es anterior a la fecha de la factura.`);
    if (fechas.fecha) {
      const dias = Math.round((fechas.fecha - hoy) / 86400000);
      // RCEL, para servicios, acepta hasta 10 días antes o después de hoy.
      if (Math.abs(dias) > 10) errores.push(`${donde}: la fecha ${fmtFecha(fechas.fecha)} está a ${Math.abs(dias)} días de hoy; ARCA sólo acepta ±10 días para servicios.`);
    }

    const alicuota = tipoK === 'C' ? 0 : emisor.alicuota;
    const iva = tipoK === 'C' || !(neto > 0) || !alicuota ? 0 : redondear(neto * alicuota);
    const total = neto > 0 ? redondear(neto + iva) : 0;
    if (tipoK && tipoK !== 'C' && Math.abs(redondear(neto + iva) - total) > 0.001) errores.push(`${donde}: neto + IVA no da el total.`);

    facturas.push({
      fila: r.fila, clienteCuit: cuit,
      cliente: (cli && cli.razonSocial) || r.cliente || '',
      condReceptor: condR, tipo: tipoK, descripcion,
      neto: neto > 0 ? redondear(neto) : null, alicuota, iva, total,
      fecha: fechas.fecha ? fmtFecha(fechas.fecha) : r.fecha,
      desde: fechas.desde ? fmtFecha(fechas.desde) : r.desde,
      hasta: fechas.hasta ? fmtFecha(fechas.hasta) : r.hasta,
      vencimiento: fechas.vencimiento ? fmtFecha(fechas.vencimiento) : r.vencimiento,
      condVenta, observaciones: r.observaciones || '',
    });
  }

  return { emisor, facturas, errores, avisos, planilla, yaEmitidas };
}

// Clave estable de una factura (para el registro) y código del lote (para
// aprobar). Cualquier cambio en un dato que va al comprobante cambia el código.
function claveFactura(emisor, planilla, f) {
  const base = [emisor.cuit, emisor.pv, planilla ? path.basename(planilla.ruta) : '', planilla ? planilla.hoja : '',
    f.fila, f.clienteCuit, f.tipo, f.descripcion, f.neto, f.iva, f.total, f.fecha, f.desde, f.hasta, f.vencimiento, f.condVenta];
  return crypto.createHash('sha256').update(JSON.stringify(base)).digest('hex').slice(0, 16);
}
function codigoLote(emisor, facturas) {
  const base = [emisor.cuit, emisor.pv, emisor.actividad, facturas.map(f => f.clave)];
  return crypto.createHash('sha256').update(JSON.stringify(base)).digest('hex').slice(0, 8).toUpperCase();
}

function tabla(emisor, facturas) {
  const filas = facturas.map(f => [
    f.fila ?? '-', f.cliente || '(padrón)', f.clienteCuit, TIPOS[f.tipo] ? TIPOS[f.tipo].nombre : '?', f.descripcion,
    ...(emisor.condicion === 'ri' ? [pesos(f.neto || 0), pesos(f.iva)] : []), pesos(f.total),
    f.fecha, `${f.desde} al ${f.hasta}`, f.vencimiento, f.condVenta,
  ]);
  const enc = ['Fila', 'Cliente', 'CUIT', 'Tipo', 'Descripción',
    ...(emisor.condicion === 'ri' ? ['Neto', 'IVA'] : []), 'Total', 'Fecha', 'Período', 'Vto. pago', 'Cond. venta'];
  const total = redondear(facturas.reduce((a, f) => a + f.total, 0));
  const lineas = [enc, enc.map(() => '---'), ...filas].map(c => `| ${c.join(' | ')} |`);
  lineas.push(`\n**Total: ${pesos(total)}** (${facturas.length} factura${facturas.length === 1 ? '' : 's'})`);
  return { markdown: lineas.join('\n'), total };
}

// ── Registro de lo emitido ────────────────────────────────────────────────────

function registroPath(cuit) { return path.join(RAIZ_SALIDA, `.registro-${cuit}.json`); }
function leerRegistro(cuit) {
  try { return JSON.parse(fs.readFileSync(registroPath(cuit), 'utf-8')); } catch { return {}; }
}
function guardarRegistro(cuit, reg) {
  fs.mkdirSync(RAIZ_SALIDA, { recursive: true });
  const p = registroPath(cuit);
  fs.writeFileSync(p + '.tmp', JSON.stringify(reg, null, 1));
  fs.renameSync(p + '.tmp', p);
}

// ── Recorrido del asistente para UNA factura, hasta la confirmación ───────────

let paso = 0;
async function dump(page, etiqueta) {
  try {
    fs.mkdirSync(DEBUG_DIR, { recursive: true });
    const base = path.join(DEBUG_DIR, `${String(++paso).padStart(3, '0')}-${etiqueta}-${Date.now()}`);
    fs.writeFileSync(`${base}.html`, await page.content());
    await page.screenshot({ path: `${base}.png`, fullPage: true }).catch(() => {});
  } catch { /* el dump es ayuda, no puede romper la corrida */ }
}

async function completarHastaConfirmacion(rcel, alertas, emisor, f) {
  const tipo = TIPOS[f.tipo];
  const etiqueta = `f${f.fila ?? 'x'}`;
  await rcelLib.irAGenerar(rcel);

  // Paso 0: punto de venta y tipo
  const elegido = await rcelLib.elegirPvYTipo(rcel, { pv: emisor.pv, tipoRe: tipo.re, tipoNombre: tipo.nombre });
  await rcelLib.continuar(rcel, 'Datos de Emisión', { alertas, esperado: 1 });
  await dump(rcel, `${etiqueta}-paso1`);

  // Paso 1: fecha, concepto, período, vencimiento, actividad
  const con = await rcelLib.elegirSelect(rcel, 'idconcepto', /^servicios$/i);
  if (!con.ok) throw new Error(`No pude elegir el concepto "Servicios": ${con.reason}. Opciones: ${JSON.stringify(con.opts || [])}`);
  await rcel.waitForTimeout(600);
  await rcelLib.setFechaComprobante(rcel, f.fecha);

  const actis = await rcelLib.opcionesSelect(rcel, 'actiAsociadaId');
  if (actis && actis.length) {
    let re = null;
    if (emisor.actividad) re = new RegExp(escapeRe(emisor.actividad), 'i');
    else if (actis.length === 1) re = /\S/;
    else throw new Error(`El emisor tiene ${actis.length} actividades y no se indicó cuál usar (hoja Emisor o --actividad): ${actis.join(' | ')}`);
    const acti = await rcelLib.elegirSelect(rcel, 'actiAsociadaId', re);
    if (!acti.ok) throw new Error(`La actividad "${emisor.actividad}" no está entre las del emisor: ${actis.join(' | ')}`);
    rcelLib.log(`actividad: ${acti.picked}`);
  }

  const per = await rcelLib.setCampos(rcel, { fsd: f.desde, fsh: f.hasta, vencimientopago: f.vencimiento });
  if (per.faltantes.length) throw new Error(`No encontré los campos de período/vencimiento: ${per.faltantes.join(', ')}`);
  // Volver a pisar la fecha: cambiar el concepto o el período puede resetearla.
  await rcelLib.setFechaComprobante(rcel, f.fecha);

  await rcelLib.continuar(rcel, 'Datos del Receptor', { alertas, esperado: 2 });
  await dump(rcel, `${etiqueta}-paso2`);

  // Paso 2: receptor y condición de venta
  const civa = await rcelLib.elegirSelect(rcel, 'idivareceptor', COND_RECEPTOR[f.condReceptor].re);
  if (!civa.ok) throw new Error(`No pude elegir la condición IVA "${COND_RECEPTOR[f.condReceptor].nombre}": ${JSON.stringify(civa.opts || [])}`);
  await rcel.waitForTimeout(500);
  const td = await rcelLib.elegirSelect(rcel, 'idtipodocreceptor', /^cuit$/i);
  if (!td.ok) throw new Error(`No pude elegir el tipo de documento CUIT: ${JSON.stringify(td.opts || [])}`);
  await rcelLib.setCampos(rcel, { nrodocreceptor: f.clienteCuit });
  const padron = await rcelLib.esperarRazonSocial(rcel);
  if (!padron.razon) throw new Error(`ARCA no devolvió la razón social del CUIT ${f.clienteCuit} (¿CUIT inexistente o padrón caído?).`);
  rcelLib.log(`receptor: ${padron.razon}${padron.domicilio ? ' · ' + padron.domicilio : ''}`);
  f.razonSocialArca = padron.razon;
  f.domicilioArca = padron.domicilio;
  if (f.cliente && !rcelLib.norm(padron.razon).includes(rcelLib.norm(f.cliente)) && !rcelLib.norm(f.cliente).includes(rcelLib.norm(padron.razon))) {
    f.avisoRazon = `La planilla dice "${f.cliente}" y ARCA "${padron.razon}"`;
    log(`AVISO fila ${f.fila ?? '-'}: ${f.avisoRazon}`);
  }

  const cv = await rcelLib.elegirCondicionVenta(rcel, condVentaRe(f.condVenta));
  if (!cv.ok) throw new Error(`No pude tildar la condición de venta "${f.condVenta}". Opciones: ${JSON.stringify(cv.opts || [])}`);

  await rcelLib.continuar(rcel, 'Detalle de la operación', { alertas, esperado: 3 });
  await dump(rcel, `${etiqueta}-paso3`);

  // Paso 3: ítem. C: precio = total. A: precio = neto + alícuota. B: RCEL pide
  // el precio con IVA incluido; se controla el total en la confirmación.
  const precio = f.tipo === 'A' ? f.neto : f.total;
  const det = await rcelLib.setCampos(rcel, {
    detalle_descripcion1: f.descripcion,
    detalle_cantidad1: '1',
    detalle_precio1: precio.toFixed(2).replace('.', ','),
  });
  if (det.faltantes.length) throw new Error(`No encontré los campos del detalle: ${det.faltantes.join(', ')}`);
  if (f.tipo !== 'C') {
    const ali = await rcelLib.elegirSelect(rcel, 'detalle_tipo_iva1', alicuotaRe(f.alicuota));
    if (!ali.ok) throw new Error(`No pude elegir la alícuota ${f.alicuota * 100}%: ${JSON.stringify(ali.opts || [])}`);
    await rcel.waitForTimeout(600);
  }
  await dump(rcel, `${etiqueta}-paso3-lleno`);

  await rcelLib.continuar(rcel, 'Confirmación', { alertas, esperado: 4 });
  await dump(rcel, `${etiqueta}-paso4`);

  const errs = await rcelLib.erroresDeValidacion(rcel);
  if (errs.length) throw new Error(`ARCA rechazó los datos: ${errs.join(' | ')}`);

  // Control de la pantalla de confirmación contra lo aprobado.
  const texto = await rcelLib.textoPantalla(rcel);
  const faltan = [];
  if (!texto.replace(/\D/g, '').includes(f.clienteCuit)) faltan.push(`CUIT receptor ${f.clienteCuit}`);
  if (!rcelLib.contieneImporte(texto, f.total)) faltan.push(`total ${pesos(f.total)}`);
  // En la B el IVA va incluido y RCEL no siempre lo muestra aparte: alcanza con el total.
  if (f.tipo === 'A' && !rcelLib.contieneImporte(texto, f.iva)) faltan.push(`IVA ${pesos(f.iva)}`);
  if (f.tipo === 'A' && !rcelLib.contieneImporte(texto, f.neto)) faltan.push(`neto ${pesos(f.neto)}`);
  if (!texto.includes(f.fecha)) faltan.push(`fecha ${f.fecha}`);
  if (!new RegExp(tipo.nombre.replace(' ', '\\s*'), 'i').test(texto)) faltan.push(tipo.nombre);
  if (faltan.length) {
    const e = new Error(`La pantalla de confirmación no coincide con la planilla (no encuentro: ${faltan.join(', ')}). No se confirma. Ver ${DEBUG_DIR}`);
    e.code = 'confirmacion_no_coincide';
    throw e;
  }
  return { elegido, texto };
}

// Nombre final del PDF: AAAAMM - [Emisor] - [Cliente].pdf (AAAAMM = período).
function destinoPdf(dir, emisor, f, comprobante) {
  const d = parseFecha(f.desde);
  const aaaamm = `${d.getFullYear()}${pad(d.getMonth() + 1)}`;
  const base = `${aaaamm} - ${nombreArchivo(emisor.nombre || emisor.cuit)} - ${nombreArchivo(f.cliente || f.razonSocialArca || f.clienteCuit)}`;
  let p = path.join(dir, `${base}.pdf`);
  if (fs.existsSync(p)) p = path.join(dir, `${base} - ${nombreArchivo(comprobante)}.pdf`);
  return p;
}

// ── Main ──────────────────────────────────────────────────────────────────────

(async () => {
  const lote = armarLote();
  const { emisor, avisos, planilla } = lote;
  let { facturas } = lote;

  // Credenciales, igual que anticipos-sct. En --validar no se tocan, y con el
  // CUIT del emisor ya conocido se resuelven recién antes de abrir el navegador
  // (así un código mal copiado no llega a buscar la clave).
  let CUIT_LOGIN, PASSWORD, TITULAR;
  const credenciales = () => {
    if (PASSWORD) return;
    try {
      const cred = resolveCredentials({ cuit: emisor.cuit, nombre: emisor.nombreClaves, cuitLogin: args['cuit-login'] });
      ({ cuitLogin: CUIT_LOGIN, password: PASSWORD, titular: TITULAR } = cred);
      if (!emisor.cuit) emisor.cuit = cred.cuit;
      if (!emisor.nombre && TITULAR) emisor.nombre = TITULAR;
    } catch (e) {
      fail(e.message, { code: 'sin_credenciales', sugerencia: `Cargá al emisor en Claves_Organismos.xlsx o en el .env (${envFile() || 'no encontrado'}) como ARCA_<cuit>_LOGIN / ARCA_<cuit>_PASSWORD.` });
    }
  };
  if (!VALIDAR && !lote.errores.length && !emisor.cuit) credenciales();

  const registro = emisor.cuit ? leerRegistro(emisor.cuit) : {};
  for (const f of facturas) f.clave = claveFactura(emisor, planilla, f);
  const enRegistro = facturas.filter(f => registro[f.clave] && registro[f.clave].estado === 'emitida');
  const trabadas = facturas.filter(f => registro[f.clave] && registro[f.clave].estado !== 'emitida');
  for (const f of enRegistro) avisos.push(`Fila ${f.fila ?? '-'}: ya emitida (${registro[f.clave].comprobante}) según el registro; no se repite.`);
  facturas = facturas.filter(f => !enRegistro.includes(f));

  if (lote.errores.length) {
    fail('La planilla tiene errores; corregilos y volvé a correr.', { errores: lote.errores, avisos });
  }
  if (trabadas.length) {
    fail('Hay facturas que quedaron a mitad de la emisión en una corrida anterior. Verificá en RCEL → Consultas si salieron ANTES de seguir.', {
      code: 'verificar_a_mano',
      facturas: trabadas.map(f => ({ fila: f.fila, cliente: f.cliente, total: f.total, ...registro[f.clave] })),
      registro: registroPath(emisor.cuit),
    });
  }
  if (!facturas.length) {
    console.log(JSON.stringify({ ok: true, mensaje: 'No hay facturas pendientes de emitir.', yaEmitidas: lote.yaEmitidas + enRegistro.length, avisos }, null, 2));
    return;
  }

  const codigo = codigoLote(emisor, facturas);
  const resumen = tabla(emisor, facturas);
  const cabecera = {
    emisor: { nombre: emisor.nombre, cuit: emisor.cuit, condicion: emisor.condicion === 'mono' ? 'Monotributo' : 'Responsable Inscripto', puntoVenta: emisor.pv, actividad: emisor.actividad || null },
    ...(planilla ? { planilla: planilla.ruta, hoja: planilla.hoja } : {}),
    facturas: facturas.length, total: resumen.total, codigo,
  };

  if (VALIDAR) {
    console.log(JSON.stringify({ ok: true, modo: 'validar', ...cabecera, avisos, tabla: resumen.markdown, detalle: facturas,
      siguiente: `Dry-run: el mismo comando sin --validar. Emitir (con OK): --emitir=${codigo}` }, null, 2));
    return;
  }
  if (EMITIR && CODIGO_PEDIDO !== codigo) {
    fail(`El código ${CODIGO_PEDIDO} no coincide con el de las facturas pendientes (${codigo}): la planilla cambió desde que se aprobó. Volvé a mostrar el resumen con --validar.`,
      { code: 'codigo_no_coincide', ...cabecera, tabla: resumen.markdown });
  }

  const dirPdf = (emisor.carpeta && fs.existsSync(emisor.carpeta))
    ? emisor.carpeta
    : (args['salida'] || path.join(RAIZ_SALIDA, slug(emisor.nombre || emisor.cuit)));
  if (emisor.carpeta && !fs.existsSync(emisor.carpeta)) avisos.push(`La carpeta del emisor "${emisor.carpeta}" no existe; los PDF van a ${dirPdf}.`);
  fs.mkdirSync(dirPdf, { recursive: true });

  credenciales();
  const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 250 } : {}));
  const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1400, height: 950 }, acceptDownloads: true });
  const page = await context.newPage();
  const resultados = [];
  let corte = null;

  try {
    log(`${emisor.nombre || emisor.cuit} · CUIT ${emisor.cuit} · ${facturas.length} factura(s) · ${pesos(resumen.total)} · ${EMITIR ? 'EMISIÓN REAL' : 'dry-run'}`);
    log('entrando a ARCA…');
    await login(page, { cuitLogin: CUIT_LOGIN, password: PASSWORD });
    const rcel = await rcelLib.abrir(context, page, abrirServicio, { cuit: emisor.cuit, nombre: emisor.nombre });
    const alertas = rcelLib.capturarAlertas(rcel);

    for (const f of facturas) {
      const id = `fila ${f.fila ?? '-'} · ${f.cliente || f.clienteCuit} · ${pesos(f.total)}`;
      log(`── ${id}`);

      // Antes de emitir: qué comprobantes ya hay en esa fecha, para reconocer
      // después el nuevo sin confundirlo con uno anterior.
      let antes = null;
      if (EMITIR) {
        antes = await rcelLib.consultarGenerados(rcel, { pv: emisor.pv, desde: f.fecha, hasta: f.fecha });
      }

      try {
        await completarHastaConfirmacion(rcel, alertas, emisor, f);
      } catch (e) {
        // Todavía no se confirmó nada: se registra el error y se sigue con la
        // siguiente (en dry-run conviene ver todos los problemas de una).
        await dump(rcel, `f${f.fila ?? 'x'}-error`);
        log(`ERROR ${id}: ${e.message}`);
        resultados.push({ fila: f.fila, cliente: f.cliente, total: f.total, ok: false, error: e.message, code: e.code || 'error' });
        alertas.splice(0);
        continue;
      }

      if (!EMITIR) {
        const dirShots = path.join(RAIZ_SALIDA, slug(emisor.nombre || emisor.cuit), 'dry-run');
        fs.mkdirSync(dirShots, { recursive: true });
        const shot = path.join(dirShots, `DRYRUN-fila${f.fila ?? 'x'}-${f.clienteCuit}.png`);
        await rcel.screenshot({ path: shot, fullPage: true }).catch(() => {});
        resultados.push({ fila: f.fila, cliente: f.cliente, razonSocialArca: f.razonSocialArca, tipo: TIPOS[f.tipo].nombre,
          total: f.total, ok: true, dryRun: true, captura: shot, ...(f.avisoRazon ? { aviso: f.avisoRazon } : {}) });
        log(`dry-run OK: la confirmación coincide (no se confirmó)`);
        continue;
      }

      // ── Emisión real ────────────────────────────────────────────────────────
      registro[f.clave] = { estado: 'confirmando', fila: f.fila, cliente: f.cliente, clienteCuit: f.clienteCuit, total: f.total, fecha: f.fecha, desde: new Date().toISOString() };
      guardarRegistro(emisor.cuit, registro);
      log('CONFIRMANDO — esto emite un comprobante fiscal con CAE');

      // "Confirmar Datos..." abre un diálogo; la emisión la dispara el
      // "Confirmar" de ese diálogo.
      const abrio = await rcel.evaluate(() => {
        const b = Array.from(document.querySelectorAll('input[type="button"], input[type="submit"], button'))
          .find(x => /confirmar datos/i.test(x.value || x.textContent || ''));
        if (!b) return false;
        b.click();
        return true;
      });
      if (!abrio) {
        delete registro[f.clave]; guardarRegistro(emisor.cuit, registro);
        corte = 'No encontré el botón "Confirmar Datos..." en la confirmación. No se emitió; corto el lote.';
        resultados.push({ fila: f.fila, cliente: f.cliente, total: f.total, ok: false, error: corte });
        break;
      }
      const dialogo = rcel.locator('.ui-dialog:visible button, .ui-dialog-buttonpane button').filter({ hasText: /^\s*Confirmar\s*$/i }).first();
      try {
        await dialogo.waitFor({ state: 'visible', timeout: 20000 });
      } catch {
        delete registro[f.clave]; guardarRegistro(emisor.cuit, registro);
        corte = 'No apareció el diálogo "¿Confirma la operación?". No se emitió; corto el lote.';
        resultados.push({ fila: f.fila, cliente: f.cliente, total: f.total, ok: false, error: corte });
        break;
      }
      await dialogo.click();
      await rcel.waitForTimeout(8000);
      await dump(rcel, `f${f.fila ?? 'x'}-emitida`);

      // La verdad está en Consultas: el comprobante nuevo es el que no estaba antes.
      const previos = new Set(antes.map(c => c.comprobante));
      const despues = await rcelLib.consultarGenerados(rcel, { pv: emisor.pv, desde: f.fecha, hasta: f.fecha });
      const nuevos = despues.filter(c => !previos.has(c.comprobante) && soloDigitos(c.nroDoc) === f.clienteCuit);
      if (nuevos.length !== 1) {
        registro[f.clave].estado = 'verificar';
        guardarRegistro(emisor.cuit, registro);
        corte = nuevos.length === 0
          ? `Confirmé la fila ${f.fila ?? '-'} pero no aparece el comprobante nuevo en RCEL → Consultas. Verificá a mano ANTES de reintentar: un reintento a ciegas duplicaría la factura.`
          : `Aparecen ${nuevos.length} comprobantes nuevos para ${f.clienteCuit} en ${f.fecha}. Verificá a mano en RCEL → Consultas.`;
        resultados.push({ fila: f.fila, cliente: f.cliente, total: f.total, ok: false, error: corte, code: 'verificar_a_mano' });
        break;
      }
      const emitido = nuevos[0];
      const totalArca = numeroArg(String(emitido.total || '').replace(/[^\d.,]/g, ''));
      log(`emitido: ${emitido.tipo} ${emitido.comprobante} · CAE ${emitido.cae} · ${emitido.total}`);

      let pdf = null, avisoPdf = '';
      if (emitido.id) {
        pdf = await rcelLib.descargarPdf(rcel, emitido.id, destinoPdf(dirPdf, emisor, f, emitido.comprobante)).catch(() => null);
      }
      if (!pdf) avisoPdf = 'El comprobante se emitió pero no pude bajar el PDF; está en RCEL → Consultas. No reemitir.';

      registro[f.clave] = { ...registro[f.clave], estado: 'emitida', comprobante: emitido.comprobante, cae: emitido.cae, totalArca: emitido.total, pdf, hasta: new Date().toISOString() };
      guardarRegistro(emisor.cuit, registro);

      let marcada = false, avisoPlanilla = '';
      if (planilla && f.fila) {
        try {
          execFileSync(PYTHON, [path.join(__dirname, 'facturacion_planilla.py'), 'marcar', planilla.ruta, planilla.hoja, String(f.fila), emitido.comprobante],
            { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
          marcada = true;
        } catch (e) {
          avisoPlanilla = `No pude marcar la fila ${f.fila} en la planilla (¿está abierta?): anotá ${emitido.comprobante} y ¿Emitida? = Sí a mano. El registro ya la tiene como emitida.`;
          log(`AVISO: ${avisoPlanilla}`);
        }
      }

      resultados.push({
        fila: f.fila, cliente: f.cliente, razonSocialArca: f.razonSocialArca, tipo: emitido.tipo || TIPOS[f.tipo].nombre,
        comprobante: emitido.comprobante, cae: emitido.cae, total: f.total, totalArca: emitido.total, ok: true,
        pdf, planillaMarcada: marcada,
        ...(totalArca != null && Math.abs(totalArca - f.total) > 0.005 ? { aviso: `ARCA registró ${emitido.total} y la planilla dice ${pesos(f.total)}` } : {}),
        ...(avisoPdf ? { avisoPdf } : {}), ...(avisoPlanilla ? { avisoPlanilla } : {}),
      });
    }
  } catch (e) {
    await dump(page, 'error').catch(() => {});
    const que = {
      captcha: 'ARCA pidió un captcha. Corré scripts/login.js --quedate, resolvelo a mano y reintentá.',
      credenciales_invalidas: 'ARCA rechazó las credenciales. Revisá la fila del emisor; ojo con reintentar, ARCA bloquea la clave.',
    }[e.code];
    corte = que || e.message;
    resultados.push({ ok: false, error: corte, code: e.code || 'error' });
  } finally {
    await browser.close().catch(() => {});
  }

  const ok = resultados.filter(r => r.ok);
  const emitidas = ok.filter(r => !r.dryRun);
  const salida = {
    ok: !corte && ok.length === facturas.length,
    modo: EMITIR ? 'emision' : 'dry-run',
    ...cabecera,
    ...(EMITIR ? { emitidas: emitidas.length, totalEmitido: redondear(emitidas.reduce((a, r) => a + r.total, 0)) } : {}),
    ...(corte ? { corte } : {}),
    carpetaPdf: dirPdf,
    avisos,
    resultados,
    ...(!EMITIR && ok.length === facturas.length ? { siguiente: `Con el OK de Joaquin: --emitir=${codigo}` } : {}),
    tabla: resumen.markdown,
  };
  console.log(JSON.stringify(salida, null, 2));
  if (!salida.ok) process.exitCode = 1;
})();
