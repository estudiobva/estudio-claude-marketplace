#!/usr/bin/env node
/**
 * papel-trabajo-compras.js — completa las compras del mes en el papel de trabajo de una o varias sociedades:
 * borrador del Libro Compras (Portal IVA) + Zetti 5.6.4 -> servicios y NCR controlados -> hojas del papel
 * (Compras Zetti, Compras Portal IVA con el cuadro de prorrateo, Cruce compras, Resumen cruce compras y
 * Clasificacion compras) -> recalculo con Excel -> controles -> se pisa el MISMO papel de la unidad.
 *
 *   node scripts/papel-trabajo-compras.js --nombre="<sociedad>" --periodo=09/2026
 *   node scripts/papel-trabajo-compras.js --cuit=<CUIT> --no-es-ncr=503-86152
 *   node scripts/papel-trabajo-compras.js --nombres="<A>;<B>" --periodo=09/2026
 *
 * Por sociedad:
 *   1. Papel del mes (carpeta AAAAMM, "AAAAMM-papeldetrabajo-<SOC>.xlsx"): tiene que existir y tener la hoja
 *      "liquidacion" (la arma el cruce de ventas). Se copia a la carpeta de trabajo; esa copia es el respaldo.
 *   2. Conciliacion NCR: [ncr]/AAAAMM-COMPRAS-SIN-NCR-*.xlsx (conciliacion-ncr). Sin eso no se arma.
 *   3. Borrador ACTUAL del Libro Compras, exportado con leer-borrador-compras.js (solo lectura, sin importar:
 *      importar volveria a traer las NCR). Asi entran los servicios que se cargaron a mano despues.
 *   4. Zetti 5.6.4 bajado de nuevo (--no-archivar) y comparado con el de la unidad. Si la unidad no lo tiene,
 *      se archiva en la carpeta IVA como "AAAAMM - ZETTI - COMPRAS.xlsx".
 *   5. armar_papel_compras.py: corta si queda alguna NCR en el borrador, si las hojas ya existen (salvo
 *      --si-existe=reemplazar) o si aparece un tipo de comprobante que el cuadro no contempla. Servicios
 *      (AYSA, EDESUR, METROGAS) que estan en Zetti y no en el borrador se agregan a mano en naranja y se avisa.
 *   6. Excel recalcula (COM, oculto) y verificar_papel.py controla: cruce y clasificacion en 0, cuadro vs
 *      libro < $1, sin errores de Excel, hojas previas iguales.
 *   7. Si todo da, se pisa el papel de la unidad (sin copias ni backups en la unidad), controlando antes que
 *      nadie lo haya cambiado ni lo tenga abierto (~$).
 *
 * Argumentos:
 *   --nombre / --cuit, --nombres / --cuits   sociedad(es), como en los otros scripts.
 *   --periodo        MM/AAAA o AAAAMM. Default: el mes anterior.
 *   --si-existe      abortar (default) | reemplazar: si el papel ya tiene las hojas de compras, las rehace.
 *   --no-es-ncr      PV-NRO de NC de Suizo/Monroe que el estudio confirmo que NO son NCR (separadas por ;).
 *   --login-nombre   Titular de la planilla con cuya clave se entra al Portal IVA si la de la sociedad no tiene
 *                    la representada (default: BVA_ARCA_LOGIN_NOMBRE del .env).
 *   --papel=<ruta>   Papel a usar, si en la carpeta del mes hay mas de uno.
 *   --zetti-unidad   Usa el Zetti de la carpeta IVA en vez de bajarlo de nuevo.
 *   --no-archivar    Arma y controla, pero deja el resultado en la carpeta de trabajo sin tocar la unidad.
 *
 * Salida: JSON por stdout; progreso por stderr. Navegador oculto (lo manejan los scripts de descarga).
 * Nunca escribe en ARCA ni en Zetti.
 */

const fs     = require('fs');
const os     = require('os');
const path   = require('path');
const crypto = require('crypto');
const { spawnSync } = require('child_process');

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const rutas = require('../lib/rutas-bva');

loadEnv();
const args = parseArgs();
const SI_EXISTE = String(args['si-existe'] || 'abortar').toLowerCase();
const ARCHIVAR = !('no-archivar' in args);
const ZETTI_UNIDAD = 'zetti-unidad' in args;
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
if (!['abortar', 'reemplazar'].includes(SI_EXISTE)) fail('--si-existe tiene que ser abortar o reemplazar');
if (process.platform !== 'win32') fail('Hace falta Windows con Excel para recalcular el papel');

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
const MM_AAAA = `${PERIODO.slice(4)}/${PERIODO.slice(0, 4)}`;
const ANTERIOR = (() => {
  const d = new Date(Number(PERIODO.slice(0, 4)), Number(PERIODO.slice(4)) - 2, 1);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
})();

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length) fail('Falta --nombre, --cuit, --nombres o --cuits');
if (args['papel'] && PEDIDAS.length !== 1) fail('--papel es para una sola sociedad');

let prefijo = '';
const log = (m) => process.stderr.write(`[papel-compras]${prefijo} ${m}\n`);
const md5 = (f) => crypto.createHash('md5').update(fs.readFileSync(f)).digest('hex');
const sello = () => new Date().toISOString().replace(/[-:]/g, '').replace('T', '-').slice(0, 13);
const ahora = () => new Date().toLocaleString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false }).replace(',', '');

function cuitDe(pedida) {
  if (pedida.cuit) return claves.normCuit(pedida.cuit);
  const { rows } = claves.loadSync();
  const { row, candidatos } = claves.buscar(rows, { nombre: pedida.nombre });
  if (candidatos) {
    throw new Error(`"${pedida.nombre}" coincide con ${candidatos.length} titulares: ` +
      candidatos.slice(0, 8).map((r) => `${r.titular} (${r.cuit})`).join(', ') + '. Afina el nombre o pasa --cuit.');
  }
  return row.cuit;
}

// Corre otro script y devuelve el JSON que imprime (el ultimo objeto de stdout).
function correr(cmd, argv, etiqueta, env) {
  log(`${etiqueta}: ${path.basename(argv[0])} ${argv.slice(1).filter((x) => !/password/i.test(x)).map((x) => x.length > 90 ? '…' + x.slice(-60) : x).join(' ')}`);
  const p = spawnSync(cmd, argv, { encoding: 'utf-8', maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'pipe'],
                                   env: { ...process.env, PYTHONIOENCODING: 'utf-8', ...(env || {}) } });
  for (const l of String(p.stderr || '').split(/\r?\n/).filter(Boolean)) process.stderr.write(`    ${l}\n`);
  const out = String(p.stdout || '').trim();
  let json = null;
  try { json = JSON.parse(out); } catch {
    const i = out.lastIndexOf('\n{');
    try { json = JSON.parse(i >= 0 ? out.slice(i + 1) : out); } catch { /* no era JSON */ }
  }
  return { status: p.status, json, stdout: out };
}
const corte = (code, msg, extra = {}) => Object.assign(new Error(msg), { code, ...extra });

// Papel del mes: el .xlsx que diga "papel" en la carpeta AAAAMM (no los de agentes ni los marcados "NO VA" o "mal").
function papelDelMes(cuit) {
  if (args['papel']) return path.resolve(args['papel']);
  const mes = rutas.carpetaMes(path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales'), PERIODO);
  if (!fs.existsSync(mes)) throw corte('sin_papel', `No existe la carpeta del mes ${mes}`);
  const xlsx = fs.readdirSync(mes).filter((f) => /papel/i.test(f) && /\.xlsx$/i.test(f) && !/^~\$/.test(f)
    && !/agente/i.test(f) && !/no va|no usar/i.test(f));
  if (xlsx.length > 1) throw corte('papel_ambiguo', `Hay ${xlsx.length} papeles de trabajo en ${mes}: ${xlsx.join(' | ')}. Indica cual con --papel.`);
  if (!xlsx.length) throw corte('sin_papel', `No hay papel de trabajo .xlsx en ${mes}: correr antes el cruce de ventas (cruze-zetti-portaliva-ventas).`);
  return path.join(mes, xlsx[0]);
}

function abierto(papel) {
  const dir = path.dirname(papel), base = path.basename(papel);
  return fs.readdirSync(dir).filter((f) => f.startsWith('~$') && base.endsWith(f.slice(2).replace(/^.{0,2}/, '')));
}

function sinNcr(cuit) {
  const dir = rutas.carpetaNcr(cuit, PERIODO, { crear: false });
  const fs_ = fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => new RegExp(`^${PERIODO}-COMPRAS-SIN-NCR-.*\\.xlsx$`, 'i').test(f)) : [];
  if (!fs_.length) {
    throw corte('sin_conciliacion', `No esta la conciliacion NCR del periodo (${PERIODO}-COMPRAS-SIN-NCR-*.xlsx en ${dir}). ` +
      'Correr antes conciliacion-ncr y eliminar las NCR del borrador.');
  }
  const ult = fs_.map((f) => path.join(dir, f)).sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)[0];
  return { dir, archivo: ult };
}

function recalcular(entrada, salida) {
  const q = (s) => s.replace(/'/g, "''");
  const ps = [
    "$ErrorActionPreference = 'Stop'",
    '$xl = New-Object -ComObject Excel.Application',
    '$xl.Visible = $false; $xl.DisplayAlerts = $false',
    'try {',
    `  $wb = $xl.Workbooks.Open('${q(entrada)}')`,
    '  $xl.CalculateFull()',
    `  $wb.SaveAs('${q(salida)}', 51)`,
    '  $wb.Close($false)',
    '} finally { $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null }',
  ].join('\n');
  log('Excel: recalculando el papel');
  const p = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', ps], { encoding: 'utf-8', timeout: 10 * 60 * 1000 });
  if (p.status !== 0 || !fs.existsSync(salida)) {
    throw corte('excel', `Excel no pudo recalcular el papel: ${String(p.stderr || p.stdout || '').trim().slice(-400)}`);
  }
}

function unaSociedad(pedida) {
  const fila = { sociedad: pedida.nombre || pedida.cuit, cuit: null, periodo: PERIODO, ok: false };
  try {
    const cuit = cuitDe(pedida);
    fila.cuit = cuit;
    const raiz = rutas.raiz();
    const iva = rutas.carpetaIva(cuit, PERIODO, { crear: false });
    const papel = papelDelMes(cuit);
    const ncr = sinNcr(cuit);
    log(`papel: ${papel}`);
    log(`conciliacion NCR: ${path.basename(ncr.archivo)}`);

    const base = process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas');
    const trabajo = path.join(base, 'papel-compras', `${cuit}-${PERIODO}-${sello()}`);
    fs.mkdirSync(path.join(trabajo, 'arca'), { recursive: true });
    const original = path.join(trabajo, 'papel-original.xlsx');
    fs.copyFileSync(papel, original);
    const md5Original = md5(papel);
    fila.trabajo = trabajo;

    // borrador actual del Libro Compras (solo lectura)
    const b = correr(process.execPath, [path.join(__dirname, 'leer-borrador-compras.js'), `--cuit=${cuit}`, `--periodo=${MM_AAAA}`,
      `--destino=${path.join(trabajo, 'arca')}`, ...(args['login-nombre'] ? [`--login-nombre=${args['login-nombre']}`] : [])], 'Portal IVA');
    if (!b.json || !b.json.ok) throw corte((b.json && b.json.code) || 'portal', `Portal IVA: ${b.json ? b.json.error : b.stdout.slice(-300)}`);
    const fechaBorrador = ahora();
    const dj = b.json.ddjj;
    const fuente = b.json.modo === 'presentada'
      ? `Libro Compras de la DDJJ presentada (${dj.formulario} ${dj.secuencia}, presentada el ${dj.presentada_el}) exportado de ARCA el ${fechaBorrador}`
      : `borrador del Libro Compras exportado de ARCA el ${fechaBorrador}`;
    fila.libro_compras = { modo: b.json.modo, archivo: b.json.archivo, en_pantalla: b.json.total_en_pantalla, login: b.json.login,
                           fecha: fechaBorrador, ...(dj ? { ddjj: dj } : {}) };
    log(`Libro Compras: ${fuente}`);

    // Zetti 5.6.4
    const zUnidad = path.join(iva, `${PERIODO} - ZETTI - COMPRAS.xlsx`);
    let zetti = zUnidad;
    if (!(ZETTI_UNIDAD && fs.existsSync(zUnidad))) {
      const z = correr(process.execPath, [path.join(__dirname, 'zetti-iva-compras.js'), `--cuit=${cuit}`, `--periodo=${MM_AAAA}`,
        '--no-archivar', '--si-existe=reemplazar'], 'Zetti', { BVA_SALIDAS_PATH: trabajo });
      const r = z.json && (z.json.resultados ? z.json.resultados[0] : z.json);
      if (!r || !r.ok) throw corte((r && r.code) || 'zetti', `Zetti: ${r ? r.error : z.stdout.slice(-300)}`);
      zetti = r.archivo;
      if (!fs.existsSync(zUnidad) && ARCHIVAR) {
        fs.copyFileSync(zetti, zUnidad);
        log(`Zetti: no estaba en la unidad, archivado en ${zUnidad}`);
        fila.zetti_archivado = zUnidad;
      }
    }

    const anterior = (() => {
      try { return path.join(rutas.carpetaIva(cuit, ANTERIOR, { crear: false }), `${ANTERIOR} - PORTAL IVA - COMPRAS.xlsx`); } catch { return ''; }
    })();
    const maestro = path.join(raiz, '0 - ESTUDIO - 00000000000', 'Maestro de Proveedores.xlsx');
    const salida = path.join(trabajo, 'papel-armado.xlsx');
    const a = correr(PYTHON, [path.join(__dirname, 'armar_papel_compras.py'), `--papel=${original}`, `--salida=${salida}`,
      `--borrador=${b.json.archivo}`, `--zetti=${zetti}`, ...(zetti !== zUnidad && fs.existsSync(zUnidad) ? [`--zetti-unidad=${zUnidad}`] : []),
      `--sin-ncr=${ncr.archivo}`, `--ncr-dir=${ncr.dir}`, `--unidad=${raiz}`, `--maestro=${maestro}`, `--cuit=${cuit}`,
      `--periodo=${PERIODO}`, ...(anterior && fs.existsSync(anterior) ? [`--portal-anterior=${anterior}`] : []),
      `--no-es-ncr=${args['no-es-ncr'] || ''}`, `--si-existe=${SI_EXISTE}`, `--fecha-borrador=${fechaBorrador}`,
      `--fuente=${fuente}`], 'Armado');
    if (!a.json || !a.json.ok) throw corte((a.json && a.json.code) || 'armado', a.json ? a.json.error : a.stdout.slice(-400), a.json && a.json.ncr ? { ncr: a.json.ncr } : {});
    const calc = path.join(trabajo, 'papel-recalculado.xlsx');
    recalcular(salida, calc);
    const v = correr(PYTHON, [path.join(__dirname, 'verificar_papel.py'), `--original=${original}`, `--nuevo=${salida}`, `--calc=${calc}`], 'Controles');
    if (!v.json) throw corte('verificacion', `No pude verificar el papel: ${v.stdout.slice(-300)}`);
    Object.assign(fila, {
      sociedad: a.json.sociedad, portal: a.json.portal, zetti: a.json.zetti, ncr: a.json.ncr, servicios: a.json.servicios,
      cruce: a.json.cruce, maestro_pendientes: a.json.maestro.pendientes, controles: v.json.controles, hallazgos: a.json.hallazgos,
      hojas: a.json.hojas, reemplazadas: a.json.reemplazadas, recalculado: calc,
    });
    if (!v.json.ok) throw corte('verificacion', `Los controles no dan: ${v.json.fallas.join(' | ')}. El papel de la unidad NO se toco; ` +
      `el armado quedo en ${calc}.`);

    if (!ARCHIVAR) {
      Object.assign(fila, { ok: true, papel: papel, archivado: false });
      log(`--no-archivar: el papel armado quedo en ${calc}`);
      return fila;
    }
    if (md5(papel) !== md5Original) throw corte('papel_cambio', `El papel ${papel} cambio mientras se armaba: no lo piso. Volver a correr.`);
    const locks = abierto(papel);
    if (locks.length) throw corte('papel_abierto', `El papel esta abierto (${locks.join(', ')}): cerrarlo y volver a correr.`);
    fs.copyFileSync(calc, papel);
    if (md5(papel) !== md5(calc)) throw corte('copia', `La copia del papel a la unidad no quedo igual: revisar ${papel}`);
    log(`papel actualizado: ${papel}`);
    Object.assign(fila, { ok: true, papel, archivado: true, respaldo: original });
  } catch (e) {
    log(`ERROR ${e.message}`);
    Object.assign(fila, { ok: false, code: e.code || 'error', error: e.message, ...(e.ncr ? { ncr_detalle: e.ncr } : {}) });
  }
  return fila;
}

const resultados = [];
for (const [i, pedida] of PEDIDAS.entries()) {
  prefijo = PEDIDAS.length > 1 ? ` [${i + 1}/${PEDIDAS.length}]` : '';
  resultados.push(unaSociedad(pedida));
}
prefijo = '';
const ok = resultados.every((r) => r.ok);
console.log(JSON.stringify(PEDIDAS.length === 1 ? resultados[0] : {
  ok, periodo: PERIODO, procesadas: resultados.length,
  con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, code: r.code, error: r.error })),
  resultados,
}, null, 2));
process.exit(ok ? 0 : 1);
