#!/usr/bin/env node
/**
 * mis-comprobantes-emitidos.js — baja de Mis Comprobantes de ARCA los
 * comprobantes EMITIDOS de un mes, los pasa a Excel y los archiva en la carpeta
 * IVA del periodo, al lado de los libros del Portal IVA y de Zetti.
 *
 *   node scripts/mis-comprobantes-emitidos.js --nombre="<sociedad>" --periodo=09/2026
 *   node scripts/mis-comprobantes-emitidos.js --cuit=<CUIT>
 *   node scripts/mis-comprobantes-emitidos.js --nombres="<A>;<B>" --periodo=09/2026
 *
 * Argumentos:
 *   --nombre / --cuit   Una sociedad (titular en Claves_Organismos.xlsx, o CUIT).
 *   --nombres / --cuits Varias, separadas por ";".
 *   --periodo           MM/AAAA o AAAAMM. Default: el mes anterior.
 *   --si-existe         abortar (default) | saltear | reemplazar, si ya esta el .xlsx.
 *   --no-archivar       Deja el .xlsx en ~/Documents/BVA-salidas/mis-comprobantes ($BVA_SALIDAS_PATH) y no toca el Drive.
 *   --ver               Navegador visible (solo si se pide).
 *
 * Archiva [IVA del periodo]/AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx, hoja
 * "Mis Comprobantes Emitidos": el CSV de ARCA completo (mismas columnas, mismo
 * orden) con tipos reales. El ZIP y el CSV se borran.
 *
 * Mis Comprobantes tiene la informacion hasta el dia de ayer: un mes bajado
 * antes de que termine puede estar incompleto.
 *
 * Salida: JSON por stdout; progreso por stderr. Solo lectura en ARCA.
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const { parseArgs, fail } = require('../lib/args');
if ('ver' in parseArgs()) process.env.PLAYWRIGHT_HEADLESS = 'false';
const { chromium } = require('playwright');
const { launchOptions } = require('../lib/launch');
const { resolveCredentials, envFile } = require('../lib/env');
const { login } = require('../lib/arca-login');
const rutas = require('../lib/rutas-bva');
const mcl   = require('../lib/mis-comprobantes');

const args = parseArgs();
const SI_EXISTE = String(args['si-existe'] || 'abortar').toLowerCase();
const ARCHIVAR = !('no-archivar' in args);
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
if (!['abortar', 'saltear', 'reemplazar'].includes(SI_EXISTE)) fail('--si-existe tiene que ser abortar, saltear o reemplazar');

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
const NOMBRE = `${PERIODO} - MIS COMPROBANTES - EMITIDOS.xlsx`;

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length) fail('Falta --nombre, --cuit, --nombres o --cuits');

let prefijo = '';
const log = (m) => process.stderr.write(`[mis-comprobantes]${prefijo} ${m}\n`);

function descomprimir(zip, dir) {
  if (process.platform === 'win32') {
    execFileSync('powershell', ['-NoProfile', '-Command',
      `Expand-Archive -LiteralPath '${zip.replace(/'/g, "''")}' -DestinationPath '${dir.replace(/'/g, "''")}' -Force`]);
  } else {
    execFileSync('unzip', ['-o', '-q', zip, '-d', dir]);
  }
  const csv = fs.readdirSync(dir).filter((f) => /\.csv$/i.test(f));
  if (csv.length !== 1) throw new Error(`El ZIP de Mis Comprobantes trae ${csv.length} CSV (esperaba 1).`);
  return path.join(dir, csv[0]);
}

async function unaSociedad(pedida) {
  let cred;
  try {
    cred = resolveCredentials({ cuit: pedida.cuit, nombre: pedida.nombre, cuitLogin: args['cuit-login'], password: args['password'] });
  } catch (e) {
    return { sociedad: pedida.nombre || pedida.cuit, ok: false, error: e.message,
      sugerencia: `Revisa el .env (${envFile() || 'no encontrado'}) o el Excel de claves.` };
  }
  const { cuit, cuitLogin, password, titular } = cred;
  const fila = { sociedad: titular || pedida.nombre || cuit, cuit, periodo: PERIODO, ok: false };

  let destinoDir;
  try {
    destinoDir = ARCHIVAR ? rutas.carpetaIva(cuit, PERIODO) : path.join(process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'mis-comprobantes');
    fs.mkdirSync(destinoDir, { recursive: true });
  } catch (e) { return { ...fila, error: e.message, code: 'destino_no_resuelto' }; }
  const destino = path.join(destinoDir, NOMBRE);
  if (fs.existsSync(destino)) {
    if (SI_EXISTE === 'saltear') return { ...fila, ok: true, salteado: true, archivo: destino };
    if (SI_EXISTE === 'abortar') return { ...fila, code: 'ya_existe', archivo: destino,
      error: `Ya esta ${destino}. Usa --si-existe=reemplazar para bajarlo de nuevo.` };
  }

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'bva-mc-'));
  const browser = await chromium.launch(launchOptions());
  const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
  const page = await context.newPage();
  try {
    log(`${fila.sociedad} · CUIT ${cuit} · emitidos ${PERIODO}`);
    log('entrando a ARCA…');
    await login(page, { cuitLogin, password });
    const mc = await mcl.abrir(context, page, cuit);
    const rango = mcl.rangoMes(PERIODO);
    const r = await mcl.descargar(mc, { tipo: 'E', ...rango, destinoZip: path.join(tmp, 'mc.zip') });
    fila.consulta = r.idConsulta;
    fila.comprobantes = r.cantidad;
    if (!r.zip) {
      return Object.assign(fila, { ok: true, archivo: null, aviso: 'Mis Comprobantes no tiene comprobantes emitidos en el periodo: no se archiva nada.' });
    }

    const csv = descomprimir(r.zip, path.join(tmp, 'csv'));
    const xlsxTmp = path.join(tmp, NOMBRE);
    const conv = JSON.parse(execFileSync(PYTHON, [path.join(__dirname, 'mc_a_excel.py'), csv, xlsxTmp, '--tipo=emitidos'], { encoding: 'utf-8' }));
    if (conv.filas !== r.cantidad) {
      throw Object.assign(new Error(`El CSV trae ${conv.filas} filas y la consulta decia ${r.cantidad}: no archivo.`), { code: 'control' });
    }
    fs.copyFileSync(xlsxTmp, destino);
    log(`archivado ${destino} (${conv.filas} comprobantes)`);
    Object.assign(fila, { ok: true, archivo: destino, filas: conv.filas, hoja: conv.hoja });
    const ultimoDia = new Date(Number(PERIODO.slice(0, 4)), Number(PERIODO.slice(4)), 1);
    if (new Date() < ultimoDia) fila.aviso = 'El periodo no termino: Mis Comprobantes tiene datos hasta ayer, el archivo puede estar incompleto.';
  } catch (e) {
    fila.code = e.code || 'error';
    fila.error = e.code === 'captcha'
      ? `ARCA pidio un captcha. Corre \`node scripts/login.js --cuit=${cuit} --quedate\`, resolvelo a mano, y reintenta.`
      : e.code === 'credenciales_invalidas'
        ? `ARCA rechazo las credenciales del CUIT ${cuit}. Verifica la fila en Claves_Organismos.xlsx; ojo con reintentar, ARCA bloquea la clave.`
        : e.message;
    log(`ERROR ${fila.error}`);
  } finally {
    await browser.close().catch(() => {});
    fs.rmSync(tmp, { recursive: true, force: true });
  }
  return fila;
}

(async () => {
  const resultados = [];
  for (const [i, pedida] of PEDIDAS.entries()) {
    prefijo = PEDIDAS.length > 1 ? ` [${i + 1}/${PEDIDAS.length}]` : '';
    const r = await unaSociedad(pedida);
    resultados.push(r);
    if (['captcha', 'credenciales_invalidas'].includes(r.code)) {
      log('corto el lote: ARCA pidio captcha o rechazo la clave');
      break;
    }
  }
  prefijo = '';
  const ok = resultados.length === PEDIDAS.length && resultados.every((r) => r.ok);
  console.log(JSON.stringify(PEDIDAS.length === 1 ? resultados[0] : {
    ok, periodo: PERIODO, procesadas: resultados.length,
    con_error: resultados.filter((r) => !r.ok).map((r) => ({ sociedad: r.sociedad, cuit: r.cuit, error: r.error })),
    resultados,
  }, null, 2));
  process.exit(ok ? 0 : 1);
})();
