#!/usr/bin/env node
/**
 * ddjj-sshh-quilmes.js — DDJJ mensual de la Tasa de Seguridad e Higiene
 * (Regimen General) de ARQui (Municipio de Quilmes): controla lo presentado
 * contra el papel de trabajo y la DJ de anticipo IIBB de ARBA (actividad
 * 477320) y archiva la boleta en la carpeta TSH del mes.
 *
 *   node scripts/ddjj-sshh-quilmes.js --nombre="<sociedad>" [--periodo=AAAAMM]
 *   node scripts/ddjj-sshh-quilmes.js --cuit=<CUIT> [--periodo=AAAAMM]
 *   node scripts/ddjj-sshh-quilmes.js --nombres="<A>;<B>" [--periodo=AAAAMM]
 *   node scripts/ddjj-sshh-quilmes.js --todas [--periodo=AAAAMM]
 *
 * Argumentos:
 *   --nombre / --cuit   Una sociedad (titular en Claves_Organismos.xlsx, o CUIT).
 *   --nombres / --cuits Varias, separadas por ";".
 *   --todas             Todas las cuentas de SSHH vinculadas al usuario de ARQui;
 *                       la sociedad sale del CUIT de sus DDJJ. Agrega la hoja
 *                       "Control planilla" (ARQui vs Tasas Municipales CLAVES).
 *   --periodo           AAAAMM o MM/AAAA. Default: el mes anterior.
 *   --no-archivar       No copia la boleta a la unidad (queda en la salida).
 *   --reemplazar        Pisa la boleta del mes si ya esta en la unidad.
 *   --ver               Navegador visible (solo si se pide).
 *
 * De donde sale cada dato:
 *   - Cuenta: bloque SEGURIDAD E HIGIENE de la hoja QUILMES de "Tasas
 *     Municipales CLAVES.xlsx", por CUIT.
 *   - Papel: [Sociedad]/01-Impuestos Mensuales/AAAAMM/AAAAMM-papeldetrabajo-*.xlsx,
 *     hoja liquidacion, Neto Gravado de "Ventas Netas (s/ Notas de Credito)".
 *     Se declara la parte entera, sin redondear.
 *   - ARBA: la DJ de anticipo IIBB del mes (cualquier *iibb*dj*.pdf de la
 *     carpeta del mes que no sea el acuse), linea 477320.
 *   - ARQui: Consultar (DDJJ del anio), la Declaracion y el Recibo de Pago.
 *
 * Login: usuario y clave del grupo MUNICIPAL (Quilmes) de Claves_Organismos.xlsx,
 * o ARQUI_USUARIO / ARQUI_CLAVE en el .env.
 *
 * Archiva el Recibo de Pago en [Sociedad]/01-Impuestos Mensuales/AAAAMM/TSH/AAAAMM-SEH-<soc>.pdf
 * (<soc> = el que ya usa el equipo para esa sociedad). Copia de todo + resumen
 * Excel en ~/Documents/BVA-salidas/ddjj-sshh-quilmes/AAAA-MM-DD/ ($BVA_SALIDAS_PATH).
 *
 * Salida: JSON por stdout; progreso por stderr. Solo consulta y descarga: no
 * presenta DDJJ ni paga.
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const { credencialesArqui, accesosQuilmes } = require('../lib/arqui-claves');
const tasas = require('../lib/tasas-claves');
const arqui = require('../lib/arqui');
const sshh = require('../lib/arqui-sshh');
const rutas = require('../lib/rutas-bva');

const args = parseArgs();
const VER = 'ver' in args;
const ARCHIVAR = !('no-archivar' in args);
const REEMPLAZAR = 'reemplazar' in args;
const TODAS = 'todas' in args;
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
// ARQui bloquea (403) si recibe muchos pedidos seguidos: pausa entre descargas
// y, ante un bloqueo, una espera antes del unico reintento.
const PAUSA_MS = Number(process.env.ARQUI_PAUSA_MS || 2500);
const ESPERA_BLOQUEO_MS = Number(process.env.ARQUI_ESPERA_BLOQUEO_MS || 300000);
// Actividad que se declara en SSHH (las de medicamentos no van).
const ACTIVIDAD = '477320';

if (VER && process.env.PLAYWRIGHT_HEADLESS == null) process.env.PLAYWRIGHT_HEADLESS = 'false';
loadEnv();

const lista = (v) => String(v || '').split(';').map((s) => s.trim()).filter(Boolean);
const PEDIDAS = [
  ...lista(args['nombres']).map((nombre) => ({ nombre })),
  ...lista(args['cuits']).map((cuit) => ({ cuit })),
];
if (args['nombre'] || args['cuit']) PEDIDAS.push({ nombre: args['nombre'], cuit: args['cuit'] });
if (!PEDIDAS.length && !TODAS) fail('Falta --nombre, --cuit, --nombres, --cuits o --todas');
if (TODAS && PEDIDAS.length) fail('--todas no se combina con --nombre/--cuit');

const hoy = new Date();
function periodoDe(v) {
  if (!v) {
    const d = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
    return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
  }
  let m = String(v).trim().match(/^(\d{4})-?(\d{2})$/);
  if (m) return `${m[1]}${m[2]}`;
  m = String(v).trim().match(/^(\d{1,2})[/-](\d{4})$/);
  if (m) return `${m[2]}${m[1].padStart(2, '0')}`;
  fail(`--periodo "${v}" no es AAAAMM ni MM/AAAA`);
}
const PERIODO = periodoDe(args['periodo']);
const ANIO = PERIODO.slice(0, 4);
const MES = Number(PERIODO.slice(4));

const log = (m) => process.stderr.write(`[ddjj-sshh-quilmes] ${m}\n`);

const FECHA = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}-${String(hoy.getDate()).padStart(2, '0')}`;
const SALIDA = path.join(process.env.BVA_SALIDAS_PATH || path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'ddjj-sshh-quilmes', FECHA);
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'ddjj-sshh-'));

function py(...a) {
  const out = execFileSync(PYTHON, [path.join(__dirname, 'sshh.py'), ...a], { encoding: 'utf-8', maxBuffer: 16 * 1024 * 1024 });
  return JSON.parse(out.trim().split(/\r?\n/).pop());
}

function cuitDe(pedida) {
  if (pedida.cuit) return claves.normCuit(pedida.cuit);
  const { rows } = claves.loadSync();
  const { row, candidatos } = claves.buscar(rows, { nombre: pedida.nombre });
  if (candidatos) {
    throw new Error(`"${pedida.nombre}" coincide con ${candidatos.length} titulares: ` +
      candidatos.slice(0, 8).map((r) => `${r.titular} (${r.cuit})`).join(', ') + '. Afina el nombre o pasa --cuit.');
  }
  if (!row.cuit) throw new Error(`"${row.titular}" no tiene CUIT en el archivo de claves.`);
  return row.cuit;
}

function nombreDe(cuit, fallback) {
  try { return rutas.nombreSociedad(cuit); } catch { return fallback || cuit; }
}

const dormir = (ms) => new Promise((r) => setTimeout(r, ms));
const redondear = (n) => Math.round(n * 100) / 100;

// ── Archivos de la unidad ─────────────────────────────────────────────────────

// Archivos (hasta 2 niveles) de la carpeta del mes, suelta o dentro del anio.
function archivosDelMes(cuit) {
  const mensuales = path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales');
  const dirs = [path.join(mensuales, PERIODO), path.join(mensuales, ANIO, PERIODO)].filter((d) => fs.existsSync(d));
  const out = [];
  const recorrer = (d, nivel) => {
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, e.name);
      if (e.isDirectory() && nivel < 2) recorrer(p, nivel + 1);
      else if (e.isFile() && !e.name.startsWith('~$')) out.push(p);
    }
  };
  dirs.forEach((d) => recorrer(d, 0));
  return out;
}

const mmaaaa = `${PERIODO.slice(4)}${ANIO}`;
const delPeriodo = (n) => n.startsWith(PERIODO) || n.startsWith(mmaaaa);

function leerPapel(archivos) {
  const papeles = archivos
    .filter((f) => /\.xlsx$/i.test(f) && /papel\s*de\s*trabajo|papeldetrabajo/i.test(path.basename(f)) && delPeriodo(path.basename(f)))
    .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
  if (!papeles.length) return { error: 'no esta el papel de trabajo del mes' };
  return py('papel', ...papeles);
}

function leerArba(archivos) {
  const djs = archivos.filter((f) => {
    const n = path.basename(f);
    return /\.pdf$/i.test(n) && /iibb/i.test(n) && /(^|[-_ ])(dj|ddjj)([-_ .]|$)/i.test(n) && !/acuse/i.test(n);
  });
  if (!djs.length) return { error: 'la DJ de IIBB del mes no esta en la unidad (si ya se presento, bajarla con constancias-anticipo-arba)' };
  const errores = [];
  for (const f of djs) {
    const r = py('arba', f);
    if (r.error) { errores.push(`${path.basename(f)}: ${r.error}`); continue; }
    if (r.periodo && r.periodo !== PERIODO) { errores.push(`${path.basename(f)} es del periodo ${r.periodo}`); continue; }
    return r;
  }
  return { error: errores.join('; ') };
}

// Token que usa el equipo en "AAAAMM-SEH-<token>.pdf" para la sociedad (el del
// mes mas reciente); si no hay ninguno, el nombre de la carpeta sin espacios.
function tokenSeh(cuit) {
  const mensuales = path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales');
  let mejor = null;
  const recorrer = (d, nivel) => {
    let es = [];
    try { es = fs.readdirSync(d, { withFileTypes: true }); } catch { return; }
    for (const e of es) {
      if (e.isDirectory() && nivel < 3) recorrer(path.join(d, e.name), nivel + 1);
      const m = e.isFile() && e.name.match(/^(\d{6})-seh-(?:boleta-)?(.+)\.pdf$/i);
      if (m && /^20\d{4}$/.test(m[1]) && (!mejor || m[1] > mejor.periodo)) mejor = { periodo: m[1], token: m[2].trim() };
    }
  };
  recorrer(mensuales, 0);
  if (mejor) return mejor.token;
  return nombreDe(cuit).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]/g, '').toLowerCase();
}

function archivarBoleta(fila, pdf, archivos) {
  const nombre = `${PERIODO}-SEH-${tokenSeh(fila.cuit)}.pdf`;
  fs.mkdirSync(SALIDA, { recursive: true });
  fs.copyFileSync(pdf, path.join(SALIDA, nombre));
  fila.archivo = path.join(SALIDA, nombre);
  if (!ARCHIVAR) { fila.boleta = 'sin archivar'; return; }
  const previa = archivos.find((f) => /\.pdf$/i.test(f) && new RegExp(`^(${PERIODO}|${mmaaaa})-seh`, 'i').test(path.basename(f)));
  if (previa && !REEMPLAZAR) {
    fila.boleta = 'ya estaba';
    fila.archivo = previa;
    return;
  }
  const destino = previa || path.join(rutas.carpetaTsh(fila.cuit, PERIODO), nombre);
  fs.copyFileSync(pdf, destino);
  fila.boleta = previa ? 'reemplazada' : 'archivada';
  fila.archivo = destino;
}

// ── Sesion de ARQui (un usuario a la vez, como en recibo-pyp-quilmes) ─────────

let browser = null;
const sesiones = new Map(); // usuario -> { cred, page, ctx, logueado }

async function sesion(cred) {
  if (!browser) browser = await chromium.launch(launchOptions());
  let s = sesiones.get(cred.usuario);
  if (!s) {
    for (const [u, otra] of sesiones) {
      await otra.ctx.close().catch(() => {});
      sesiones.delete(u);
    }
    const ctx = await browser.newContext({ acceptDownloads: true });
    s = { cred, ctx, page: await ctx.newPage(), logueado: false };
    s.page.on('dialog', (d) => d.dismiss().catch(() => {}));
    await arqui.prepararPagina(s.page);
    sesiones.set(cred.usuario, s);
  }
  if (!s.logueado) {
    log(`login en ARQui (${cred.origen})...`);
    await arqui.login(s.page, cred);
    s.logueado = true;
  }
  return s;
}

let bloqueado = null;
async function conSesion(cred, fn) {
  if (bloqueado) throw bloqueado;
  const intento = async () => fn((await sesion(cred)).page);
  try {
    return await intento();
  } catch (e) {
    if (e.code === 'sesion') {
      log('la sesion vencio: vuelvo a loguear');
      sesiones.get(cred.usuario).logueado = false;
      return intento();
    }
    if (e.code !== 'bloqueo') throw e;
    log(`ARQui bloqueo los pedidos (403): espero ${Math.round(ESPERA_BLOQUEO_MS / 1000)} s y reintento una vez...`);
    await dormir(ESPERA_BLOQUEO_MS);
    for (const s of sesiones.values()) s.logueado = false;
    try {
      return await intento();
    } catch (e2) {
      if (e2.code === 'bloqueo') {
        bloqueado = Object.assign(new Error(`${e2.message} No se siguio con el resto: volver a correr en un rato.`), { code: 'bloqueo' });
        throw bloqueado;
      }
      throw e2;
    }
  }
}

// Toda consulta o descarga pasa por aca: una pausa entre pedidos para no
// disparar el 403 (una corrida de --todas son mas de 80 pedidos).
let pedidos = 0;
async function pedir(cred, fn) {
  if (pedidos++ > 0) await dormir(PAUSA_MS);
  return conSesion(cred, fn);
}

async function bajar(cred, ddjj, cual) {
  const pdf = path.join(TMP, `${ddjj.id}-${cual}.pdf`);
  await pedir(cred, (page) => sshh.descargar(page, ddjj, cual, pdf));
  return pdf;
}

// DDJJ del anio; si no hay ninguna, se mira el anio anterior solo para saber
// de quien es la cuenta (CUIT).
async function ddjjYCuit(cred, cuenta) {
  const ddjjs = await pedir(cred, (page) => sshh.ddjjDelAnio(page, cuenta, ANIO));
  let cuit = (ddjjs.find((d) => d.cuit) || {}).cuit || null;
  let ultimaAntes = null;
  if (!ddjjs.length) {
    const antes = await pedir(cred, (page) => sshh.ddjjDelAnio(page, cuenta, Number(ANIO) - 1));
    cuit = (antes.find((d) => d.cuit) || {}).cuit || null;
    const meses = antes.filter((d) => d.origen !== 'Anual' && d.mes).map((d) => d.mes);
    if (meses.length) ultimaAntes = `${Math.max(...meses)}/${Number(ANIO) - 1}`;
  }
  return { ddjjs, cuit, ultimaAntes };
}

// ── Control de una cuenta ─────────────────────────────────────────────────────

const agregarObs = (fila, m) => { fila.observacion = [fila.observacion, m].filter(Boolean).join(' '); };

async function controlar(fila, cred, ddjjs) {
  // Datos de la unidad: papel y DJ de ARBA.
  let archivos = [];
  try { archivos = archivosDelMes(fila.cuit); } catch (e) { agregarObs(fila, `Unidad: ${e.message}`); }
  const papel = leerPapel(archivos);
  if (papel.error) agregarObs(fila, `Papel: ${papel.error}.`);
  else { fila.papel = redondear(papel.neto_gravado); fila.papel_entero = Math.trunc(papel.neto_gravado); }
  const arba = leerArba(archivos);
  if (arba.error) agregarObs(fila, `ARBA: ${arba.error}.`);
  else if (arba.actividades[ACTIVIDAD] == null) agregarObs(fila, `ARBA: la DJ no tiene la actividad ${ACTIVIDAD}.`);
  else fila.arba = arba.actividades[ACTIVIDAD];

  // DDJJ del periodo en ARQui (si esta repetida o rectificada, la ultima).
  const delMes = ddjjs.filter((d) => d.origen !== 'Anual' && d.mes === MES);
  if (!delMes.length) {
    fila.estado = 'sin DDJJ';
    const monto = fila.papel_entero != null ? fila.papel_entero : (fila.arba != null ? Math.trunc(fila.arba) : null);
    agregarObs(fila, monto != null
      ? `No hay DDJJ de ${MES}/${ANIO} en ARQui. A declarar: ${ACTIVIDAD} = ${monto} (tasa estimada ${redondear(monto * 0.013)} al 1,3%).`
      : `No hay DDJJ de ${MES}/${ANIO} en ARQui y no hay monto para declarar (falta el papel).`);
    return;
  }
  const ultima = delMes.slice().sort((a, b) => a.nro - b.nro || (a.id || 0) - (b.id || 0)).pop();
  Object.assign(fila, { fecha: ultima.fecha, caracter: `${ultima.caracter}${ultima.nro > 1 ? ` ${ultima.nro}` : ''}` });
  if (new Set(delMes.map((d) => `${d.caracter}${d.nro}`)).size > 1) agregarObs(fila, `ARQui tiene ${delMes.length} DDJJ del mes: se toma la ultima (${fila.caracter}).`);

  const dj = py('ddjj', await bajar(cred, ultima, 'ddjj'));
  if (dj.error) throw Object.assign(new Error(`Declaracion: ${dj.error}`), { code: 'pdf' });
  if (dj.cuit && dj.cuit !== fila.cuit) throw Object.assign(new Error(`La Declaracion es del CUIT ${dj.cuit}, no del ${fila.cuit}.`), { code: 'control' });
  if (dj.periodo && dj.periodo !== PERIODO) throw Object.assign(new Error(`La Declaracion es del periodo ${dj.periodo}, no del ${PERIODO}.`), { code: 'control' });
  fila.cuenta = dj.cuenta || fila.cuenta;
  const act = dj.actividades.filter((a) => a.codigo === ACTIVIDAD);
  const otras = dj.actividades.filter((a) => a.codigo !== ACTIVIDAD && a.ingreso);
  fila.declarado = act.reduce((s, a) => s + a.ingreso, 0);
  fila.tasa = dj.total != null ? dj.total : redondear(dj.actividades.reduce((s, a) => s + a.tasa, 0));
  if (otras.length) fila.otras = otras.map((a) => `${a.codigo}: ${a.ingreso}`).join('; ');
  Object.assign(fila, { vencimiento: dj.vencimiento, recibo: dj.recibo });
  fs.mkdirSync(SALIDA, { recursive: true });
  fs.copyFileSync(dj.archivo, path.join(SALIDA, `${PERIODO}-SSHH-DDJJ-${fila.cuenta || fila.cuit}.pdf`));

  // Boleta (Recibo de Pago) -> TSH.
  try {
    const rec = py('recibo', await bajar(cred, ultima, 'recibo'));
    if (rec.error) throw new Error(rec.error);
    if (rec.cuenta && dj.cuenta && rec.cuenta !== dj.cuenta) throw new Error(`el recibo es de la cuenta ${rec.cuenta}`);
    if (dj.total != null && Math.abs(rec.total - dj.total) > 0.01) agregarObs(fila, `El recibo suma ${rec.total} y la DDJJ ${dj.total}.`);
    archivarBoleta(fila, rec.archivo, archivos);
  } catch (e) {
    if (e.code === 'bloqueo') throw e;
    fila.boleta = 'error';
    agregarObs(fila, `Boleta: ${e.message}.`);
  }

  // Comparacion. Menos de $1 no cuenta: hay quien declara con centavos o
  // redondea en vez de truncar (el papel tiene centavos).
  const difs = [];
  if (fila.papel_entero != null) {
    fila.dif_papel = redondear(fila.declarado - fila.papel_entero);
    if (Math.abs(fila.dif_papel) >= 1) difs.push(`papel ${fila.papel_entero}`);
  }
  if (fila.arba != null) {
    fila.dif_arba = redondear(fila.declarado - fila.arba);
    if (Math.abs(fila.dif_arba) >= 1) difs.push(`ARBA ${fila.arba}`);
  }
  if (!difs.length && ((fila.dif_papel || 0) || (fila.dif_arba || 0))) agregarObs(fila, 'Diferencia de centavos por redondeo: no cuenta.');
  if (otras.length) difs.push(`otras actividades declaradas (${fila.otras})`);
  if (difs.length) {
    fila.estado = 'diferencia';
    agregarObs(fila, `Declarado ${fila.declarado} vs ${difs.join(', ')}. No se toca: rectificar lo decide la persona.`);
  } else {
    fila.estado = (fila.papel_entero == null && fila.arba == null) ? 'sin control' : 'OK';
  }
}

// ── Modos ─────────────────────────────────────────────────────────────────────

function credGeneral() {
  if ((process.env.ARQUI_USUARIO || '').trim() && process.env.ARQUI_CLAVE) return credencialesArqui();
  const [g] = accesosQuilmes();
  if (!g) throw new Error('No hay ningun usuario de ARQui en Claves_Organismos.xlsx (grupo MUNICIPAL con municipio Quilmes).');
  return { ...g, origen: 'archivo (general de Quilmes)' };
}

const errorFila = (fila, e) => Object.assign(fila, { estado: 'error', codigo: e.code || 'error', observacion: [fila.observacion, e.message].filter(Boolean).join(' ') });

async function porSociedad(planilla) {
  const out = [];
  for (const pedida of PEDIDAS) {
    const fila = { sociedad: pedida.nombre || pedida.cuit, periodo: `${MES}/${ANIO}`, estado: 'error' };
    out.push(fila);
    try {
      fila.cuit = cuitDe(pedida);
      fila.sociedad = nombreDe(fila.cuit, pedida.nombre);
      const enPlanilla = planilla.filas.filter((f) => f.cuit === fila.cuit);
      const buena = enPlanilla.find((f) => !f.problema);
      if (!buena) {
        fila.estado = 'sin cuenta';
        fila.observacion = enPlanilla.length
          ? `En la planilla de tasas la fila ${enPlanilla[0].fila} tiene ${enPlanilla[0].problema}.`
          : 'No esta en el bloque SEGURIDAD E HIGIENE de la planilla de tasas (hoja QUILMES).';
        continue;
      }
      fila.cuenta = buena.cuenta;
      const cred = credencialesArqui({ cuit: fila.cuit });
      log(`${fila.sociedad}: cuenta ${fila.cuenta}...`);
      const disponibles = await pedir(cred, (page) => sshh.cuentas(page));
      if (!disponibles.includes(fila.cuenta)) {
        throw Object.assign(new Error(`La cuenta ${fila.cuenta} no esta vinculada al usuario de ARQui (Seguridad e Higiene). Vincularla con "Adherir cuenta" lo decide la persona.`), { code: 'no_vinculada' });
      }
      const ddjjs = await pedir(cred, (page) => sshh.ddjjDelAnio(page, fila.cuenta, ANIO));
      await controlar(fila, cred, ddjjs);
      log(`  ${fila.estado}${fila.declarado != null ? `: declarado ${fila.declarado}` : ''}`);
    } catch (e) {
      errorFila(fila, e);
      log(`  error: ${e.message}`);
    }
  }
  return { filas: out };
}

async function todas(planilla) {
  const cred = credGeneral();
  const porCuenta = new Map(planilla.filas.filter((f) => !f.problema).map((f) => [f.cuenta, f]));
  const cuentas = await pedir(cred, (page) => sshh.cuentas(page));
  log(`${cuentas.length} cuentas de SSHH en ARQui; ${porCuenta.size} en la planilla`);
  const out = [], control = [];
  for (const cuenta of cuentas) {
    const p = porCuenta.get(cuenta);
    const fila = { sociedad: '', cuenta, periodo: `${MES}/${ANIO}`, estado: 'error' };
    const ctl = { cuenta, nombre_planilla: p && p.nombre, cuit_planilla: p && p.cuit };
    try {
      const { ddjjs, cuit: cuitArqui, ultimaAntes } = await ddjjYCuit(cred, cuenta);
      ctl.cuit_arqui = cuitArqui;
      fila.cuit = cuitArqui || (p && p.cuit) || null;
      if (!fila.cuit) {
        fila.estado = 'sin CUIT';
        fila.observacion = `La cuenta no tiene DDJJ en ${Number(ANIO) - 1} ni en ${ANIO} y no esta en la planilla: no se sabe de que sociedad es.`;
        ctl.resultado = `Sin DDJJ en ${Number(ANIO) - 1}-${ANIO} y falta en la planilla`;
        out.push(fila); control.push(ctl);
        continue;
      }
      fila.sociedad = nombreDe(fila.cuit, p && p.nombre);
      ctl.sociedad = fila.sociedad;
      ctl.resultado = !p ? 'Falta en la planilla (bloque SEGURIDAD E HIGIENE)'
        : (cuitArqui && p.cuit !== cuitArqui) ? `CUIT distinto: la planilla dice ${p.cuit}, ARQui ${cuitArqui}` : 'OK';
      log(`${cuenta} ${fila.sociedad}...`);
      if (ultimaAntes) agregarObs(fila, `La cuenta no tiene ninguna DDJJ en ${ANIO}: la ultima es de ${ultimaAntes} (¿dada de baja?).`);
      await controlar(fila, cred, ddjjs);
      log(`  ${fila.estado}${fila.declarado != null ? `: declarado ${fila.declarado}` : ''}`);
    } catch (e) {
      errorFila(fila, e);
      ctl.resultado = ctl.resultado || `No se pudo consultar: ${e.message}`;
      log(`  error: ${e.message}`);
      if (e.code === 'bloqueo') { out.push(fila); control.push(ctl); break; }
    }
    out.push(fila); control.push(ctl);
  }
  for (const f of planilla.filas) {
    if (f.problema) control.push({ cuenta: f.cuenta, nombre_planilla: f.nombre, cuit_planilla: f.cuit, resultado: `Fila ${f.fila} de la planilla con ${f.problema}` });
    else if (!cuentas.includes(f.cuenta)) control.push({ cuenta: f.cuenta, nombre_planilla: f.nombre, cuit_planilla: f.cuit, resultado: 'En la planilla pero no vinculada al usuario de ARQui' });
  }
  return { filas: out, control };
}

// ── Main ──────────────────────────────────────────────────────────────────────

(async () => {
  let res;
  try {
    const planilla = tasas.cuentasSshh();
    log(`periodo ${MES}/${ANIO}; planilla de tasas: ${planilla.filas.length} filas en SEGURIDAD E HIGIENE`);
    res = TODAS ? await todas(planilla) : await porSociedad(planilla);
  } catch (e) {
    if (browser) await browser.close().catch(() => {});
    fail(e.message, e.code ? { codigo: e.code } : {});
  }
  if (browser) await browser.close().catch(() => {});

  const cuenta = (est) => res.filas.filter((f) => f.estado === est).length;
  const totales = {
    sociedades: res.filas.length, ok: cuenta('OK'), diferencias: cuenta('diferencia'), sin_ddjj: cuenta('sin DDJJ'),
    sin_control: cuenta('sin control'), sin_cuenta: cuenta('sin cuenta') + cuenta('sin CUIT'), errores: cuenta('error'),
    boletas_archivadas: res.filas.filter((f) => f.boleta === 'archivada' || f.boleta === 'reemplazada').length,
    bloqueado: Boolean(bloqueado),
  };
  fs.mkdirSync(SALIDA, { recursive: true });
  const hora = `${String(hoy.getHours()).padStart(2, '0')}${String(hoy.getMinutes()).padStart(2, '0')}${String(hoy.getSeconds()).padStart(2, '0')}`;
  const datos = path.join(TMP, 'datos.json');
  fs.writeFileSync(datos, JSON.stringify(res));
  const excel = path.join(SALIDA, `DDJJ SSHH QUILMES ${PERIODO} ${FECHA} ${hora}.xlsx`);
  const x = py('excel', datos, excel);
  console.log(JSON.stringify({ periodo: PERIODO, ...res, totales, excel: x.error ? null : excel, excel_error: x.error, salida: SALIDA }, null, 2));
})();
