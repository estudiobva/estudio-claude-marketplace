#!/usr/bin/env node
/**
 * cuadro-marcha.js — baja el cuadro de marcha (cronograma de cuotas) de los
 * prestamos vigentes de todas las sociedades, banco por banco, con login automatico.
 *
 *   node scripts/cuadro-marcha.js                              todos los bancos, todas las sociedades
 *   node scripts/cuadro-marcha.js --banco=NACION,MACRO
 *   node scripts/cuadro-marcha.js --banco=MACRO --nombre="<sociedad>"
 *   node scripts/cuadro-marcha.js --banco=NACION --relevar     guarda la pantalla de prestamos de cada empresa
 *
 * Argumentos:
 *   --banco       Uno o varios separados por coma, o "todos" (default): los que tienen
 *                 navegacion de prestamos en lib/prestamos/.
 *   --cuit / --nombre   Solo esas sociedades: una o varias separadas por coma (CUIT o
 *                 nombre de la carpeta de la unidad). Sirve para reprocesar las que fallaron.
 *   --salida      Carpeta destino (default ~/Documents/BVA-salidas/cuadro-marcha,
 *                 o $BVA_SALIDAS_PATH/cuadro-marcha).
 *   --unidad      Archiva en la unidad: [Sociedad]/03-Bancos y Conciliaciones/[BANCO]/Prestamos/.
 *   --relevar     Ademas guarda texto + captura de la pantalla de prestamos de cada
 *                 empresa en ~/bva-inspeccion/cuadro-marcha-<fecha>/ (para ajustar la
 *                 navegacion de un banco). Quedan en la PC, no en la unidad.
 *   --ver         Navegador visible (solo si el usuario lo pide). Por defecto, oculto.
 *
 * Por cada banco: entra con Claves_Bancos.xlsx (una vez, o una por empresa si el
 * banco no deja cambiar de empresa adentro), recorre las empresas que muestra el
 * banco, las empareja con las carpetas de la unidad (lib/prestamos/sociedades.js),
 * entra a Prestamos y baja el cuadro de marcha de cada prestamo vigente:
 *   [salida]/[N] - [Sociedad] - [CUIT]/[BANCO]/[sociedad]_[banco]_prestamo-[nro]_cuadro-marcha_[AAAA-MM-DD].[ext]
 * Nunca pisa: si ya hay uno igual lo saltea; si hay otro distinto agrega _2.
 * Siempre cierra la sesion del banco al terminar (Macro bloquea 10 minutos si no).
 *
 * Salida: JSON por stdout con prestamos y archivos por banco/sociedad; progreso por
 * stderr. Nunca imprime claves.
 */

const fs     = require('fs');
const os     = require('os');
const path   = require('path');
const crypto = require('crypto');
const { chromium } = require('playwright');

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const cb     = require('../lib/claves-bancos');
const bancos = require('../lib/bancos');
const ex     = require('../lib/extractos');
const prest  = require('../lib/prestamos');
const socs   = require('../lib/prestamos/sociedades');

loadEnv();
const args = parseArgs();
const VER = 'ver' in args;
const UNIDAD = 'unidad' in args;
const RELEVAR = 'relevar' in args;
const { launchOptions, HEADLESS: HEADLESS_ENV } = require('../lib/launch');
const VISIBLE = VER || !HEADLESS_ENV;
const SALIDA = args['salida'] || path.join(process.env.BVA_SALIDAS_PATH ||
  path.join(os.homedir(), 'Documents', 'BVA-salidas'), 'cuadro-marcha');
const HOY = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Argentina/Buenos_Aires' });
const SELLO = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
const DIR_RELEVAR = path.join(os.homedir(), 'bva-inspeccion', `cuadro-marcha-${SELLO}`);

const QUE_HACER = {
  credenciales_invalidas: 'Verifica la fila del banco en Claves_Bancos.xlsx. No reintentes enseguida: el banco bloquea el usuario.',
  bloqueado: 'El usuario esta bloqueado: hay que desbloquearlo desde el banco. No reintentar.',
  segundo_factor: 'El banco pidio token o codigo: correr con --ver (si el usuario lo pide) para ingresarlo.',
  captcha: 'El banco pidio captcha: correr con --ver (si el usuario lo pide) para resolverlo.',
  sin_credenciales: 'Completa usuario y clave de ese banco en Claves_Bancos.xlsx.',
  sesion_activa: 'Quedo una sesion abierta en el banco: esperar 10 minutos y volver a correr ese banco.',
  clave_vencida: 'La clave esta vencida: hay que cambiarla entrando al banco a mano y actualizar Claves_Bancos.xlsx.',
  antibots: 'El banco no deja entrar al navegador automatizado. Probar la clave a mano en Chrome.',
  sin_funcionalidad: 'El banco no habilita la consulta de prestamos para esa empresa en el home banking: pedirlo al administrador del banco o pedir el cuadro al oficial de cuenta.',
  banco_sin_servicio: 'El banco no respondio en ese momento (pasa de noche en Macro): volver a correr solo esas sociedades con --nombre=A,B,... mas tarde.',
  sin_conexion: 'Se corto internet o la PC se suspendio: volver a correr ese banco (con la PC despierta). Si quedo una sesion abierta en Macro, esperar 10 minutos.',
  sin_navegacion: 'Ese banco todavia no tiene la navegacion de prestamos (lib/prestamos/).',
  sin_relevar: 'Hay prestamos pero la pantalla todavia no se sabe leer: correr con --relevar y ajustar lib/prestamos/<banco>.js.',
};
const conQueHacer = (code, mensaje) => ({ code, mensaje, ...(QUE_HACER[code] ? { que_hacer: QUE_HACER[code] } : {}) });
const codigoDe = (e) => e.code || (/net::ERR_/.test(e.message) ? 'sitio_rechazo' : 'error');
const log = (m) => process.stderr.write(`[cuadro-marcha] ${m}\n`);
const slugArchivo = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').toLowerCase();

// ── destino y archivado ─────────────────────────────────────────────────────
function dirDestino(soc, banco) {
  if (!soc) return path.join(SALIDA, '_sin_sociedad', banco);
  if (!UNIDAD) return path.join(SALIDA, path.basename(soc.dir), banco);
  const c = ex.carpetaBanco(soc, banco);
  if (c.candidatas) throw Object.assign(new Error(`Hay varias carpetas de ${banco} en ${soc.nombre}: ${c.candidatas.join(' | ')}`), { code: 'sin_carpeta' });
  return path.join(c.dir, 'Prestamos');
}

const sha = (b) => crypto.createHash('sha256').update(b).digest('hex');
function archivar(buf, dir, nombre) {
  if (!buf || !buf.length) throw new Error(`archivo vacio: ${nombre}`);
  if (/\.pdf$/i.test(nombre) && buf.slice(0, 1024).indexOf('%PDF') < 0) throw new Error(`${nombre} no es un PDF valido`);
  fs.mkdirSync(dir, { recursive: true });
  const h = sha(buf);
  for (const f of fs.readdirSync(dir)) {
    const p = path.join(dir, f);
    try { if (fs.statSync(p).isFile() && sha(fs.readFileSync(p)) === h) return { estado: 'ya_estaba', destino: p }; } catch { /* en uso */ }
  }
  const ext = path.extname(nombre), raiz = nombre.slice(0, -ext.length || undefined);
  let destino = path.join(dir, nombre);
  for (let n = 2; fs.existsSync(destino); n++) destino = path.join(dir, `${raiz}_${n}${ext}`);
  fs.writeFileSync(destino, buf, { flag: 'wx' });
  return { estado: 'guardado', destino };
}

// Sin --relevar igual se guarda cuando hay algo que el script no sabe leer.
async function relevar(page, banco, empresa, extra = '', forzar = false) {
  if (!RELEVAR && !forzar) return null;
  const dir = path.join(DIR_RELEVAR, banco);
  fs.mkdirSync(dir, { recursive: true });
  const base = path.join(dir, slugArchivo(empresa) + extra);
  await page.screenshot({ path: `${base}.png`, fullPage: true }).catch(() => {});
  fs.writeFileSync(`${base}.txt`, `${page.url()}\n\n${await page.innerText('body').catch(() => '')}`);
  fs.writeFileSync(`${base}.html`, await page.content().catch(() => ''));
  return `${base}.png`;
}

// ── navegador ───────────────────────────────────────────────────────────────
async function abrirNavegador() {
  const extra = { headless: !VISIBLE, slowMo: VISIBLE ? 250 : 0 };
  try { return await chromium.launch(launchOptions({ channel: 'chrome', ...extra })); }
  catch { log('Chrome no disponible, uso Chromium de Playwright.'); return chromium.launch(launchOptions(extra)); }
}

async function entrarAlBanco(browser, banco) {
  const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
  const page = await context.newPage();
  let cred = await cb.buscar({ banco: bancos.nombreEnClaves(banco) });
  try { await bancos.adaptador(banco).login(page, cred, {}); }
  catch (e) { await context.close().catch(() => {}); throw e; }
  finally { cred = null; }
  return { context, page };
}

// ── una empresa ─────────────────────────────────────────────────────────────
async function procesarEmpresa(page, mod, banco, empresa, soc) {
  const r = { empresa, sociedad: soc ? `${soc.nombre} (${soc.cuit})` : null, prestamos: [], guardados: [], ya_estaban: [], error: null };
  await mod.entrar(page, empresa);
  const lst = await mod.listar(page);
  r.relevado = await relevar(page, banco, empresa);
  if (lst.prestamos === null) {
    r.error = conQueHacer('sin_relevar', 'La pantalla de prestamos muestra algo que el script todavia no sabe leer.');
    if (!RELEVAR) r.relevado = await relevar(page, banco, empresa, '', true);
    return r;
  }
  r.prestamos = lst.prestamos.map(p => ({ id: p.id, descripcion: p.descripcion }));
  if (!lst.prestamos.length) return r;
  if (typeof mod.cuadro !== 'function') {
    r.error = conQueHacer('sin_relevar', `${lst.prestamos.length} prestamo(s): falta la descarga del cuadro de marcha en lib/prestamos/${banco.toLowerCase()}.js.`);
    return r;
  }
  const dir = dirDestino(soc, banco);
  const nombreSoc = soc ? ex.slug(soc.nombre) : slugArchivo(empresa).replace(/-/g, '_');
  r.prestamos = [];
  for (const p of lst.prestamos) {
    const avisos = [];
    const guardar = (buf, sugerido, tipo = 'cuadro-marcha') => {
      const ext = (path.extname(sugerido || '') || '.pdf').toLowerCase();
      const res = archivar(buf, dir, `${nombreSoc}_${banco.toLowerCase()}_prestamo-${slugArchivo(p.id)}_${tipo}_${HOY}${ext}`);
      (res.estado === 'guardado' ? r.guardados : r.ya_estaban).push(res.destino);
      log(`  ${res.estado === 'guardado' ? 'guardado' : 'ya estaba'}: ${res.destino}`);
    };
    const aviso = (t) => { avisos.push(t); log(`  aviso: ${t}`); };
    // La tabla del cuadro (cuota, vencimiento, capital, intereses...) va ademas como planilla.
    const tabla = (encabezado, filas) => {
      const csv = (v) => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
      guardar(Buffer.from('\ufeff' + [encabezado, ...filas].map(f => f.map(csv).join(';')).join('\r\n'), 'utf8'), 'plan.csv');
    };
    try { await mod.cuadro(page, p, { guardar, aviso, tabla }); }
    catch (e) { r.error = conQueHacer(codigoDe(e), `prestamo ${p.id}: ${e.message.split('\n')[0]}`); }
    r.prestamos.push({ id: p.id, descripcion: p.descripcion, ...(p.datos ? { datos: p.datos } : {}), ...(avisos.length ? { avisos } : {}) });
  }
  return r;
}

// ── un banco ────────────────────────────────────────────────────────────────
async function procesarBanco(browser, banco, filtro) {
  const mod = prest.MODULOS[banco];
  const out = { banco, empresas: [], sin_sociedad: [], error: null };
  if (!mod) { out.error = conQueHacer('sin_navegacion', `${banco}: sin navegacion de prestamos.`); return out; }

  log(`${banco}: entrando…`);
  let { context, page } = await entrarAlBanco(browser, banco);
  try {
    const lista = await mod.empresas(page);
    log(`${banco}: ${lista.length} empresas.`);
    const objetivo = lista.map(e => (typeof e === 'string' ? { nombre: e } : e))
      .map(e => ({ empresa: e.nombre, soc: socs.emparejar(e.nombre, { cuit: e.cuit }) }))
      .filter(x => !filtro || (x.soc && filtro.cuits.has(x.soc.cuit)));
    out.sin_sociedad = objetivo.filter(x => !x.soc).map(x => x.empresa);
    if (filtro && !objetivo.length) log(`${banco}: ninguna de ${filtro.nombre} figura entre las empresas del banco.`);

    for (let i = 0; i < objetivo.length; i++) {
      const { empresa, soc } = objetivo[i];
      log(`${banco} [${i + 1}/${objetivo.length}] ${empresa}${soc ? ` → ${soc.nombre}` : ' → (sin sociedad)'}`);
      if (mod.sesion === 'por_empresa' && i > 0) {
        await mod.salir(page);
        await context.close().catch(() => {});
        ({ context, page } = await entrarAlBanco(browser, banco));
      }
      try {
        const r = await procesarEmpresa(page, mod, banco, empresa, soc);
        out.empresas.push(r);
        log(`  ${r.prestamos.length} prestamo(s)${r.error ? ` · ${r.error.code}` : ''}`);
      } catch (e) {
        const r = { empresa, sociedad: soc ? `${soc.nombre} (${soc.cuit})` : null, error: conQueHacer(codigoDe(e), e.message.split('\n')[0]) };
        r.relevado = await relevar(page, banco, empresa, '-error', true);
        out.empresas.push(r);
        log(`  ERROR ${r.error.code} — ${r.error.mensaje}`);
        if (/ERR_(INTERNET_DISCONNECTED|NETWORK_CHANGED|NAME_NOT_RESOLVED|CONNECTION_)/.test(e.message) || /^chrome-error:/.test(page.url())) {
          throw Object.assign(new Error('Se corto la conexion a internet (o la PC se suspendio) en el medio de la corrida.'), { code: 'sin_conexion' });
        }
        if (/sesi[oó]n.*(expir|venc)/i.test(await page.innerText('body').catch(() => ''))) {
          throw Object.assign(new Error('La sesion del banco vencio en el medio de la corrida.'), { code: 'sesion_vencida' });
        }
      }
    }
  } finally {
    await mod.salir(page).catch(() => {});
    await context.close().catch(() => {});
  }
  return out;
}

// ── main ────────────────────────────────────────────────────────────────────
(async () => {
  let filtro = null, lista;
  try {
    if (args['cuit'] || args['nombre']) {
      const pedidas = [...String(args['cuit'] || '').split(',').filter(Boolean).map(c => ex.sociedad({ cuit: c.trim() })),
                       ...String(args['nombre'] || '').split(',').filter(Boolean).map(n => ex.sociedad({ nombre: n.trim() }))];
      filtro = { nombre: pedidas.map(s => s.nombre).join(', '), cuits: new Set(pedidas.map(s => s.cuit)) };
    }
    const pedido = !args['banco'] || /^todos$/i.test(args['banco']) ? prest.bancos
      : args['banco'].split(',').map(b => bancos.canonico(b)).filter(Boolean);
    lista = pedido;
  } catch (e) { fail(e.message, { code: e.code || 'parametros' }); }

  log(`bancos: ${lista.join(', ')}${filtro ? ` · sociedades: ${filtro.nombre}` : ' · todas las sociedades'}`);
  log(`salida: ${UNIDAD ? '[Sociedad]/03-Bancos y Conciliaciones/[BANCO]/Prestamos (unidad)' : SALIDA}`);
  log(`navegador: ${VISIBLE ? 'visible (pedido por el usuario)' : 'oculto'}${RELEVAR ? ` · relevando en ${DIR_RELEVAR}` : ''}.`);

  const browser = await abrirNavegador();
  const resultados = [];
  for (const banco of lista) {
    try { resultados.push(await procesarBanco(browser, banco, filtro)); }
    catch (e) {
      resultados.push({ banco, error: conQueHacer(codigoDe(e), e.message.split('\n')[0]) });
      log(`${banco}: ERROR ${codigoDe(e)} — ${e.message.split('\n')[0]}`);
    }
  }
  await browser.close().catch(() => {});

  const todas = resultados.flatMap(r => (r.empresas || []).map(e => ({ banco: r.banco, ...e })));
  const resumen = {
    con_prestamos: todas.filter(e => e.prestamos && e.prestamos.length)
      .map(e => ({ banco: e.banco, sociedad: e.sociedad || e.empresa, prestamos: e.prestamos.length, archivos: (e.guardados || []).length + (e.ya_estaban || []).length })),
    errores: todas.filter(e => e.error).map(e => ({ banco: e.banco, empresa: e.empresa, ...e.error }))
      .concat(resultados.filter(r => r.error).map(r => ({ banco: r.banco, ...r.error }))),
  };
  const salida = { ok: !resumen.errores.length, fecha: HOY, salida: UNIDAD ? 'unidad' : SALIDA, ...(RELEVAR ? { relevado: DIR_RELEVAR } : {}), resumen, resultados };
  try {
    fs.mkdirSync(SALIDA, { recursive: true });
    fs.writeFileSync(path.join(SALIDA, `cuadro-marcha-${SELLO}.json`), JSON.stringify(salida, null, 2));
    // Planilla para el estudio: una fila por prestamo (y una por empresa con error).
    const COLS = ['banco', 'empresa_en_banco', 'sociedad', 'prestamo', 'producto', 'fecha_otorgamiento', 'importe', 'tna',
      'capital_adeudado / deuda', 'proxima_cuota', 'proximo_pago', 'vencimiento', 'cuotas', 'cuotas_pagas', 'ultima_cuota_paga', 'archivos', 'avisos_o_error'];
    const filas = [];
    for (const e of todas) {
      for (const p of e.prestamos || []) {
        const d = p.datos || {};
        filas.push([e.banco, e.empresa, e.sociedad, p.id, d.producto, d.fecha_otorgamiento, d.importe, d.tna,
          d.capital_adeudado || d.deuda_a_la_fecha, d.proxima_cuota, d.proximo_pago, d.vencimiento, d.cuotas, d.cuotas_pagas, d.ultima_cuota_paga,
          [...(e.guardados || []), ...(e.ya_estaban || [])].filter(f => f.includes(slugArchivo(p.id))).map(f => path.basename(f)).join(' | '),
          (p.avisos || []).join(' | ')]);
      }
      if (e.error) filas.push([e.banco, e.empresa, e.sociedad, ...Array(COLS.length - 4).fill(''), `${e.error.code}: ${e.error.mensaje}`]);
    }
    const csv = (v) => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
    fs.writeFileSync(path.join(SALIDA, `cuadro-marcha-${SELLO}.csv`),
      '\ufeff' + [COLS, ...filas].map(f => f.map(csv).join(';')).join('\r\n'));
    salida.planilla = path.join(SALIDA, `cuadro-marcha-${SELLO}.csv`);
  } catch { /* informativo */ }
  console.log(JSON.stringify(salida, null, 2));
  process.exit(salida.ok ? 0 : 1);
})();
