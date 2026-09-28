#!/usr/bin/env node
/**
 * extractos-banco.js — login + descarga + renombrado + archivado de extractos.
 *
 *   node scripts/extractos-banco.js --nombre="<sociedad>" --banco=GALICIA --desde=202601 --hasta=202606
 *   node scripts/extractos-banco.js --cuit=30000000007 --banco=GALICIA,SANTANDER,COMAFI --periodo=08/2026
 *   node scripts/extractos-banco.js --nombre="<sociedad>" --banco=todos --desde=202601 --hasta=202603
 *   node scripts/extractos-banco.js --nombre="<sociedad>" --banco=MACRO --periodo=202608 --grabar
 *   node scripts/extractos-banco.js --nombre="<sociedad>" --banco=GALICIA --periodo=202608 --unidad
 *
 * Argumentos:
 *   --cuit / --nombre   Sociedad (el CUIT desempata; el nombre es el de la carpeta de la unidad).
 *   --banco             Uno o varios separados por coma, o "todos" = los que el Maestro de
 *                       Sociedades marca con "Si" para esa sociedad.
 *   --desde / --hasta   Periodos AAAAMM o MM/AAAA (o --periodo para uno solo). Maximo 24 meses.
 *   --sociedad-claves   Solo Galicia: nombre de la sociedad en el bloque Galicia de
 *                       Claves_Bancos.xlsx, si difiere del de la carpeta.
 *   --grabar            Ademas guarda lo que se clickea en el banco (sin valores tipeados)
 *                       en ~/bva-inspeccion/, para automatizar la navegacion de ese banco.
 *   --salida            Carpeta destino (default: ~/Documents/BVA-salidas/extractos,
 *                       o $BVA_SALIDAS_PATH/extractos).
 *   --unidad            Archiva directo en la unidad compartida, en
 *                       [N] - [Sociedad] - [CUIT]/03-Bancos y Conciliaciones/[BANCO]/[AAAA]/.
 *   --ver               Navegador visible. Sin --ver corre oculto, salvo que algun banco
 *                       pedido no tenga descarga automatica (hoy ninguno): ahi se abre
 *                       visible igual, porque hay que navegar el banco en la ventana.
 *   --control           Despues del login deja la ventana abierta a un puerto local
 *                       (127.0.0.1, --puerto=9333 por defecto) para que Claude la maneje con
 *                       scripts/control.js: navega y descarga; el script sigue renombrando y
 *                       archivando. El login lo sigue haciendo este script.
 *
 * Por cada banco:
 *   1. Loguea con Claves_Bancos.xlsx (una sola vez; si piden captcha/token, te lo pide a vos).
 *   2. Descarga: si el banco tiene descarga automatica la hace sola; si no, en la ventana
 *      elegis la sociedad y bajas los resumenes del periodo, y el script toma cada archivo.
 *   3. Renombra: [sociedad]_[banco]_[AAAA-MM]_mensual.pdf (o _semanaN en Galicia). El periodo
 *      sale del nombre que le pone el banco; si no se puede leer, te lo pregunta.
 *   4. Guarda en [salida]/[N] - [Sociedad] - [CUIT]/[BANCO]/[AAAA]/ (con --unidad, en
 *      03-Bancos y Conciliaciones/[BANCO]/[AAAA]/ reusando la carpeta del banco que ya
 *      exista). Nunca pisa: si ya esta igual lo saltea, si hay otro distinto con el mismo
 *      nombre agrega _2.
 *
 * Salida: JSON con lo archivado y los periodos que quedaron sin archivo. Nunca imprime claves.
 */

const fs       = require('fs');
const os       = require('os');
const path     = require('path');
const readline = require('readline');
const { chromium } = require('playwright');

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const cb  = require('../lib/claves-bancos');
const bancos = require('../lib/bancos');
const ex  = require('../lib/extractos');

loadEnv();
const args = parseArgs();
const VER = 'ver' in args;
const UNIDAD = 'unidad' in args;
const GRABAR = 'grabar' in args;
const CONTROL = 'control' in args;
const PUERTO = Number(args['puerto'] || 9333);
const { launchOptions, HEADLESS: HEADLESS_ENV } = require('../lib/launch');
const SALIDA = ex.salidaBase(args['salida']);
// Se decide en main(), cuando ya se sabe que bancos se piden.
let VISIBLE = VER || !HEADLESS_ENV;

// Que hacer ante cada error, para el JSON (como en los scripts de ARCA).
const QUE_HACER = {
  credenciales_invalidas: 'Verifica la fila del banco en Claves_Bancos.xlsx. No reintentes enseguida: el banco bloquea el usuario.',
  bloqueado: 'El usuario esta bloqueado: hay que desbloquearlo desde el banco. No reintentar.',
  segundo_factor: 'El banco pidio token o codigo: volve a correr con --ver para ingresarlo en la ventana.',
  captcha: 'El banco pidio captcha: volve a correr con --ver para resolverlo en la ventana.',
  sin_credenciales: 'Completa usuario y clave de ese banco en Claves_Bancos.xlsx (node scripts/claves-bancos.js --faltantes).',
  requiere_ventana: 'Ese banco todavia no tiene descarga automatica: correr con --ver (o --control).',
  sin_carpeta: 'Hay varias carpetas para ese banco en la unidad: correr sin --unidad o elegir la carpeta.',
};
const conQueHacer = (code, mensaje) => ({ code, mensaje, ...(QUE_HACER[code] ? { que_hacer: QUE_HACER[code] } : {}) });

const log = (m) => process.stderr.write(`[extractos] ${m}\n`);

// ── consola: todas las preguntas pasan por una sola cola ────────────────────
const rl = readline.createInterface({ input: process.stdin, output: process.stderr });
// Si la entrada se cierra (sin terminal, o stdin redirigido) las preguntas se
// contestan solo desde la ventana del banco; sin ventana, se corta con error.
// Si no hay terminal (Claude lo corre en segundo plano) solo cuenta la ventana.
let terminalCerrada = !process.stdin.isTTY;
rl.on('close', () => { terminalCerrada = true; });
let colaPreguntas = Promise.resolve();
// Cada pregunta se puede contestar en la terminal (Enter) o, si hay ventana,
// con el boton de la barra que se muestra dentro de la pagina del banco: asi no
// depende de que el foco del teclado este en la terminal.
function preguntar(texto, { ventana, conTexto = false } = {}) {
  const p = colaPreguntas.then(() => new Promise((res) => {
    const ac = new AbortController();
    let listo = false;
    const fin = (r) => {
      if (listo) return;
      listo = true;
      ac.abort();
      if (ventana) ventana.ocultar();
      res(String(r == null ? '' : r).trim());
    };
    if (!terminalCerrada) {
      try { rl.question(`\n${texto}\n> `, { signal: ac.signal }, fin); }
      catch { terminalCerrada = true; }
    }
    if (terminalCerrada) process.stderr.write(`\n${texto}\n  (respondé con el boton de la ventana)\n`);
    if (ventana) ventana.mostrar(texto, conTexto).then(fin).catch(() => {});
    else if (terminalCerrada) fin('');
  }));
  colaPreguntas = p.catch(() => {});
  return p;
}

// Barra flotante "Estudio BVA" con el pedido y un boton Listo, en todas las
// pestanas del contexto; se vuelve a dibujar si la pagina navega.
function barraEnVentana(context) {
  let pendiente = null;
  let resolver = null;
  const pintar = (page) => page.evaluate((pend) => {
    const vieja = document.getElementById('bva-barra');
    if (vieja) vieja.remove();
    if (!pend) return;
    const b = document.createElement('div');
    b.id = 'bva-barra';
    b.style.cssText = 'position:fixed;top:12px;right:12px;max-width:380px;z-index:2147483647;background:#1f3a5f;' +
      'color:#fff;font:13px/1.45 system-ui,sans-serif;padding:12px 14px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.35)';
    const t = document.createElement('div');
    t.style.cssText = 'white-space:pre-wrap;margin-bottom:8px';
    t.textContent = 'Estudio BVA\n' + pend.msg.trim();
    b.appendChild(t);
    let inp = null;
    if (pend.conTexto) {
      inp = document.createElement('input');
      inp.placeholder = 'AAAAMM (vacio = descartar)';
      inp.style.cssText = 'padding:6px;border-radius:4px;border:0;width:180px;color:#000;margin-right:8px';
      b.appendChild(inp);
    }
    const btn = document.createElement('button');
    btn.textContent = pend.conTexto ? 'Aceptar' : 'Listo';
    btn.style.cssText = 'padding:7px 16px;border:0;border-radius:4px;background:#2e9d5b;color:#fff;font-weight:600;cursor:pointer';
    btn.onclick = (ev) => { ev.stopPropagation(); window.__bvaListo(inp ? inp.value : ''); };
    b.appendChild(btn);
    document.documentElement.appendChild(b);
  }, pendiente).catch(() => {});
  const binding = context.exposeBinding('__bvaListo', (_src, v) => {
    if (resolver) { const r = resolver; resolver = null; r(v); }
  }).catch(() => {});
  const enganchar = (p) => {
    p.on('load', () => pintar(p));
    // Si la persona cierra la ultima ventana del banco, cuenta como "Listo".
    p.on('close', () => {
      if (!context.pages().length && resolver) { const r = resolver; resolver = null; pendiente = null; r(''); }
    });
  };
  context.pages().forEach(enganchar);
  context.on('page', enganchar);
  return {
    mostrar: async (msg, conTexto) => {
      await binding;
      return new Promise((res) => {
        pendiente = { msg, conTexto };
        resolver = res;
        context.pages().forEach(pintar);
      });
    },
    ocultar: () => { pendiente = null; resolver = null; context.pages().forEach(pintar); },
  };
}

async function abrirNavegador() {
  // Con --control se abre ademas un puerto de depuracion SOLO en 127.0.0.1.
  const extra = { headless: !VISIBLE, slowMo: VISIBLE ? 250 : 0,
    ...(CONTROL ? { args: [...launchOptions().args,
      `--remote-debugging-port=${PUERTO}`, '--remote-debugging-address=127.0.0.1'] } : {}) };
  try { return await chromium.launch(launchOptions({ channel: 'chrome', ...extra })); }
  catch { log('Chrome no disponible, uso Chromium de Playwright.'); return chromium.launch(launchOptions(extra)); }
}

async function credenciales(banco, soc) {
  if (banco === 'GALICIA') {
    // Galicia se entra con el login propio de cada sociedad (bloque Galicia del
    // archivo); la fila de login unico hoy esta vacia y queda solo de respaldo.
    try { return await cb.buscar({ banco: 'GALICIA', cuit: soc.cuit, sociedad: args['sociedad-claves'] || soc.nombre }); }
    catch (e) {
      if (args['sociedad-claves']) throw e;
      try { return await cb.buscar({ banco: 'GALICIA' }); }
      catch {
        throw Object.assign(new Error(`${e.message} Si en el archivo figura con otro nombre, ` +
          'pasa --sociedad-claves="<nombre como figura en el bloque Galicia>".'), { code: 'sin_credenciales' });
      }
    }
  }
  return cb.buscar({ banco: bancos.nombreEnClaves(banco), sociedad: soc.nombre });
}

// ── grabacion de la navegacion (solo clicks y selects; nunca lo tipeado) ────
const GRABADOR = `(() => {
  if (window.__bvaGrabador) return; window.__bvaGrabador = true;
  const desc = (e) => {
    const el = e.closest('a,button,[role=button],[role=menuitem],[role=tab],[role=option],li,select,input,label,td,span,div') || e;
    return { tag: el.tagName, id: el.id || null, name: el.getAttribute('name'), type: el.getAttribute('type'),
      role: el.getAttribute('role'), aria: el.getAttribute('aria-label'), testid: el.getAttribute('data-testid'),
      texto: (el.innerText || el.value || '').trim().replace(/\\s+/g, ' ').slice(0, 80),
      clase: String(el.className || '').split(' ').slice(0, 3).join(' '), href: el.getAttribute('href') };
  };
  const propio = (ev) => ev.target && ev.target.closest && ev.target.closest('#bva-barra');
  document.addEventListener('click', (ev) => !propio(ev) && window.__bvaGrabar && window.__bvaGrabar({ ev: 'click', url: location.href, el: desc(ev.target) }), true);
  document.addEventListener('change', (ev) => {
    const t = ev.target; if (!window.__bvaGrabar) return;
    const d = { ev: 'change', url: location.href, el: desc(t) };
    if (t.tagName === 'SELECT') d.opcion = t.options[t.selectedIndex] && t.options[t.selectedIndex].text;
    else if (t.type === 'checkbox' || t.type === 'radio') d.marcado = t.checked;
    else d.completado = !!t.value;   // nunca el valor
    window.__bvaGrabar(d);
  }, true);
})();`;

// ── un banco ────────────────────────────────────────────────────────────────
async function procesarBanco(browser, banco, soc, periodos) {
  const r = { banco, guardados: [], ya_estaban: [], descartados: [], rechazados: [], sin_archivo: [], error: null };
  const adaptador = bancos.adaptador(banco);

  // Carpeta de destino, antes de loguear (si hay que elegir, se pregunta ya).
  let cb_ = UNIDAD ? ex.carpetaBanco(soc, banco) : ex.carpetaSalida(SALIDA, soc, banco);
  if (cb_.candidatas) {
    const lista = cb_.candidatas.map((c, i) => `  ${i + 1}. ${c}`).join('\n');
    const r2 = await preguntar(`Hay ${cb_.candidatas.length} carpetas para ${banco} en ${soc.nombre}:\n${lista}\n  Numero de la que va:`);
    const elegida = cb_.candidatas[Number(r2) - 1];
    if (!elegida) throw Object.assign(new Error('No se eligio carpeta de banco.'), { code: 'sin_carpeta' });
    cb_ = ex.carpetaBanco(soc, banco, { elegida });
  }
  const dirBanco = cb_.dir;
  log(`${banco}: destino ${dirBanco}${cb_.creada ? ' (se crea)' : ''}`);

  let cred = await credenciales(banco, soc);
  const context = await browser.newContext({ locale: 'es-AR', acceptDownloads: true });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), `bva-${banco.toLowerCase()}-`));
  const ventana = VISIBLE ? barraEnVentana(context) : null;
  const pedirAccion = VISIBLE ? (msg) => preguntar(`  ${msg}`, { ventana }) : null;

  // Grabacion
  const grabacion = [];
  let grabando = false;
  if (GRABAR) {
    await context.exposeBinding('__bvaGrabar', (_src, d) => { if (grabando) grabacion.push({ t: new Date().toISOString(), ...d }); });
    await context.addInitScript(GRABADOR);
  }

  // Captura de archivos: descargas y PDFs que el banco abre en el visor.
  let cola = Promise.resolve();
  const vistos = new Set();
  const procesar = (origen, sugerido) => {
    cola = cola.then(() => archivarUno(origen, sugerido)).catch((e) => {
      r.rechazados.push({ archivo: sugerido, motivo: e.message });
      log(`rechazado ${sugerido}: ${e.message}`);
    });
  };
  const engancharPagina = (page) => {
    page.on('download', async (d) => {
      const sugerido = d.suggestedFilename();
      const destino = path.join(tmp, `${Date.now()}-${sugerido}`);
      try { await d.saveAs(destino); procesar(destino, sugerido); if (GRABAR && grabando) grabacion.push({ ev: 'download', archivo: sugerido, url: page.url() }); }
      catch (e) { log(`no pude guardar la descarga ${sugerido}: ${e.message}`); }
    });
  };
  context.on('page', engancharPagina);
  context.on('response', async (resp) => {
    const h = resp.headers();
    if (!/application\/pdf/i.test(h['content-type'] || '') || /attachment/i.test(h['content-disposition'] || '')) return;
    if (vistos.has(resp.url())) return;
    vistos.add(resp.url());
    try {
      const buf = await resp.body();
      const m = (h['content-disposition'] || '').match(/filename\*?=(?:UTF-8'')?"?([^";]+)/i);
      const sugerido = decodeURIComponent(m ? m[1] : path.basename(new URL(resp.url()).pathname) || 'extracto.pdf');
      const destino = path.join(tmp, `${Date.now()}-${sugerido.endsWith('.pdf') ? sugerido : sugerido + '.pdf'}`);
      fs.writeFileSync(destino, buf);
      procesar(destino, path.basename(destino).replace(/^\d+-/, ''));
    } catch { /* respuestas que no se pueden leer: se ignoran */ }
  });

  const conArchivo = new Set();
  // Rutas del JSON relativas a la carpeta de la sociedad (en la salida o en la unidad).
  const baseSoc = UNIDAD ? soc.dir : path.join(SALIDA, path.basename(soc.dir));
  async function archivarUno(origen, sugerido) {
    let det = ex.periodoDeArchivo(sugerido);
    if (!det || !periodos.includes(det.periodo)) {
      const leido = det ? ` (el nombre dice ${det.periodo}, fuera de lo pedido)` : '';
      const resp = await preguntar(`  ${banco}: bajaste "${sugerido}"${leido}.\n  Periodo AAAAMM al que corresponde (vacio = descartar):`,
                                   { ventana, conTexto: true });
      const p = ex.periodo(resp);
      if (!p) { r.descartados.push(sugerido); log(`descartado: ${sugerido}`); return; }
      det = { periodo: p, dia: det && det.periodo === p ? det.dia : null };
    }
    const semana = banco === 'GALICIA' && det.dia ? Math.ceil(det.dia / 7) : null;
    const ext = (path.extname(sugerido) || '.pdf').toLowerCase();
    const nombre = ex.nombreArchivo({ soc, banco, periodo: det.periodo, semana, ext });
    const res = ex.archivar({ origen, dirBanco, periodo: det.periodo, nombre });
    conArchivo.add(det.periodo);
    const rel = path.relative(baseSoc, res.destino);
    if (res.estado === 'guardado') { r.guardados.push(rel); log(`guardado: ${rel}`); }
    else { r.ya_estaban.push(rel); log(`ya estaba (mismo contenido): ${rel}`); }
  }

  try {
    const page = await context.newPage();   // context.on('page') ya la engancha
    log(`${banco}: entrando…`);
    await adaptador.login(page, cred, { pedirAccion });
    cred = null;
    log(`${banco}: adentro.`);
    grabando = true;

    if (typeof adaptador.descargar === 'function') {
      await adaptador.descargar(page, { sociedad: soc, periodos });
    } else {
      if (!pedirAccion) throw Object.assign(new Error(`${banco} todavia no tiene descarga automatica: corre con --ver.`), { code: 'requiere_ventana' });
      if (CONTROL) log(`CONTROL ${banco} adentro — ventana disponible en http://127.0.0.1:${PUERTO} (scripts/control.js)`);
      await preguntar(CONTROL ?
        `  ${banco} — Claude maneja esta ventana: va a bajar los resumenes de ${periodos[0]} a ` +
        `${periodos[periodos.length - 1]}.\n  No hace falta que toques nada; al terminar toca "Listo" el propio Claude ` +
        `(o vos, o Enter aca).` :
        `  ${banco} — en la ventana del navegador:\n` +
        `    1. Elegi la sociedad ${soc.nombre} (CUIT ${soc.cuit}) si el banco lo pide.\n` +
        `    2. Anda a Resumenes / Extractos y baja los de ${periodos[0]} a ${periodos[periodos.length - 1]}` +
        `${banco === 'GALICIA' ? ' (semanales y mensual)' : ''}, todas las cuentas.\n` +
        `    Cada archivo que bajes lo renombro y lo guardo solo.\n` +
        `  Cuando termines toca "Listo" en la ventana (o Enter aca).`, { ventana });
    }
    await page.waitForTimeout(1500).catch(() => {});   // la ventana puede haberse cerrado ("Listo" por cierre)
    await cola;
  } catch (e) {
    r.error = conQueHacer(e.code || 'error', e.message.split('\n')[0]);
    log(`${banco}: ERROR ${r.error.code} — ${r.error.mensaje}`);
    try {
      const dir = path.join(os.homedir(), 'bva-inspeccion');
      fs.mkdirSync(dir, { recursive: true });
      const base = path.join(dir, `${banco.toLowerCase()}-error-${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}`);
      const pg = context.pages()[context.pages().length - 1];
      if (pg) {
        await pg.screenshot({ path: `${base}.png`, fullPage: true });
        fs.writeFileSync(`${base}.txt`, `${pg.url()}\n\n${(await pg.innerText('body').catch(() => '')).slice(0, 4000)}`);
        r.captura = `${base}.png`;
        log(`${banco}: captura del error en ${base}.png`);
      }
    } catch { /* sin captura */ }
    await cola.catch(() => {});
  } finally {
    if (GRABAR && grabacion.length) {
      const dir = path.join(os.homedir(), 'bva-inspeccion');
      fs.mkdirSync(dir, { recursive: true });
      const f = path.join(dir, `${banco.toLowerCase()}-grabacion-${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.json`);
      fs.writeFileSync(f, JSON.stringify({ banco, sociedad: soc.nombre, periodos, pasos: grabacion }, null, 2));
      r.grabacion = f;
      log(`grabacion guardada en ${f}`);
    }
    await context.close().catch(() => {});
    fs.rmSync(tmp, { recursive: true, force: true });   // solo la carpeta temporal propia
  }

  r.sin_archivo = periodos.filter(p => !conArchivo.has(p));
  return r;
}

// ── main ────────────────────────────────────────────────────────────────────
(async () => {
  let soc, periodos, lista;
  try {
    soc = ex.sociedad({ cuit: args['cuit'], nombre: args['nombre'] });
    periodos = args['periodo'] ? ex.rango(args['periodo']) : ex.rango(args['desde'], args['hasta']);
    if (!args['banco']) throw new Error('Falta --banco (uno, varios separados por coma, o "todos").');
    lista = /^todos$/i.test(args['banco'])
      ? ex.bancosDelMaestro(soc.cuit)
      : args['banco'].split(',').map(b => bancos.canonico(b)).filter(Boolean);
  } catch (e) {
    rl.close();
    fail(e.message, { code: e.code || 'parametros' });
  }

  const sinAdaptador = lista.filter(b => !bancos.bancos.includes(b));
  lista = lista.filter(b => bancos.bancos.includes(b));
  log(`sociedad: ${soc.nombre} (${soc.cuit})`);
  log(`periodos: ${periodos[0]} a ${periodos[periodos.length - 1]} (${periodos.length})`);
  log(`bancos: ${lista.join(', ') || '(ninguno)'}${sinAdaptador.length ? ` · sin adaptador: ${sinAdaptador.join(', ')}` : ''}`);
  log(`salida: ${UNIDAD ? path.join(soc.dir, ex.BANCOS_DIR) + ' (unidad)' : path.join(SALIDA, path.basename(soc.dir))}`);
  // Sin descarga automatica hay que navegar el banco en la ventana: se abre visible.
  const manuales = lista.filter(b => typeof bancos.adaptador(b).descargar !== 'function');
  // Con --control tambien: la barra "Listo" de la ventana es la que marca el fin.
  if (!VISIBLE && (manuales.length || CONTROL)) {
    VISIBLE = true;
    log(`navegador visible: ${CONTROL ? '--control' : `${manuales.join(', ')} todavia no tiene(n) descarga automatica`}.`);
  }

  const resultados = [];
  const browser = lista.length ? await abrirNavegador() : null;
  // Si se cierra el navegador entero, se termina con lo que haya en vez de colgarse.
  let cerrado = false;
  if (browser) browser.on('disconnected', () => {
    if (cerrado) return;
    log('se cerro el navegador: termino la corrida con lo que se llego a guardar.');
    console.log(JSON.stringify({ ok: false, cortado: 'navegador cerrado', sociedad: { nombre: soc.nombre, cuit: soc.cuit }, periodos, resultados }, null, 2));
    process.exit(1);
  });
  for (const banco of lista) {
    try { resultados.push(await procesarBanco(browser, banco, soc, periodos)); }
    catch (e) {
      resultados.push({ banco, error: conQueHacer(e.code || 'error', e.message) });
      log(`${banco}: ERROR ${e.code || 'error'} — ${e.message}`);
    }
    const ultimo = resultados[resultados.length - 1];
    if (ultimo.error && ['credenciales_invalidas', 'bloqueado'].includes(ultimo.error.code)) {
      log(`${banco}: ${ultimo.error.mensaje} No se reintenta.`);
    }
  }
  cerrado = true;
  if (browser) await browser.close().catch(() => {});
  rl.close();

  console.log(JSON.stringify({
    ok: resultados.every(r => !r.error),
    sociedad: { nombre: soc.nombre, cuit: soc.cuit },
    periodos,
    salida: UNIDAD ? path.join(soc.dir, ex.BANCOS_DIR) : path.join(SALIDA, path.basename(soc.dir)),
    bancos_sin_adaptador: sinAdaptador,
    resultados,
  }, null, 2));
  process.exit(resultados.some(r => r.error) ? 1 : 0);
})();
