#!/usr/bin/env node
/**
 * correos-organismos.js — releva los correos electronicos declarados de TODAS
 * las sociedades del archivo de claves (o de las que se pidan) en ARCA y ARBA,
 * y arma UN Excel con lo que hay declarado, lo que hay que revisar y lo que
 * falta.
 *
 *   node scripts/correos-organismos.js
 *   node scripts/correos-organismos.js --solo=3,7,12            (posiciones del Excel de claves)
 *   node scripts/correos-organismos.js --nombres="<A>;<B>"
 *   node scripts/correos-organismos.js --organismos=arba        (solo ARBA; default arca,arba)
 *   node scripts/correos-organismos.js --consolidar             (no entra a ningun sitio: rearma el Excel)
 *
 * Argumentos:
 *   --solo / --nombres   Filtran que se releva. El Excel junta todo lo relevado en el dia.
 *   --organismos         arca, arba o arca,arba.
 *   --rehacer            Vuelve a relevar lo ya hecho hoy con exito (por defecto se saltea: reanudable).
 *   --dominio-estudio    Dominio propio del estudio (default estudiobva.com): los correos de otro
 *                        dominio se marcan como externos para confirmar que sigan activos.
 *   --salida             Carpeta base (default ~/Documents/BVA-salidas/correos, $BVA_SALIDAS_PATH). Subcarpeta por dia.
 *   --ver                Navegador visible.
 *
 * Que se lee (ver lib/correos.js):
 *   ARCA  Sistema Registral: correo principal + estado de verificacion, telefono,
 *         adhesion al DFE y la lista completa de correos (Datos registrales).
 *         DFE > "Asociar temas a un correo": a que correo llegan los avisos.
 *   ARBA  Autogestion > Datos de contacto: correos con caracter y fecha de
 *         verificacion, telefonos, y el correo con que se registro la CIT.
 *   AGIP  No se automatiza (el firewall de AGIP bloquea el navegador
 *         automatizado). Queda la columna para completar a mano en las
 *         sociedades que tienen clave de AGIP.
 *
 * Que se puede saber y que no: se ve si hay correo, si el organismo lo da por
 * verificado, si coincide entre organismos, si es del estudio o de terceros y
 * si el dominio tiene servidor de correo (DNS). NO se puede saber desde aca si
 * un buzon de terceros sigue leyendose: esos quedan marcados para confirmar.
 *
 * Solo LEE. Nunca modifica correos, telefonos ni adhesiones.
 * Por que un login por apoderado en ARCA: ver deudas-ddjj-sct-lote.js (loguear
 * 69 veces hace rebotar y un rechazo repetido bloquea la clave de todo el grupo).
 * En ARBA cada sociedad tiene su CIT: un intento por sociedad, sin reintentos
 * de login (ARBA bloquea la CIT tras varios intentos fallidos).
 */

const fs   = require('fs');
const path = require('path');
const dns  = require('dns').promises;
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const { launchOptions } = require('../lib/launch');
const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const claves = require('../lib/claves');
const { login } = require('../lib/arca-login');
const correosLib = require('../lib/correos');

const args = parseArgs();
const VER = 'ver' in args;
const REHACER = 'rehacer' in args;
const SOLO_CONSOLIDAR = 'consolidar' in args;
const ORGS = new Set(String(args['organismos'] || 'arca,arba').toLowerCase().split(',').map(s => s.trim()).filter(Boolean));
const DOMINIO_ESTUDIO = String(args['dominio-estudio'] || 'estudiobva.com').toLowerCase();
const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const log = (m) => process.stderr.write(`[correos] ${m}\n`);

for (const o of ORGS) if (!['arca', 'arba'].includes(o)) fail(`--organismos no reconoce "${o}" (usar arca, arba o arca,arba)`);

const HOY = new Date();
const iso = [HOY.getFullYear(), String(HOY.getMonth() + 1).padStart(2, '0'),
             String(HOY.getDate()).padStart(2, '0')].join('-');
const BASE = args['salida'] || path.join(process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas'), 'correos');
const DIR = path.join(BASE, iso);
const RESUMEN = path.join(DIR, 'lote-resumen.json');

// ── A quien relevar ─────────────────────────────────────────────────────────
loadEnv();
let rows;
try { ({ rows } = claves.loadSync()); }
catch (e) { fail(`No pude leer el archivo de claves: ${e.message}`); }

const todos = [];
const omitidos = [];
const vistos = new Map();
for (const r of rows) {
  if (!r.cuit) { omitidos.push({ pos: r.pos, titular: r.titular, motivo: 'sin CUIT' }); continue; }
  if (vistos.has(r.cuit)) {
    omitidos.push({ pos: r.pos, titular: r.titular, motivo: `CUIT repetido (ya esta como "${vistos.get(r.cuit)}")` });
    continue;
  }
  vistos.set(r.cuit, r.titular);
  const arca = claves.credenciales(r, 'ARCA');
  const arba = claves.credenciales(r, 'ARBA');
  const agip = claves.credenciales(r, 'AGIP');
  todos.push({
    pos: r.pos, titular: r.titular, cuit: r.cuit, tipo: r.tipo,
    arca: arca.clave ? { login: claves.normCuit(arca.usuario) || r.cuit, clave: arca.clave } : null,
    arba: arba.clave ? { usuario: claves.normCuit(arba.usuario) || r.cuit, clave: arba.clave } : null,
    agip: Boolean(agip.clave),
  });
}

const solo = args['solo'] ? new Set(String(args['solo']).split(',').map(x => x.trim())) : null;
const nombres = args['nombres'] ? String(args['nombres']).split(';').map(s => s.trim()).filter(Boolean) : null;
let objetivos = todos;
if (solo) objetivos = objetivos.filter(o => solo.has(String(o.pos).trim()));
if (nombres) {
  const elegidos = new Set();
  for (const n of nombres) {
    const r = claves.buscar(rows, { nombre: n });
    if (r.candidatos) fail(`"${n}" coincide con varios titulares: ${r.candidatos.map(c => c.titular).join(', ')}`);
    elegidos.add(r.row.cuit);
  }
  objetivos = objetivos.filter(o => elegidos.has(o.cuit));
}
if (SOLO_CONSOLIDAR) objetivos = [];
else if (!objetivos.length) fail('No hay ninguna sociedad para relevar con ese filtro.');

// ── Estado (reanudable) ─────────────────────────────────────────────────────
fs.mkdirSync(DIR, { recursive: true });
const estado = fs.existsSync(RESUMEN)
  ? JSON.parse(fs.readFileSync(RESUMEN, 'utf-8'))
  : { fecha: iso, sociedades: {}, dfe: {} };
estado.dfe ||= {};
const guardar = () => fs.writeFileSync(RESUMEN, JSON.stringify(estado, null, 1), 'utf-8');

function registrar(o, org, datos) {
  const s = (estado.sociedades[o.cuit] ||= { pos: o.pos, sociedad: o.titular, cuit: o.cuit });
  s[org] = { ...datos, cuando: new Date().toISOString() };
  guardar();
}
const yaHecha = (o, org) => !REHACER && estado.sociedades[o.cuit]?.[org]?.ok;

const navegadorCerrado = (browser, e) => !browser.isConnected() || /has been closed/i.test(e.message);

// ── ARCA ────────────────────────────────────────────────────────────────────
async function faseArca() {
  const grupos = new Map();
  for (const o of objetivos.filter(o => o.arca)) {
    if (!grupos.has(o.arca.login)) grupos.set(o.arca.login, []);
    grupos.get(o.arca.login).push(o);
  }
  const total = [...grupos.values()].reduce((a, g) => a + g.length, 0);
  log(`ARCA: ${total} sociedad(es) en ${grupos.size} login(s)`);
  let n = 0;

  for (const [cuitLogin, grupo] of grupos) {
    const pendientes = grupo.filter(o => !yaHecha(o, 'arca'));
    const faltaDfe = REHACER || !estado.dfe[cuitLogin]?.ok;
    n += grupo.length - pendientes.length;
    if (!pendientes.length && !faltaDfe) continue;
    log(`── ARCA login ${cuitLogin.slice(0, 2)}…${cuitLogin.slice(-3)}: ${pendientes.length} sociedad(es)`);

    const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 150 } : {}));
    const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1500, height: 950 } });
    const portal = await context.newPage();
    try {
      try {
        await login(portal, { cuitLogin, password: grupo[0].arca.clave });
      } catch (e) {
        const motivo = e.code === 'credenciales_invalidas'
          ? 'ARCA rechazo las credenciales (no se reintento para no bloquear la clave)'
          : e.code === 'captcha' ? 'ARCA pidio captcha en el login' : `login fallo: ${e.message.split('\n')[0]}`;
        log(`   ${motivo}`);
        for (const o of pendientes) { n++; registrar(o, 'arca', { ok: false, error: motivo, code: e.code || 'login' }); }
        continue;
      }

      if (faltaDfe) {
        try {
          const mapa = await correosLib.leerAsociacionesDfe(context, portal);
          estado.dfe[cuitLogin] = { ok: true, mapa, cuando: new Date().toISOString() };
          log(`   DFE: ${Object.keys(mapa).length} representado(s) con correo de avisos`);
        } catch (e) {
          if (navegadorCerrado(browser, e)) throw Object.assign(new Error('Se cerro el navegador. Volve a correrlo y retoma donde quedo.'), { code: 'navegador_cerrado' });
          estado.dfe[cuitLogin] = { ok: false, error: e.message.split('\n')[0] };
          log(`   DFE: no pude leer las asociaciones (${e.message.split('\n')[0]})`);
        }
        guardar();
      }

      let reg = null;
      for (const o of pendientes) {
        n++;
        log(`[ARCA ${n}/${total}] #${o.pos} ${o.titular} (${o.cuit})`);
        let ultimo = null;
        for (let intento = 1; intento <= 2; intento++) {
          try {
            if (!reg || reg.isClosed()) reg = await correosLib.abrirRegistral(context, portal);
            const r = await correosLib.relevarArca(reg, o.cuit);
            registrar(o, 'arca', { ok: true, login: cuitLogin, ...r });
            const todosCorreos = new Set([r.correo, ...(r.sabana?.correos || []).map(c => c.correo)].filter(Boolean));
            log(`   ${todosCorreos.size ? [...todosCorreos].join(', ') : 'SIN CORREO'}${r.estado ? ` (${r.estado})` : ''} · DFE ${r.dfe || '?'}`);
            ultimo = null;
            break;
          } catch (e) {
            ultimo = e;
            if (navegadorCerrado(browser, e)) throw Object.assign(new Error('Se cerro el navegador. Volve a correrlo y retoma donde quedo.'), { code: 'navegador_cerrado' });
            if (['cuit_no_disponible', 'servicio_no_disponible'].includes(e.code) || intento === 2) break;
            log(`   fallo (${e.code || 'error'}): ${e.message.split('\n')[0]} — reabro Sistema Registral y reintento`);
            if (reg) await reg.close().catch(() => {});
            reg = null;
          }
        }
        if (ultimo) {
          const msg = ultimo.message.split('\n')[0];
          log(`   ERROR (${ultimo.code || 'error'}): ${msg}`);
          registrar(o, 'arca', { ok: false, login: cuitLogin, error: msg, code: ultimo.code || 'error' });
          if (ultimo.code === 'servicio_no_disponible') {
            // Sin Sistema Registral en este login no hay nada que hacer con el resto del grupo.
            for (const r of pendientes.slice(pendientes.indexOf(o) + 1)) {
              n++; registrar(r, 'arca', { ok: false, login: cuitLogin, error: msg, code: ultimo.code });
            }
            break;
          }
          if (reg) await reg.close().catch(() => {});
          reg = null;
        }
      }
    } finally {
      await browser.close().catch(() => {});
    }
  }
}

// ── ARBA ────────────────────────────────────────────────────────────────────
async function faseArba() {
  const lista = objetivos.filter(o => o.arba);
  const pendientes = lista.filter(o => !yaHecha(o, 'arba'));
  log(`ARBA: ${lista.length} sociedad(es) con CIT, ${pendientes.length} por relevar`);
  if (!pendientes.length) return;
  const browser = await chromium.launch(launchOptions(VER ? { headless: false, slowMo: 150 } : {}));
  try {
    let n = lista.length - pendientes.length;
    for (const o of pendientes) {
      n++;
      log(`[ARBA ${n}/${lista.length}] #${o.pos} ${o.titular} (${o.cuit})`);
      let ultimo = null;
      for (let intento = 1; intento <= 2; intento++) {
        try {
          const r = await correosLib.relevarArba(browser, { cuit: o.cuit, ...o.arba });
          registrar(o, 'arba', { ok: true, ...r });
          log(`   ${r.correos.length ? r.correos.map(c => `${c.correo} [${c.caracter}${c.verificado ? '' : ', sin verificar'}]`).join(', ') : 'SIN CORREO'}` +
              `${r.correoCit ? ` · CIT: ${r.correoCit}` : ''}`);
          ultimo = null;
          break;
        } catch (e) {
          ultimo = e;
          if (navegadorCerrado(browser, e)) throw Object.assign(new Error('Se cerro el navegador. Volve a correrlo y retoma donde quedo.'), { code: 'navegador_cerrado' });
          // Un rechazo de la CIT no se reintenta: ARBA la bloquea.
          if (e.code === 'arba_login_rechazado' || intento === 2) break;
          log(`   fallo (${e.code || 'error'}): ${e.message.split('\n')[0]} — reintento`);
        }
      }
      if (ultimo) {
        const msg = ultimo.message.split('\n')[0];
        log(`   ERROR (${ultimo.code || 'error'}): ${msg}`);
        registrar(o, 'arba', { ok: false, error: msg, code: ultimo.code || 'error' });
      }
    }
  } finally {
    await browser.close().catch(() => {});
  }
}

// ── Analisis ────────────────────────────────────────────────────────────────
const PUBLICOS = /^(gmail|googlemail|hotmail|outlook|live|yahoo|icloud|me|fibertel|speedy|arnet|ciudad)\./i;
const dominio = (c) => String(c || '').split('@')[1]?.toLowerCase() || '';
const clasificar = (c) => {
  const d = dominio(c);
  if (!d) return '';
  if (d === DOMINIO_ESTUDIO || d.endsWith('.' + DOMINIO_ESTUDIO)) return 'Estudio';
  if (PUBLICOS.test(d)) return 'Personal (correo gratuito)';
  return 'Externo';
};

// El dominio recibe correo si tiene MX (o, en su defecto, A: MX implicito).
async function revisarDominios(lista) {
  const out = {};
  for (const d of [...new Set(lista.filter(Boolean))]) {
    try { const mx = await dns.resolveMx(d); out[d] = mx.length ? 'Sí' : 'No'; }
    catch (e) {
      if (['ENOTFOUND', 'ENODATA', 'ESERVFAIL'].includes(e.code)) {
        try { await dns.resolve4(d); out[d] = 'Sin MX (usa A)'; } catch { out[d] = 'No'; }
      } else out[d] = `No se pudo verificar (${e.code})`;
    }
  }
  return out;
}

const minus = (s) => String(s || '').trim().toLowerCase();

// "42242744" y "11-42242744" son el mismo: se deduplica por los ultimos 8
// digitos y queda la primera version (la sabana de ARCA trae el codigo de area).
function unirTelefonos(lista) {
  const vistos = new Map();
  for (const t of lista.filter(Boolean)) {
    const k = String(t).replace(/\D/g, '').slice(-8);
    if (k && !vistos.has(k)) vistos.set(k, t);
  }
  return [...vistos.values()].join(', ');
}

async function consolidar() {
  const filas = todos.map(o => ({ o, s: estado.sociedades[o.cuit] || {} }));
  const todosLosCorreos = [];
  for (const { s } of filas) {
    if (s.arca?.ok) todosLosCorreos.push(s.arca.correo, ...(s.arca.sabana?.correos || []).map(c => c.correo));
    if (s.arba?.ok) todosLosCorreos.push(...s.arba.correos.map(c => c.correo), s.arba.correoCit);
  }
  for (const d of Object.values(estado.dfe)) if (d.ok) for (const l of Object.values(d.mapa)) todosLosCorreos.push(...l.map(x => x.correo));
  const mx = await revisarDominios(todosLosCorreos.filter(Boolean).map(dominio));

  const resumen = [], detalle = [], control = [];
  const unicos = new Map();
  const anotar = (correo, o, org) => {
    const k = minus(correo); if (!k) return;
    const u = unicos.get(k) || { correo: k, sociedades: new Set(), organismos: new Set() };
    u.sociedades.add(`${o.pos} - ${o.titular}`); u.organismos.add(org); unicos.set(k, u);
  };
  const fila = (o, org, fuente, correo, tipo, est) => {
    detalle.push({ pos: o.pos, sociedad: o.titular, cuit: o.cuit, organismo: org, fuente, correo,
                   tipo, estado: est, dominio: dominio(correo), clasificacion: clasificar(correo),
                   recibeCorreo: mx[dominio(correo)] || '' });
    anotar(correo, o, org);
  };

  for (const { o, s } of filas) {
    const obs = [];
    let grave = false;
    const a = s.arca, b = s.arba;

    // ARCA
    let arcaCorreos = [];
    let arcaVerif = '', arcaDfe = '', arcaAvisos = '';
    if (!o.arca) { obs.push('ARCA: sin clave en el archivo de claves'); }
    else if (!a) { obs.push('ARCA: no relevado'); }
    else if (!a.ok) { obs.push(`ARCA: no relevado (${a.error})`); }
    else {
      const sab = a.sabana?.correos || [];
      const vistosA = new Set();
      if (a.correo) { fila(o, 'ARCA', 'Sistema Registral (principal)', a.correo, (sab.find(c => minus(c.correo) === minus(a.correo)) || {}).tipo || '', a.estado || ''); vistosA.add(minus(a.correo)); }
      for (const c of sab) if (!vistosA.has(minus(c.correo))) { fila(o, 'ARCA', 'Sistema Registral (datos registrales)', c.correo, c.tipo, ''); vistosA.add(minus(c.correo)); }
      arcaCorreos = [...vistosA];
      arcaVerif = a.correo ? (a.estado || 'Sin estado') : '';
      arcaDfe = [a.dfe, a.sabana?.dfe && /debe asociar/i.test(a.sabana.dfe) ? 'debe asociar email/teléfono' : ''].filter(Boolean).join(' · ');
      if (!arcaCorreos.length) { obs.push('ARCA: sin correo declarado'); grave = true; }
      else if (a.correo && !/^VERIFICADO$/i.test(a.estado || '')) { obs.push(`ARCA: correo principal no verificado (${a.estado || 'sin estado'})`); grave = true; }
      if (a.dfe && !/^ADHERIDO$/i.test(a.dfe)) { obs.push(`ARCA: DFE "${a.dfe}" (no figura adherido)`); grave = true; }
      if (a.sabana?.dfe && /debe asociar/i.test(a.sabana.dfe)) obs.push('ARCA: el DFE pide asociar email y teléfono al servicio');
      if (a.errorSabana) obs.push(`ARCA: no se pudo leer la lista completa de correos (${a.errorSabana})`);
      const d = estado.dfe[a.login];
      if (d?.ok) {
        const asoc = d.mapa[o.cuit] || [];
        arcaAvisos = asoc.map(x => `${x.correo}${x.tema && x.tema !== 'Todos' ? ` (${x.tema})` : ''}`).join(', ');
        for (const x of asoc) fila(o, 'ARCA', 'DFE - avisos por correo', x.correo, [x.tipoCorreo, x.tema && `tema: ${x.tema}`].filter(Boolean).join(' · '), '');
        if (!asoc.length) obs.push('ARCA: ningún correo asociado a los avisos del DFE (en el login del apoderado)');
      } else if (d && !d.ok) obs.push('ARCA: no se pudo leer los avisos del DFE');
    }

    // ARBA
    let arbaCorreos = [];
    let arbaVerif = '';
    if (!o.arba) { obs.push('ARBA: sin CIT en el archivo de claves'); }
    else if (!b) { obs.push('ARBA: no relevado'); }
    else if (!b.ok) { obs.push(`ARBA: no relevado (${b.error})`); }
    else {
      for (const c of b.correos) fila(o, 'ARBA', 'Datos de contacto', c.correo, c.caracter, c.verificado ? `Verificado ${c.verificado}` : 'No verificado');
      if (b.correoCit) fila(o, 'ARBA', 'Registro de la CIT', b.correoCit, '', '');
      arbaCorreos = b.correos.map(c => minus(c.correo));
      arbaVerif = b.correos.map(c => c.verificado ? `✓ ${c.verificado}` : 'no verificado').join(', ');
      if (!b.correos.length) { obs.push('ARBA: sin correo declarado'); grave = true; }
      else {
        if (!b.correos.some(c => c.verificado)) { obs.push('ARBA: ningún correo verificado'); grave = true; }
        if (!b.correos.some(c => /contribuyente/i.test(c.caracter))) obs.push('ARBA: ningún correo con carácter Contribuyente');
      }
      if (b.correoCit && !arbaCorreos.includes(minus(b.correoCit))) obs.push(`ARBA: el correo de la CIT (${b.correoCit}) no está en los datos de contacto`);
    }

    // Cruce
    if (arcaCorreos.length && arbaCorreos.length && !arcaCorreos.some(c => arbaCorreos.includes(c))) {
      obs.push('Correo distinto en ARCA y en ARBA');
    }
    const propios = new Set([...arcaCorreos, ...arbaCorreos]);
    for (const c of propios) {
      const cl = clasificar(c);
      if (cl !== 'Estudio') obs.push(`Correo ${cl === 'Externo' ? 'externo' : 'personal'}: ${c} — confirmar que siga activo`);
      if (mx[dominio(c)] === 'No') { obs.push(`El dominio de ${c} no recibe correo (sin MX)`); grave = true; }
    }

    // AGIP
    const agip = o.agip ? 'Relevar a mano' : '';
    if (o.agip) obs.push('AGIP: relevar a mano (el sitio bloquea el navegador automatizado)');

    const relevadoAlgo = (a && a.ok) || (b && b.ok);
    const soloAgip = obs.every(x => /^AGIP|sin clave|sin CIT/.test(x));
    resumen.push({
      pos: o.pos, sociedad: o.titular, cuit: o.cuit, tipo: o.tipo || '',
      estado: !relevadoAlgo ? 'Sin relevar' : grave ? 'Revisar' : soloAgip ? 'OK' : 'Ver observaciones',
      arcaCorreos: arcaCorreos.join(', '), arcaVerif, arcaDfe, arcaAvisos,
      arbaCorreos: b?.ok ? b.correos.map(c => `${c.correo} (${c.caracter})`).join(', ') : '',
      arbaVerif, arbaCit: b?.ok ? (b.correoCit || '') : '',
      agip,
      telefonos: unirTelefonos([...(a?.ok ? (a.sabana?.telefonos || []).map(t => t.numero) : []), ...(a?.ok ? a.telefonos || [] : []),
                                ...(b?.ok ? [...b.telefonosFijos, ...b.telefonosMoviles] : [])]),
      observaciones: obs.join('\n'),
    });
    control.push({
      pos: o.pos, sociedad: o.titular, cuit: o.cuit,
      arca: !o.arca ? 'Sin clave' : !a ? 'No relevado' : a.ok ? 'OK' : `Error: ${a.error}`,
      arba: !o.arba ? 'Sin CIT' : !b ? 'No relevado' : b.ok ? 'OK' : `Error: ${b.error}`,
      agip: o.agip ? 'Tiene clave: relevar a mano' : 'Sin clave',
      cuando: [a?.cuando, b?.cuando].filter(Boolean).sort().pop() || '',
    });
  }

  const listaUnicos = [...unicos.values()].map(u => ({
    correo: u.correo, clasificacion: clasificar(u.correo), recibeCorreo: mx[dominio(u.correo)] || '',
    cantidad: u.sociedades.size, organismos: [...u.organismos].join(', '), sociedades: [...u.sociedades].join('\n'),
  })).sort((x, y) => y.cantidad - x.cantidad || x.correo.localeCompare(y.correo));

  const lote = { fecha: iso, dominioEstudio: DOMINIO_ESTUDIO, resumen, detalle, unicos: listaUnicos, control, omitidos };
  const tmp = path.join(DIR, 'lote.json');
  fs.writeFileSync(tmp, JSON.stringify(lote, null, 1), 'utf-8');
  const destino = path.join(DIR, `CORREOS_ORGANISMOS_${iso}.xlsx`);
  try {
    execFileSync(PYTHON, [path.join(__dirname, 'correos_a_excel.py'), tmp, destino],
                 { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] });
  } catch (e) {
    log(`no pude armar el Excel: ${String(e.stderr || e.message).trim().split(/\r?\n/).pop()}`);
  }
  return { lote, destino };
}

(async () => {
  log(`salida: ${DIR}`);
  for (const om of omitidos) log(`omitida #${om.pos} ${om.titular}: ${om.motivo}`);
  if (ORGS.has('arca')) await faseArca();
  if (ORGS.has('arba')) await faseArba();

  const { lote, destino } = await consolidar();
  const r = lote.resumen.filter(x => x.estado !== 'Sin relevar');
  console.log(JSON.stringify({
    ok: true,
    fecha: iso,
    carpeta: DIR,
    excel: destino,
    relevadas: r.length,
    revisar: r.filter(x => x.estado === 'Revisar').map(x => `${x.pos} - ${x.sociedad}`),
    conObservaciones: r.filter(x => x.estado === 'Ver observaciones').length,
    sinObservaciones: r.filter(x => x.estado === 'OK').length,
    correosDistintos: lote.unicos.length,
    errores: lote.control.filter(c => /^Error/.test(c.arca) || /^Error/.test(c.arba))
      .map(c => ({ pos: c.pos, sociedad: c.sociedad, arca: c.arca, arba: c.arba })),
    omitidos,
  }, null, 2));
})().catch(e => fail(e.message, { code: e.code || 'error' }));
