#!/usr/bin/env node
/**
 * descargar-ncr.js — deja en la carpeta ncr/ del periodo las Notas de Credito de
 * Recupero (Suizo + Monroe) de una sociedad, listas para la skill conciliacion-ncr.
 *
 * El SharePoint del cliente solo se abre desde el perfil de Chrome "Trabajo"
 * (estudiobva), con la cuenta de Google que autorizaron las farmacias: no hay
 * claves para que un script inicie sesion. Por eso son dos pasos:
 *
 *   1. En esa pestaña del SharePoint se corre scripts/sharepoint/bundle-ncr.js
 *      (con window.__NCR_PERIODO = 'AAAAMM'). Baja SUIZO/ y MONROE/ del mes en
 *      UNA descarga: ~/Downloads/SP-NCR-AAAAMM.bundle. Un paquete sirve para
 *      todas las sociedades del mes.
 *   2. Este script abre el paquete y guarda lo de la sociedad:
 *
 *   node scripts/descargar-ncr.js --nombre="<sociedad>" --periodo=08/2026
 *   node scripts/descargar-ncr.js --cuit=<CUIT> --periodo=202608 --bundle="C:\...\SP-NCR-202608.bundle"
 *   node scripts/descargar-ncr.js --nombre="<sociedad>" --periodo=08/2026 --actualizar
 *   node scripts/descargar-ncr.js --cuit=<CUIT> --periodo=202608 --solo-revisar   (no escribe nada)
 *
 * Como elige los archivos de Monroe: por el CUIT que trae cada archivo en la
 * columna "Cuit", no por el nombre (los nombres cambian de un mes a otro:
 * "F ALEMANA" en julio, "FRANCO ALEMANA " en agosto).
 *
 * El nombre de cada archivo de Monroe se usa para resolver el CLIENTE de Suizo
 * (igual o via lib/ncr-alias.json). Lo que no resuelva se informa y el script
 * termina con error: no se adivina por parecido.
 *
 * Nunca pisa: si el archivo ya esta en ncr/ con el mismo contenido lo deja; si
 * cambio, avisa y no lo toca. Con --actualizar lo reemplaza y guarda el anterior
 * en ncr/_anteriores/.
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');

const { parseArgs, fail } = require('../lib/args');
const { resolveCredentials, envFile } = require('../lib/env');
const { readWorkbook } = require('../lib/xlsx-min');
const rutas = require('../lib/rutas-bva');

const args = parseArgs();
const log = (m) => process.stderr.write(`[descargar-ncr] ${m}\n`);

const norm = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .toUpperCase().replace(/\.XLSX?$/i, '').replace(/\s+/g, ' ').trim();
const digitos = (s) => String(s || '').replace(/\D/g, '');

function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)))          return s;
  return null;
}

// ── lectura de los archivos (en memoria, con el lector minimo del proyecto) ──
function tabla(buf) {
  const filas = readWorkbook(buf).rows();
  return filas.filter((f) => f && f.some((c) => String(c).trim() !== ''));
}

// Monroe: encabezado con "Tipo Linea"; filas "Detalle"; pie "Cantidad de Registros: N".
function leerMonroe(buf) {
  const t = tabla(buf);
  const cab = t[0].map((c) => String(c).trim());
  const col = (n) => cab.indexOf(n);
  if (col('Tipo Linea') < 0 || col('Cuit') < 0 || col('Numero Formateado') < 0) return null;
  const det = t.slice(1).filter((f) => f[col('Tipo Linea')] === 'Detalle');
  const pie = t.find((f) => /Cantidad de Registros/i.test(f[0] || ''));
  return {
    cuits: [...new Set(det.map((f) => digitos(f[col('Cuit')])))],
    claves: det.map((f) => `${f[col('Numero Formateado')]}|${Number(f[col('Importe Total')]).toFixed(2)}`).sort(),
    leidas: det.length,
    declaradas: pie ? Number((pie[0].match(/(\d+)/) || [])[1]) : null,
  };
}

// Suizo: encabezado con CLIENTE, terminal, numero, total.
function leerSuizo(buf) {
  const t = tabla(buf);
  const i = t.findIndex((f) => f.some((c) => String(c).trim().toUpperCase() === 'CLIENTE'));
  if (i < 0) return null;
  const cab = t[i].map((c) => String(c).trim().toLowerCase());
  const col = (n) => cab.indexOf(n);
  const filas = t.slice(i + 1).filter((f) => f[col('cliente')]);
  return {
    filas,
    clientes: new Map([...new Set(filas.map((f) => norm(f[col('cliente')])))]
      .map((c) => [c, filas.filter((f) => norm(f[col('cliente')]) === c).length])),
    claves: (clientes) => filas.filter((f) => clientes.includes(norm(f[col('cliente')])))
      .map((f) => `${Number(f[col('terminal')])}|${Number(f[col('numero')])}|${Number(f[col('total')]).toFixed(2)}`).sort(),
  };
}

// Paquete armado por scripts/sharepoint/bundle-ncr.js.
function abrirBundle(ruta) {
  const b = fs.readFileSync(ruta);
  const n = b.readUInt32BE(0);
  const indice = JSON.parse(b.subarray(4, 4 + n).toString('utf-8'));
  let pos = 4 + n;
  for (const a of indice.archivos) { a.buf = b.subarray(pos, pos + a.bytes); pos += a.bytes; }
  if (pos !== b.length) throw new Error(`El paquete ${ruta} esta incompleto o dañado.`);
  return indice;
}

// El mas nuevo de ~/Downloads (Chrome agrega " (1)", " (2)" si se baja otra vez).
function ubicarBundle(periodo) {
  if (args['bundle']) return path.resolve(String(args['bundle']));
  const dir = path.join(os.homedir(), 'Downloads');
  const re = new RegExp(`^SP-NCR-${periodo}( \\(\\d+\\))?\\.bundle$`, 'i');
  const hits = fs.existsSync(dir)
    ? fs.readdirSync(dir).filter((f) => re.test(f))
        .map((f) => ({ f, t: fs.statSync(path.join(dir, f)).mtimeMs })).sort((x, y) => y.t - x.t)
    : [];
  if (!hits.length) {
    throw new Error(`No encuentro SP-NCR-${periodo}.bundle en ${dir}. Primero hay que correr ` +
      'scripts/sharepoint/bundle-ncr.js en la pestaña del SharePoint (perfil de Chrome "Trabajo").');
  }
  return path.join(dir, hits[0].f);
}

// Guarda sin pisar. `mismo(bufViejo)` decide si el contenido es equivalente
// (el Drive re-guarda los .xlsx y cambian los bytes aunque el contenido no).
function guardar(dir, nombre, buf, mismo, actualizar, soloRevisar) {
  const destino = path.join(dir, nombre);
  if (!fs.existsSync(destino)) { if (!soloRevisar) fs.writeFileSync(destino, buf); return soloRevisar ? 'faltaria' : 'nuevo'; }
  let igual = false;
  try { igual = mismo(fs.readFileSync(destino)); } catch { igual = false; }
  if (igual) return 'sin_cambios';
  if (!actualizar || soloRevisar) return 'distinto_no_tocado';
  const ant = path.join(dir, '_anteriores');
  fs.mkdirSync(ant, { recursive: true });
  const d = new Date(), z = (n) => String(n).padStart(2, '0');
  const sello = `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`;
  fs.renameSync(destino, path.join(ant, nombre.replace(/(\.xlsx?)$/i, ` (${sello})$1`)));
  fs.writeFileSync(destino, buf);
  return 'reemplazado';
}

(() => {
  const PERIODO = normalizarPeriodo(args['periodo']);
  if (!PERIODO) fail('Falta --periodo (MM/AAAA o AAAAMM)');
  const ACTUALIZAR = 'actualizar' in args;
  const SOLO_REVISAR = 'solo-revisar' in args;   // no escribe nada: solo informa

  let CUIT, TITULAR;
  try {
    ({ cuit: CUIT, titular: TITULAR } = args['cuit'] && !args['nombre']
      ? { cuit: digitos(args['cuit']), titular: null }
      : resolveCredentials({ cuit: args['cuit'], nombre: args['nombre'] }));
  } catch (e) {
    fail(e.message, { sugerencia: `Revisa el .env (${envFile() || 'no encontrado'}) o el Excel de claves.` });
  }
  const alias = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'lib', 'ncr-alias.json'), 'utf-8')).alias;

  try {
    let DESTINO;
    try { DESTINO = rutas.carpetaNcr(CUIT, PERIODO, { crear: !SOLO_REVISAR }); }
    catch (e) { if (!SOLO_REVISAR) throw e; DESTINO = `(todavia no existe: ${e.message})`; }
    const BUNDLE = ubicarBundle(PERIODO);
    const pkg = abrirBundle(BUNDLE);
    if (pkg.periodo !== PERIODO) throw new Error(`El paquete es del periodo ${pkg.periodo}, no de ${PERIODO}.`);
    const horas = (Date.now() - Date.parse(pkg.generado)) / 36e5;
    log(`${TITULAR || CUIT} · CUIT ${CUIT} · paquete ${path.basename(BUNDLE)} (armado hace ${horas.toFixed(1)} h) → ${DESTINO}`);
    if (horas > 24) log('⚠ el paquete tiene mas de un dia: si las farmacias subieron cambios, volver a armarlo.');

    // ── Monroe: quedarse con los archivos del CUIT ──
    const listaMonroe = pkg.archivos.filter((a) => a.drogueria === 'MONROE' && /\.xlsx$/i.test(a.nombre));
    const propios = [], sinFilas = [], mixtos = [], ilegibles = [];
    for (const a of listaMonroe) {
      let m;
      try { m = leerMonroe(a.buf); } catch { m = null; }
      if (!m) { ilegibles.push(a.nombre); continue; }
      if (!m.cuits.length) { sinFilas.push(a.nombre); continue; }
      if (m.cuits.length > 1) { if (m.cuits.includes(CUIT)) mixtos.push(`${a.nombre} (${m.cuits.join(', ')})`); continue; }
      if (m.cuits[0] === CUIT) propios.push({ ...a, m, farmacia: norm(a.nombre) });
    }
    log(`Monroe: ${listaMonroe.length} archivos en el paquete, ${propios.length} con el CUIT ${CUIT}`);
    if (mixtos.length) throw new Error(`Archivos de Monroe con mas de un CUIT adentro: ${mixtos.join(' | ')}. Revisar a mano.`);
    if (!propios.length) {
      throw Object.assign(new Error(`Ningun archivo de Monroe de ${pkg.origen} tiene el CUIT ${CUIT}.`), { code: 'sin_archivos_monroe' });
    }
    const repetidas = propios.map((p) => p.farmacia).filter((f, i, v) => v.indexOf(f) !== i);
    if (repetidas.length) throw new Error(`Dos archivos de Monroe dan el mismo nombre de farmacia: ${repetidas.join(', ')}.`);

    // ── Suizo: un solo archivo ──
    const listaSuizo = pkg.archivos.filter((a) => a.drogueria === 'SUIZO' && /\.xlsx?$/i.test(a.nombre));
    if (listaSuizo.length !== 1) {
      throw new Error(`En SUIZO hay ${listaSuizo.length} archivos (${listaSuizo.map((a) => a.nombre).join(', ')}); se espera uno.`);
    }
    if (/\.xls$/i.test(listaSuizo[0].nombre)) {
      throw new Error(`El archivo de Suizo es .xls viejo (${listaSuizo[0].nombre}); hay que abrirlo y guardarlo como .xlsx.`);
    }
    const bufSuizo = listaSuizo[0].buf;
    const suizo = leerSuizo(bufSuizo);
    if (!suizo) throw new Error(`El archivo de Suizo (${listaSuizo[0].nombre}) no tiene la columna CLIENTE.`);

    // ── CLIENTE de Suizo para cada farmacia ──
    const resueltos = [], sinResolver = [];
    for (const p of propios) {
      const cliente = alias[p.farmacia] || p.farmacia;
      if (suizo.clientes.has(cliente)) {
        resueltos.push({ farmacia: p.farmacia, cliente, via: alias[p.farmacia] ? 'alias' : 'igual',
                         ncr_suizo: suizo.clientes.get(cliente) });
      } else sinResolver.push(p.farmacia);
    }
    const clientes = resueltos.map((r) => r.cliente);

    // ── guardar ──
    if (!SOLO_REVISAR) fs.mkdirSync(DESTINO, { recursive: true });
    const guardados = [];
    for (const p of propios) {
      const estado = guardar(DESTINO, `${p.farmacia}.xlsx`, p.buf, (viejo) => {
        const v = leerMonroe(viejo); return v && v.claves.join('\n') === p.m.claves.join('\n');
      }, ACTUALIZAR, SOLO_REVISAR);
      guardados.push({ archivo: `${p.farmacia}.xlsx`, origen: p.nombre, estado, leidas: p.m.leidas,
                       declaradas: p.m.declaradas, pie_ok: p.m.declaradas == null || p.m.declaradas === p.m.leidas });
    }
    const nombreSuizo = `${PERIODO}-SUIZO NCR.xlsx`;
    const estadoSuizo = guardar(DESTINO, nombreSuizo, bufSuizo, (viejo) => {
      const v = leerSuizo(viejo);
      return v && v.filas.length === suizo.filas.length && v.claves(clientes).join('\n') === suizo.claves(clientes).join('\n');
    }, ACTUALIZAR, SOLO_REVISAR);
    guardados.push({ archivo: nombreSuizo, origen: listaSuizo[0].nombre, estado: estadoSuizo,
                     ncr_de_la_sociedad: clientes.reduce((s, c) => s + suizo.clientes.get(c), 0), filas_grupo: suizo.filas.length });

    for (const g of guardados) log(`${g.archivo}: ${g.estado}${g.pie_ok === false ? `  ⚠ pie ${g.leidas}/${g.declaradas}` : ''}`);
    const distintos = guardados.filter((g) => g.estado === 'distinto_no_tocado');
    const truncados = guardados.filter((g) => g.pie_ok === false);

    const res = {
      ok: !sinResolver.length && !distintos.length,
      titular: TITULAR, cuit: CUIT, periodo: PERIODO, origen: pkg.origen,
      paquete: BUNDLE, paquete_armado: pkg.generado, destino: DESTINO, archivos: guardados,
      clientes_suizo: clientes.join(','),   // listo para el ultimo argumento de reconciliar_ncr.py
      resolucion_suizo: resueltos,
      sin_resolver_en_suizo: sinResolver,
      exportaciones_monroe_truncadas: truncados.map((g) => `${g.archivo}: ${g.leidas}/${g.declaradas}`),
      distintos_no_tocados: distintos.map((g) => g.archivo),
      monroe_sin_filas: sinFilas, monroe_ilegibles: ilegibles,
    };
    if (sinResolver.length) {
      res.que_hacer = `Estas farmacias de Monroe no aparecen como CLIENTE en Suizo: ${sinResolver.join(', ')}. ` +
        'Confirmar a que CLIENTE corresponden y agregarlas a lib/ncr-alias.json (o no tienen NCR de Suizo este mes).';
    }
    if (distintos.length) {
      res.que_hacer = (res.que_hacer ? res.que_hacer + ' ' : '') +
        'Hay archivos que cambiaron en el SharePoint respecto de los que ya estan en ncr/. Revisar y volver a correr con --actualizar.';
    }
    console.log(JSON.stringify(res, null, 2));
    if (!res.ok) process.exitCode = 1;
  } catch (e) {
    fail(e.message, { code: e.code || 'error' });
  }
})();
