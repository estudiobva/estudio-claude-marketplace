#!/usr/bin/env node
/**
 * archivar-descargas.js — Renombra y archiva extractos bajados A MANO o por
 * Claude in Chrome (bancos donde el script no puede entrar: captcha, antibots).
 *
 * Toma los archivos de la carpeta de descargas del Chrome de la persona (por
 * defecto ~/Downloads) y los guarda igual que extractos-banco.js: mismo nombre
 * ([sociedad]_[banco]_AAAA-MM_mensual.pdf), misma carpeta, sin pisar nada, sin
 * duplicar (compara contenido) y rechazando "PDF" que no son PDF. No borra ni
 * mueve los originales de Descargas.
 *
 *   # los bajados en los ultimos 30 minutos (o desde una hora)
 *   node scripts/archivar-descargas.js --banco=COMAFI --nombre="Sociedad Ejemplo" --recientes=30
 *   node scripts/archivar-descargas.js --banco=COMAFI --cuit=30000000007 --desde-hora=10:15
 *   # uno puntual, con el periodo leido del PDF cuando el nombre no lo trae
 *   node scripts/archivar-descargas.js --banco=COMAFI --cuit=30000000007 --archivo="~/Downloads/Resumen.pdf" --periodo=202610
 *
 * Opciones: --carpeta (default ~/Downloads), --unidad (archiva en
 * 03-Bancos y Conciliaciones/[BANCO]/[AAAA]/ de la sociedad), --salida, --semana=N (Galicia).
 *
 * Los archivos cuyo periodo no se puede leer del nombre quedan en `sin_periodo`:
 * hay que abrirlos, ver el periodo del resumen y volver a correr con
 * --archivo=... --periodo=AAAAMM. Salida: JSON por stdout.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const ex = require('../lib/extractos');
const bancos = require('../lib/bancos');

loadEnv();
const args = parseArgs();
const casa = (p) => String(p).replace(/^~(?=$|\/)/, os.homedir());
const EXT = /\.(pdf|xlsx?|csv|txt)$/i;

function desdeCuando() {
  if (args['recientes']) return Date.now() - Number(args['recientes']) * 60000;
  if (args['desde-hora']) {
    const m = String(args['desde-hora']).match(/^(\d{1,2}):(\d{2})$/);
    if (!m) throw new Error('--desde-hora va como HH:MM (hora de hoy).');
    const d = new Date(); d.setHours(+m[1], +m[2], 0, 0);
    return d.getTime();
  }
  return null;
}

(async () => {
  if (!args['banco']) throw new Error('Falta --banco.');
  const banco = bancos.canonico(args['banco']);
  const adaptador = bancos.adaptador(banco);
  if (!args['cuit'] && !args['nombre']) throw new Error('Falta --cuit o --nombre de la sociedad.');
  const soc = ex.sociedad({ cuit: args['cuit'], nombre: args['nombre'] });

  // Archivos: los --archivo= pasados (se pueden repetir) o los recientes de la carpeta.
  let archivos = process.argv.slice(2).filter(a => a.startsWith('--archivo=')).map(a => casa(a.slice(10)));
  if (!archivos.length) {
    const desde = desdeCuando();
    if (!desde) throw new Error('Pasa --archivo=<ruta>, --recientes=<minutos> o --desde-hora=HH:MM.');
    const carpeta = casa(args['carpeta'] || path.join(os.homedir(), 'Downloads'));
    archivos = fs.readdirSync(carpeta)
      .filter(f => EXT.test(f) && !f.endsWith('.crdownload'))
      .map(f => path.join(carpeta, f))
      .filter(f => fs.statSync(f).isFile() && fs.statSync(f).mtimeMs >= desde)
      .sort((a, b) => fs.statSync(a).mtimeMs - fs.statSync(b).mtimeMs);
  }
  if (args['periodo'] && archivos.length !== 1) throw new Error('--periodo va con un solo --archivo.');

  const cb = 'unidad' in args
    ? ex.carpetaBanco(soc, banco, args['carpeta-banco'] ? { elegida: args['carpeta-banco'] } : {})
    : ex.carpetaSalida(ex.salidaBase(args['salida']), soc, banco);
  if (cb.candidatas) throw new Error(`Hay varias carpetas para ${banco}: ${cb.candidatas.join(' | ')}. Pasa --carpeta-banco=.`);

  const r = { banco, sociedad: soc.nombre, cuit: soc.cuit, destino: cb.dir,
              guardados: [], ya_estaban: [], sin_periodo: [], rechazados: [] };
  for (const f of archivos) {
    if (!fs.existsSync(f)) { r.rechazados.push({ archivo: f, motivo: 'no existe' }); continue; }
    let det = args['periodo'] ? { periodo: ex.periodo(args['periodo']), dia: null } : null;
    if (args['periodo'] && !det.periodo) throw new Error('--periodo invalido (AAAAMM).');
    if (!det) {
      det = adaptador.periodoDeArchivo ? adaptador.periodoDeArchivo(path.basename(f)) : undefined;
      if (det === undefined) det = ex.periodoDeArchivo(f);
    }
    if (!det || !det.periodo) { r.sin_periodo.push(det && det.ambiguo ? { archivo: f, nota: det.ambiguo } : { archivo: f }); continue; }
    const semana = args['semana'] ? Number(args['semana'])
      : (banco === 'GALICIA' && det.dia ? Math.ceil(det.dia / 7) : null);
    const nombre = ex.nombreArchivo({ soc, banco, periodo: det.periodo, semana, ext: path.extname(f).toLowerCase() });
    try {
      const res = ex.archivar({ origen: f, dirBanco: cb.dir, periodo: det.periodo, nombre });
      (res.estado === 'guardado' ? r.guardados : r.ya_estaban).push({ archivo: path.basename(f), periodo: det.periodo, destino: res.destino });
    } catch (e) {
      r.rechazados.push({ archivo: f, motivo: e.message });
    }
  }
  if (r.sin_periodo.length) r.que_hacer = 'Abrir cada archivo de sin_periodo, ver el periodo del resumen y correr con --archivo="<ruta>" --periodo=AAAAMM.';
  console.log(JSON.stringify({ ok: true, ...r }, null, 2));
})().catch((e) => fail(e.message, { code: e.code || 'error' }));
