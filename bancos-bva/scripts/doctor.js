#!/usr/bin/env node
/**
 * doctor.js — diagnostico del plugin bancos-bva. No entra a ningun banco.
 *
 *   node scripts/doctor.js          reporte legible
 *   node scripts/doctor.js --json   para parsear
 *
 * Chequea: Node, dependencias, que Chromium abra de verdad, que .env se cargo,
 * la unidad compartida, el Maestro de Sociedades y que Claves_Bancos.xlsx abra
 * con la contrasena (informa que bancos tienen datos cargados, sin mostrarlos).
 * Exit 1 si hay algun ❌.
 */

const fs = require('fs');
const path = require('path');
const { parseArgs } = require('../lib/args');
const { loadEnv, envFile } = require('../lib/env');

loadEnv();
const JSON_OUT = 'json' in parseArgs();
const checks = [];
const add = (id, estado, detalle, arreglo) => checks.push({ id, estado, detalle, ...(arreglo ? { arreglo } : {}) });

(async () => {
  // Node
  const major = Number(process.versions.node.split('.')[0]);
  add('node', major >= 20 ? 'ok' : 'error', `v${process.versions.node}`, major >= 20 ? null : 'Instala Node 20 o superior.');

  // Dependencias
  let pw = null;
  try { pw = require('playwright'); add('playwright', 'ok', 'instalado'); }
  catch { add('playwright', 'error', 'no instalado', 'cd a la raiz del plugin y corre: npm install'); }
  try { require('officecrypto-tool'); add('officecrypto-tool', 'ok', 'instalado'); }
  catch { add('officecrypto-tool', 'error', 'no instalado (hace falta para abrir Claves_Bancos.xlsx)', 'npm install'); }

  // Chromium (que abra de verdad) y Chrome instalado
  if (pw) {
    try { const b = await pw.chromium.launch({ headless: true }); await b.close(); add('chromium', 'ok', 'abre'); }
    catch (e) { add('chromium', 'error', e.message.split('\n')[0], 'npx playwright install chromium'); }
    try { const b = await pw.chromium.launch({ headless: true, channel: 'chrome' }); await b.close(); add('chrome', 'ok', 'Google Chrome disponible (se usa por defecto)'); }
    catch { add('chrome', 'aviso', 'Google Chrome no disponible: se usa Chromium (algunos bancos lo miran con mas desconfianza)'); }
  }

  // .env
  const env = envFile();
  add('env', env ? 'ok' : 'aviso', env || 'no se encontro ningun .env',
      env ? null : 'Crea ~/.fisco-ar/.env a partir de .env.example.');

  // Unidad compartida
  const rutas = require('../lib/rutas-bva');
  let raiz = null;
  try { raiz = rutas.raiz(); add('unidad', 'ok', raiz); }
  catch (e) { add('unidad', 'error', e.message, 'Monta Google Drive o define BVA_UNIDAD_PATH.'); }

  // Maestro de Sociedades
  if (raiz) {
    try {
      const ex = require('../lib/extractos');
      const { readWorkbook } = require('../lib/xlsx-min');
      const filas = readWorkbook(fs.readFileSync(path.join(raiz, '0 - ESTUDIO - 00000000000', 'Maestro de Sociedades.xlsx'))).rows('MAESTRO');
      const cuits = filas.map(f => String(f[1] || '').replace(/\D/g, '')).filter(c => c.length === 11);
      let conBancos = 0;
      for (const c of cuits) { try { if (ex.bancosDelMaestro(c).length) conBancos++; } catch { /* fila rara */ } }
      add('maestro', 'ok', `${cuits.length} entidades, ${conBancos} con bancos marcados`);
    } catch (e) { add('maestro', 'error', e.message.split('\n')[0], 'Revisa Maestro de Sociedades.xlsx (hoja MAESTRO).'); }
  }

  // Claves_Bancos.xlsx
  const cb = require('../lib/claves-bancos');
  const archivo = cb.filePath();
  if (!fs.existsSync(archivo)) {
    add('claves-bancos', 'error', `no existe ${archivo}`, 'Define CLAVES_BANCOS_PATH en el .env.');
  } else {
    try {
      const { bancos: filas, galicia, avisoGalicia } = await cb.load();
      const completos = filas.filter(b => b.usuario && b.clave).map(b => b.banco);
      const incompletos = filas.filter(b => !b.usuario || !b.clave).map(b => b.banco);
      add('claves-bancos', incompletos.length ? 'aviso' : 'ok',
        `abre OK · login unico con datos: ${completos.join(', ') || '(ninguno)'}` +
        (incompletos.length ? ` · sin usuario o clave: ${incompletos.join(', ')}` : ''),
        incompletos.length ? 'Completa esas filas en Claves_Bancos.xlsx.' : null);
      const g = galicia.filter(x => x.usuario && x.clave).length;
      add('claves-galicia', avisoGalicia ? 'error' : 'ok',
        avisoGalicia || `${g} de ${galicia.length} sociedades con login propio de Galicia`);
    } catch (e) {
      add('claves-bancos', 'error', e.message.split('\n')[0],
          'Revisa CLAVES_BANCOS_PASSWORD (o CLAVES_ORGANISMOS_PASSWORD) en el .env.');
    }
  }

  // Adaptadores
  const bancos = require('../lib/bancos');
  const auto = bancos.bancos.filter(b => typeof bancos.adaptador(b).descargar === 'function');
  add('adaptadores', 'ok', `login: ${bancos.bancos.join(', ')} · descarga automatica: ${auto.join(', ') || 'ninguno todavia (modo asistido)'}`);

  if (JSON_OUT) {
    console.log(JSON.stringify({ ok: !checks.some(c => c.estado === 'error'), checks }, null, 2));
  } else {
    const ico = { ok: '✅', aviso: '⚠️ ', error: '❌' };
    for (const c of checks) {
      console.log(`${ico[c.estado]} ${c.id.padEnd(18)} ${c.detalle}`);
      if (c.arreglo) console.log(`   ${' '.repeat(18)} → ${c.arreglo}`);
    }
  }
  process.exit(checks.some(c => c.estado === 'error') ? 1 : 0);
})();
