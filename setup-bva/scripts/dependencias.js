#!/usr/bin/env node
/**
 * dependencias.js — instala lo que les falta a los plugins del marketplace
 * estudiobva-skills instalados en esta computadora: node_modules (npm install),
 * el Chromium de Playwright y los paquetes de Python de requirements.txt.
 *
 *   node scripts/dependencias.js              instala solo lo que falta
 *   node scripts/dependencias.js --revisar    no instala nada: informa (exit 1 si falta algo)
 *   node scripts/dependencias.js --forzar     vuelve a correr npm install en todos
 *   node scripts/dependencias.js --plugin=conciliacion-ncr   uno solo
 *   node scripts/dependencias.js --json
 *
 * Hay que correrlo despues de cada actualizacion de un plugin: Claude lo deja en
 * una carpeta nueva, sin node_modules. No toca el .env ni entra a ningun sitio.
 */

const p = require('../lib/plugins');

const args = Object.fromEntries(process.argv.slice(2).map(a => {
  const m = a.match(/^--([^=]+)(?:=(.*))?$/);
  return m ? [m[1], m[2] === undefined ? true : m[2]] : [a, true];
}));
const REVISAR = Boolean(args.revisar || args.check);
const FORZAR = Boolean(args.forzar);
const JSON_OUT = Boolean(args.json);
const log = (...x) => { if (!JSON_OUT) console.log(...x); };

let plugins = p.instalados();
if (args.plugin) plugins = plugins.filter(x => x.nombre === args.plugin);
if (!plugins.length) {
  const msg = args.plugin
    ? `El plugin ${args.plugin} no esta instalado.`
    : `No hay plugins de ${p.MARKET} instalados. Corre primero el instalador (instalar.ps1).`;
  if (JSON_OUT) console.log(JSON.stringify({ ok: false, error: msg }));
  else console.error(msg);
  process.exit(1);
}

const py = p.python();
const chromiumVisto = new Map(); // version de playwright -> instalado
const resultado = [];

for (const pl of plugins) {
  const r = { plugin: pl.nombre, version: pl.version, dir: pl.dir, faltaba: [], hecho: [], errores: [] };
  log(`\n${pl.nombre} ${pl.version}`);

  // npm
  const faltanNode = p.nodeFaltantes(pl.dir);
  if (faltanNode.length) r.faltaba.push(...faltanNode.map(d => `npm:${d}`));
  if ((faltanNode.length || FORZAR) && !REVISAR) {
    log('  npm install...');
    const x = p.correr('npm', ['install', '--no-audit', '--no-fund', '--no-update-notifier', '--loglevel=error'], { cwd: pl.dir, heredar: !JSON_OUT });
    if (x.ok && !p.nodeFaltantes(pl.dir).length) r.hecho.push('npm install');
    else r.errores.push(`npm install fallo${x.err ? ': ' + x.err.split('\n').pop() : ''}`);
  }

  // Chromium, una vez por version de Playwright
  if (p.dependenciasNode(pl.dir).includes('playwright') && p.versionPlaywright(pl.dir)) {
    const v = p.versionPlaywright(pl.dir);
    if (!chromiumVisto.has(v)) chromiumVisto.set(v, p.chromiumInstalado(pl.dir));
    if (!chromiumVisto.get(v)) {
      r.faltaba.push(`chromium (playwright ${v})`);
      if (!REVISAR) {
        log(`  instalando Chromium para Playwright ${v}...`);
        const x = p.correr('npx', ['playwright', 'install', 'chromium'], { cwd: pl.dir, heredar: !JSON_OUT });
        const ok = x.ok && p.chromiumInstalado(pl.dir);
        chromiumVisto.set(v, ok);
        if (ok) r.hecho.push('chromium'); else r.errores.push('no se pudo instalar Chromium (npx playwright install chromium)');
      }
    }
  }

  // Python
  const mods = p.modulosRequeridos(pl.dir);
  const faltanPy = p.modulosFaltantes(py, mods);
  if (faltanPy.length) r.faltaba.push(...faltanPy.map(m => `python:${m}`));
  if (faltanPy.length && !REVISAR) {
    log(`  pip install -r requirements.txt (${faltanPy.join(', ')})...`);
    const x = p.correr(py, ['-m', 'pip', 'install', '--user', '--disable-pip-version-check', '-q', '-r', 'requirements.txt'],
      { cwd: pl.dir, heredar: !JSON_OUT });
    if (x.ok && !p.modulosFaltantes(py, mods).length) r.hecho.push('pip install');
    else r.errores.push(`pip install fallo con ${py}${x.error ? ' (' + x.error.code + ')' : ''}`);
  }

  if (!r.faltaba.length) log('  OK, no falta nada');
  else if (REVISAR) log(`  falta: ${r.faltaba.join(', ')}`);
  for (const e of r.errores) log(`  ERROR ${e}`);
  resultado.push(r);
}

const pendiente = REVISAR ? resultado.some(r => r.faltaba.length) : resultado.some(r => r.errores.length);
if (JSON_OUT) console.log(JSON.stringify({ ok: !pendiente, python: py, plugins: resultado }, null, 2));
else log(pendiente ? (REVISAR ? '\nFaltan dependencias: corre este script sin --revisar.' : '\nQuedaron errores (ver arriba).') : '\nListo.');
process.exit(pendiente ? 1 : 0);
