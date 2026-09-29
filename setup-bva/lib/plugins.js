// plugins.js — lo que comparten dependencias.js, configurar.js, diagnostico.js
// y el chequeo de inicio: donde estan los plugins del marketplace instalados,
// el .env del estudio, el Python y que dependencias le faltan a cada plugin.
//
// No usa dependencias de npm: tiene que andar antes de que haya ninguna.

const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

const MARKET = 'estudiobva-skills';
const REPO = 'estudiobva/estudio-claude-marketplace';
const WIN = process.platform === 'win32';

const claudeDir = () => process.env.CLAUDE_CONFIG_DIR || path.join(os.homedir(), '.claude');

function leerJson(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf-8').replace(/^﻿/, '')); }
  catch { return null; }
}

// Plugins de este marketplace instalados: [{ nombre, dir, version }].
function instalados() {
  const j = leerJson(path.join(claudeDir(), 'plugins', 'installed_plugins.json'));
  const out = [];
  for (const [id, entradas] of Object.entries((j && j.plugins) || {})) {
    const [nombre, market] = id.split('@');
    if (market !== MARKET) continue;
    const lista = Array.isArray(entradas) ? entradas : [entradas];
    const e = lista.find(x => x && x.scope === 'user') || lista[0];
    if (e && e.installPath && fs.existsSync(e.installPath)) {
      out.push({ nombre, dir: e.installPath, version: e.version || '' });
    }
  }
  return out.sort((a, b) => a.nombre.localeCompare(b.nombre));
}

// Nombres de los plugins que publica el marketplace (la copia local que baja Claude).
function publicados() {
  const j = leerJson(path.join(claudeDir(), 'plugins', 'marketplaces', MARKET, '.claude-plugin', 'marketplace.json'));
  return j && Array.isArray(j.plugins) ? j.plugins.map(p => p.name) : null;
}

// ── .env ─────────────────────────────────────────────────────────────────────
// El mismo archivo que leen todos los plugins (lib/env.js de cada uno):
// $FISCO_ENV_FILE o ~/.fisco-ar/.env. Sobrevive a las actualizaciones.

const envFile = () => process.env.FISCO_ENV_FILE || path.join(os.homedir(), '.fisco-ar', '.env');

// Mismo parser que lib/env.js de los plugins, para leer igual que ellos.
function parseEnv(text) {
  const out = {};
  for (const raw of String(text).split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith('#')) continue;
    const m = line.replace(/^export\s+/, '').match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
    if (!m) continue;
    let value = m[2].trim();
    if ((value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    } else {
      value = value.replace(/\s+#.*$/, '').trim();
    }
    out[m[1]] = value;
  }
  return out;
}

function leerEnv() {
  const f = envFile();
  return fs.existsSync(f) ? parseEnv(fs.readFileSync(f, 'utf-8')) : {};
}

// Variable efectiva: la del entorno pisa la del .env (igual que en los plugins).
function variable(nombre, env = leerEnv()) {
  const v = process.env[nombre];
  return v !== undefined && v !== '' ? v : (env[nombre] || '');
}

// ── Procesos ─────────────────────────────────────────────────────────────────

// npm y npx son .cmd en Windows: sin shell no arrancan. Los argumentos que se
// pasan aca son fijos (sin rutas), las rutas van por cwd.
function correr(cmd, args, { cwd, heredar = false, timeout, env } = {}) {
  const r = spawnSync(cmd, args, {
    cwd, timeout, env: env ? { ...process.env, ...env } : process.env, shell: WIN && /^(npm|npx)$/.test(cmd),
    stdio: heredar ? 'inherit' : 'pipe', encoding: 'utf-8', windowsHide: true,
  });
  return { ok: r.status === 0, status: r.status, out: (r.stdout || '').trim(), err: (r.stderr || '').trim(), error: r.error };
}

// ── Python ───────────────────────────────────────────────────────────────────

// El Python que usan los scripts: BVA_PYTHON, o python / python3 como ellos.
function python(env) {
  return variable('BVA_PYTHON', env) || (WIN ? 'python' : 'python3');
}

// Busca un Python real. En Windows "python.exe" de WindowsApps es un atajo a la
// Microsoft Store que no ejecuta nada: se descarta.
function detectarPython() {
  const intentos = WIN ? [['py', ['-3']], ['python', []], ['python3', []]] : [['python3', []], ['python', []]];
  for (const [cmd, pre] of intentos) {
    const r = correr(cmd, [...pre, '-c', 'import sys;print(sys.executable);print("%d.%d" % sys.version_info[:2])'], { timeout: 20000 });
    if (!r.ok) continue;
    const [exe, ver] = r.out.split(/\r?\n/);
    if (!exe || /WindowsApps/i.test(exe)) continue;
    const [ma, mi] = (ver || '0.0').split('.').map(Number);
    if (ma === 3 && mi >= 9) return { exe: exe.trim(), version: ver.trim() };
  }
  return null;
}

// "openpyxl>=3.1" -> "openpyxl". Algunos paquetes se importan con otro nombre.
const MODULO = { 'python-dateutil': 'dateutil', beautifulsoup4: 'bs4', pyyaml: 'yaml', 'pillow': 'PIL' };
function modulosRequeridos(dir) {
  const f = path.join(dir, 'requirements.txt');
  if (!fs.existsSync(f)) return [];
  return fs.readFileSync(f, 'utf-8').split(/\r?\n/)
    .map(l => l.replace(/#.*/, '').trim()).filter(Boolean)
    .map(l => l.split(/[<>=!~;\[\s]/)[0].toLowerCase())
    .filter(Boolean)
    .map(n => MODULO[n] || n.replace(/-/g, '_'));
}

// Modulos que no se pueden importar con ese Python (o todos, si no arranca).
function modulosFaltantes(py, modulos) {
  if (!modulos.length) return [];
  const r = correr(py, ['-c',
    'import importlib.util,sys;print(",".join(m for m in sys.argv[1:] if importlib.util.find_spec(m) is None))',
    ...modulos], { timeout: 30000 });
  if (!r.ok) return modulos;
  return r.out ? r.out.split(',') : [];
}

// ── Node ─────────────────────────────────────────────────────────────────────

function dependenciasNode(dir) {
  const pkg = leerJson(path.join(dir, 'package.json'));
  return pkg ? Object.keys(pkg.dependencies || {}) : [];
}

// Las dependencias sin carpeta en node_modules. Tras una actualizacion el
// plugin queda en una carpeta nueva sin node_modules: es el caso tipico.
function nodeFaltantes(dir) {
  return dependenciasNode(dir)
    .filter(d => !fs.existsSync(path.join(dir, 'node_modules', d, 'package.json')));
}

function versionPlaywright(dir) {
  const p = leerJson(path.join(dir, 'node_modules', 'playwright', 'package.json'));
  return p ? p.version : null;
}

// Chromium de la version de Playwright del plugin (se guarda en una cache
// compartida, ms-playwright, asi que suele estar aunque el plugin sea nuevo).
function chromiumInstalado(dir) {
  const r = correr(process.execPath, ['-e',
    "const p=require('playwright').chromium.executablePath();process.stdout.write(require('fs').existsSync(p)?'si':'no')"],
    { cwd: dir, timeout: 30000 });
  return r.ok && r.out === 'si';
}

module.exports = {
  MARKET, REPO, WIN, claudeDir, leerJson, instalados, publicados,
  envFile, parseEnv, leerEnv, variable, correr,
  python, detectarPython, modulosRequeridos, modulosFaltantes,
  dependenciasNode, nodeFaltantes, versionPlaywright, chromiumInstalado,
};
