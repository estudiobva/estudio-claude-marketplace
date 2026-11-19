// Helpers de los orquestadores del papel de trabajo: ubicar otros plugins del
// marketplace instalados, correr un script hijo y leer el JSON que imprime.
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const PYTHON = process.env.BVA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');

// Ultima version instalada de un plugin de estudiobva-skills; si no esta en la
// cache, la copia del marketplace. Corta si no hay ninguna.
function pluginDir(nombre) {
  const base = path.join(os.homedir(), '.claude', 'plugins', 'cache', 'estudiobva-skills', nombre);
  if (fs.existsSync(base)) {
    const vers = fs.readdirSync(base).filter((v) => fs.existsSync(path.join(base, v, 'scripts')))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    if (vers.length) return path.join(base, vers[vers.length - 1]);
  }
  const mk = path.join(os.homedir(), '.claude', 'plugins', 'marketplaces', 'estudiobva-skills', nombre);
  if (fs.existsSync(path.join(mk, 'scripts'))) return mk;
  throw Object.assign(new Error(`Falta el plugin ${nombre} de estudiobva-skills (instalarlo con setup-bva o /plugin install ${nombre}@estudiobva-skills)`), { code: 'falta_plugin' });
}

// Ultimo bloque JSON de un texto (los scripts imprimen el JSON al final del stdout).
function ultimoJson(txt) {
  const s = String(txt || '').trim();
  if (!s) return null;
  try { return JSON.parse(s); } catch { /* sigue */ }
  // JSON con indentacion: buscar desde el ultimo "{" o "[" que abre en columna 0.
  const lineas = s.split(/\r?\n/);
  for (let i = lineas.length - 1; i >= 0; i--) {
    if (/^[{[]/.test(lineas[i])) {
      try { return JSON.parse(lineas.slice(i).join('\n')); } catch { /* sigue */ }
    }
  }
  return null;
}

// Corre `cmd args` en cwd; el stderr del hijo va al log con un prefijo.
// Devuelve { code, json, stdout, stderr }.
function correr(cmd, args, { cwd, log = () => {}, prefijo = '', env = {} } = {}) {
  return new Promise((resolve) => {
    const hijo = spawn(cmd, args, { cwd, env: { ...process.env, PYTHONIOENCODING: 'utf-8', ...env }, windowsHide: true });
    let out = '', err = '', resto = '';
    hijo.stdout.on('data', (d) => { out += d; });
    hijo.stderr.on('data', (d) => {
      err += d;
      resto += d;
      const lineas = resto.split(/\r?\n/);
      resto = lineas.pop();
      for (const l of lineas) if (l.trim()) log(`${prefijo}${l}`);
    });
    hijo.on('error', (e) => resolve({ code: -1, json: null, stdout: out, stderr: String(e.message) }));
    hijo.on('close', (code) => {
      if (resto.trim()) log(`${prefijo}${resto}`);
      resolve({ code, json: ultimoJson(out), stdout: out, stderr: err });
    });
  });
}

// "<N> - <Farmacia> - <CUIT>" -> "<Farmacia>"
function nombreFarmacia(socDir, cuit) {
  return path.basename(socDir).replace(/^\s*\d+\s*-\s*/, '').replace(new RegExp(`\\s*-\\s*${cuit}\\s*$`), '').trim();
}

// AAAAMM de MM/AAAA, AAAAMM o AAAA-MM; null si no reconoce.
function normalizarPeriodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + m[1].padStart(2, '0');
  if ((m = s.match(/^(\d{4})(\d{2})$/)) && Number(m[2]) >= 1 && Number(m[2]) <= 12) return s;
  if ((m = s.match(/^(\d{4})[\/\-.](\d{1,2})$/))) return m[1] + m[2].padStart(2, '0');
  return null;
}

function rangoPeriodos(desde, hasta) {
  const out = [];
  let a = Number(desde.slice(0, 4)), m = Number(desde.slice(4));
  const fin = Number(hasta);
  while (a * 100 + m <= fin) {
    out.push(`${a}${String(m).padStart(2, '0')}`);
    if (++m > 12) { m = 1; a++; }
    if (out.length > 60) break;
  }
  return out;
}

const mmaaaa = (per) => `${per.slice(4)}/${per.slice(0, 4)}`;

// Excel abierto en Excel/LibreOffice: deja un "~$<nombre>" (o ".~lock.<nombre>#") al lado.
function abiertoEnExcel(archivo) {
  const dir = path.dirname(archivo), base = path.basename(archivo);
  return [`~$${base}`, `~$${base.slice(2)}`, `.~lock.${base}#`].some((n) => fs.existsSync(path.join(dir, n)));
}

module.exports = { PYTHON, pluginDir, ultimoJson, correr, nombreFarmacia, normalizarPeriodo, rangoPeriodos, mmaaaa, abiertoEnExcel };
