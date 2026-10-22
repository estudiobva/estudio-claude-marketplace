// sociedades.js — empareja el nombre con que un banco muestra a cada empresa
// ("SOY FARMACIA EJEMPLO 9 S.C.S.") con la sociedad del estudio (carpeta de la unidad
// + CUIT), usando la hoja MAESTRO del Maestro de Sociedades.
//
// Orden: 1) alias explicito (~/.fisco-ar/prestamos-alias.json o
// lib/prestamos/alias.json: { "NOMBRE EN EL BANCO": "CUIT" }); 2) nombre igual
// al del Maestro o al de la carpeta, sin la forma societaria; 3) el mas
// parecido por palabras, solo si gana claro. Si no, queda sin sociedad (se
// informa y se guarda aparte): nunca se adivina entre dos.

const fs = require('fs');
const os = require('os');
const path = require('path');
const rutas = require('../rutas-bva');
const { readWorkbook } = require('../xlsx-min');

const SOCIETARIO = new Set(['S', 'C', 'A', 'R', 'L', 'SA', 'SRL', 'SCS', 'SAS', 'SOC', 'SOCIEDAD', 'EN',
  'COMANDITA', 'SIMPLE', 'ANONIMA', 'RESPONSABILIDAD', 'LIMITADA', 'DE', 'DEL', 'LA', 'LAS', 'Y']);
// Palabras que casi todas las sociedades comparten: no sirven para distinguir.
const COMUNES = new Set(['SOY', 'MAGA', 'FARMACIA', 'SANAR']);

const NUMEROS = { UNO: '1', DOS: '2', TRES: '3', CUATRO: '4', CINCO: '5', SEIS: '6', SIETE: '7',
  OCHO: '8', NUEVE: '9', DIEZ: '10', ONCE: '11', DOCE: '12' };
// "CASERO" / "CASEROS" y "TRES" / "3" se comparan igual.
const raizPalabra = (t) => NUMEROS[t] || (t.length > 4 ? t.replace(/S$/, '') : t);
const tokens = (s) => String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/\([^)]*\)/g, ' ').replace(/[^A-Z0-9]+/g, ' ').trim().split(/\s+/)
  .filter(t => t && !SOCIETARIO.has(t)).map(raizPalabra);
const clave = (s) => tokens(s).join(' ');
const sinComunes = (s) => [...new Set(tokens(s).filter(t => !COMUNES.has(t)))];

let _cache = null;
function sociedades() {
  if (_cache) return _cache;
  const base = rutas.raiz();
  const carpetas = new Map();
  for (const d of fs.readdirSync(base, { withFileTypes: true })) {
    const m = d.isDirectory() && d.name.match(/^\s*\d+\s*-\s*(.+?)\s*-\s*(\d{11})\s*$/);
    if (m && m[2] !== '00000000000') carpetas.set(m[2], { nombre: m[1], cuit: m[2], dir: path.join(base, d.name) });
  }
  const nombres = [];   // { nombre, cuit }
  for (const s of carpetas.values()) nombres.push({ nombre: s.nombre, cuit: s.cuit });
  try {
    const archivo = path.join(base, '0 - ESTUDIO - 00000000000', 'Maestro de Sociedades.xlsx');
    const filas = readWorkbook(fs.readFileSync(archivo)).rows('MAESTRO');
    const iEnc = filas.findIndex(f => f.some(c => /^CUIT$/i.test(String(c).trim())));
    const enc = filas[iEnc].map(c => String(c).trim());
    const cN = enc.findIndex(c => /^NOMBRE$/i.test(c)), cC = enc.findIndex(c => /^CUIT$/i.test(c));
    for (const f of filas.slice(iEnc + 1)) {
      const cuit = String(f[cC] || '').replace(/\D/g, '');
      if (carpetas.has(cuit) && String(f[cN] || '').trim()) nombres.push({ nombre: String(f[cN]).trim(), cuit });
    }
  } catch { /* sin Maestro: solo nombres de carpeta */ }
  _cache = { carpetas, nombres };
  return _cache;
}

function alias() {
  const out = {};
  for (const f of [path.join(__dirname, 'alias.json'), path.join(os.homedir(), '.fisco-ar', 'prestamos-alias.json')]) {
    try { Object.assign(out, JSON.parse(fs.readFileSync(f, 'utf8'))); } catch { /* opcional */ }
  }
  const norm = {};
  for (const [k, v] of Object.entries(out)) norm[clave(k)] = String(v).replace(/\D/g, '');
  return norm;
}

// Devuelve { nombre, cuit, dir, como } o null.
function emparejar(nombreBanco, { cuit: cuitVisto } = {}) {
  const { carpetas, nombres } = sociedades();
  const res = (cuit, como) => (carpetas.has(cuit) ? { ...carpetas.get(cuit), como } : null);
  if (cuitVisto && carpetas.has(String(cuitVisto))) return res(String(cuitVisto), 'cuit');
  const a = alias()[clave(nombreBanco)];
  if (a) return res(a, 'alias');

  const k = clave(nombreBanco);
  const exactos = [...new Set(nombres.filter(n => clave(n.nombre) === k).map(n => n.cuit))];
  if (exactos.length === 1) return res(exactos[0], 'nombre');

  // Por palabras distintivas: todas las de la sociedad tienen que estar en el
  // nombre del banco (o al reves), y gana solo si no empata con otra.
  const tb = new Set(sinComunes(nombreBanco));
  if (!tb.size) return null;
  const puntajes = new Map();
  for (const n of nombres) {
    const ts = sinComunes(n.nombre);
    if (!ts.length) continue;
    const comunes = ts.filter(t => tb.has(t)).length;
    const cubre = comunes === ts.length || comunes === tb.size;
    if (!cubre || !comunes) continue;
    const p = comunes / Math.max(ts.length, tb.size);
    if (p > (puntajes.get(n.cuit) || 0)) puntajes.set(n.cuit, p);
  }
  const orden = [...puntajes.entries()].sort((x, y) => y[1] - x[1]);
  if (orden.length && (orden.length === 1 || orden[0][1] > orden[1][1])) return res(orden[0][0], 'parecido');
  return null;
}

module.exports = { emparejar, sociedades, clave };
