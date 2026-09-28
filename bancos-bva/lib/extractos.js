// extractos.js — donde y con que nombre se archiva cada extracto bancario.
//
// Destino (como el resto de los scripts del estudio):
//   ~/Documents/BVA-salidas/extractos/[N] - [Sociedad] - [CUIT]/[BANCO]/[AAAA]/
//   (--salida o $BVA_SALIDAS_PATH/extractos). Con --unidad, directo en la unidad:
//   [N] - [Sociedad] - [CUIT]/03-Bancos y Conciliaciones/[BANCO]/[AAAA]/
//   En los dos casos van todos los meses del ano juntos: el mes ya va en el nombre.
//
// Nombre:
//   [sociedad]_[banco]_[AAAA-MM]_mensual.pdf
//   [sociedad]_[banco]_[AAAA-MM]_semana[N].pdf     (Galicia entrega semanales)
//   Si ya hay otro archivo DISTINTO con ese nombre: ..._2.pdf, _3.pdf...
//   Si ya hay uno IGUAL (mismo contenido): no se copia de nuevo.
//
// Nunca pisa ni borra nada de la unidad.

const fs     = require('fs');
const path   = require('path');
const crypto = require('crypto');
const rutas  = require('./rutas-bva');
const { readWorkbook } = require('./xlsx-min');
const { canonico, norm } = require('./bancos');

const BANCOS_DIR = '03-Bancos y Conciliaciones';

// ── periodos ────────────────────────────────────────────────────────────────
function periodo(p) {
  const s = String(p || '').trim();
  let m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{4})$/))) return m[2] + String(m[1]).padStart(2, '0');
  if ((m = s.match(/^(\d{4})[\/\-.]?(\d{2})$/)) && +m[2] >= 1 && +m[2] <= 12) return m[1] + m[2];
  return null;
}

function rango(desde, hasta) {
  const a = periodo(desde), b = periodo(hasta || desde);
  if (!a || !b) throw new Error('Periodo invalido. Usa AAAAMM o MM/AAAA (ej. --desde=202601 --hasta=202606).');
  if (a > b) throw new Error(`--desde (${a}) es posterior a --hasta (${b}).`);
  const out = [];
  let y = +a.slice(0, 4), mo = +a.slice(4);
  while (`${y}${String(mo).padStart(2, '0')}` <= b) {
    out.push(`${y}${String(mo).padStart(2, '0')}`);
    if (++mo > 12) { mo = 1; y++; }
  }
  if (out.length > 24) throw new Error(`El rango tiene ${out.length} meses: pedi de a 24 como maximo.`);
  return out;
}

// Periodo (y dia, si lo trae) a partir del nombre que le pone el banco.
// Ej: Extracto_Cuentas_Galicia_2026_01_23.pdf -> { periodo: '202601', dia: 23 }
function periodoDeArchivo(nombre) {
  const s = path.basename(String(nombre || '')).replace(/\(\d+\)/g, ' ');
  let m = s.match(/(?:^|[^\d])(20\d{2})[-_. ]?(0[1-9]|1[0-2])(?:[-_. ]?(0[1-9]|[12]\d|3[01]))?(?!\d)/);
  if (m) return { periodo: m[1] + m[2], dia: m[3] ? +m[3] : null };
  m = s.match(/(?:^|[^\d])(?:(0[1-9]|[12]\d|3[01])[-_.])?(0[1-9]|1[0-2])[-_.](20\d{2})(?!\d)/);
  if (m) return { periodo: m[3] + m[2], dia: m[1] ? +m[1] : null };
  return null;
}

// ── sociedad ────────────────────────────────────────────────────────────────
const partes = (carpeta) => {
  const m = path.basename(carpeta).match(/^\s*\d+\s*-\s*(.+?)\s*-\s*(\d{11})\s*$/);
  return m ? { nombre: m[1], cuit: m[2] } : null;
};

// Por CUIT (preferido) o por nombre de carpeta. Nunca adivina entre varias.
function sociedad({ cuit, nombre }) {
  if (cuit) {
    const dir = rutas.carpetaSociedad(cuit);
    return { dir, ...partes(dir) };
  }
  if (!nombre) throw new Error('Falta --cuit o --nombre de la sociedad.');
  const base = rutas.raiz();
  const buscado = norm(nombre);
  const todas = fs.readdirSync(base, { withFileTypes: true })
    .filter(d => d.isDirectory() && partes(d.name))
    .map(d => ({ dir: path.join(base, d.name), ...partes(d.name) }));
  const exactas = todas.filter(s => norm(s.nombre) === buscado);
  const cand = exactas.length ? exactas : todas.filter(s => norm(s.nombre).includes(buscado));
  if (cand.length === 1) return cand[0];
  if (!cand.length) throw new Error(`No hay ninguna sociedad en la unidad que se llame "${nombre}". Pasa --cuit.`);
  throw new Error(`"${nombre}" coincide con ${cand.length} sociedades: ` +
    cand.map(s => `${s.nombre} (${s.cuit})`).join(' | ') + '. Pasa --cuit.');
}

const slug = (nombre) => String(nombre).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');

// ── bancos de la sociedad segun el Maestro ──────────────────────────────────
// Hoja MAESTRO: fila 2 de encabezados; columnas Galicia, Galicia +, Comafi,
// Hipotecario, Patagonia, Provincia, Nacion, Frances, Macro, Santander con "Si".
function bancosDelMaestro(cuit) {
  const archivo = path.join(rutas.raiz(), '0 - ESTUDIO - 00000000000', 'Maestro de Sociedades.xlsx');
  const filas = readWorkbook(fs.readFileSync(archivo)).rows('MAESTRO');
  const iEnc = filas.findIndex(f => f.some(c => /^CUIT$/i.test(String(c).trim())));
  if (iEnc < 0) throw new Error('No encuentro la fila de encabezados en la hoja MAESTRO.');
  const enc = filas[iEnc].map(c => String(c).trim());
  const cCuit = enc.findIndex(c => /^CUIT$/i.test(c));
  const COLS = ['Galicia', 'Galicia +', 'Comafi', 'Hipotecario', 'Patagonia', 'Provincia',
                'Nacion', 'Frances', 'Macro', 'Santander'];
  const fila = filas.slice(iEnc + 1).find(f => String(f[cCuit]).replace(/\D/g, '') === String(cuit));
  if (!fila) throw new Error(`El CUIT ${cuit} no esta en la hoja MAESTRO del Maestro de Sociedades.`);
  const si = (v) => /^s[ií]/i.test(String(v || '').trim());
  const out = [];
  for (const col of COLS) {
    // Comparacion exacta: con norm() "Galicia +" quedaria igual a "Galicia".
    const i = enc.findIndex(c => c.replace(/\s+/g, ' ').toLowerCase() === col.toLowerCase());
    if (i >= 0 && si(fila[i])) out.push(col === 'Galicia +' ? 'GALICIA +' : canonico(col));
  }
  return out;
}

// ── carpeta de salida (default) ─────────────────────────────────────────────
function salidaBase(salida) {
  return salida || path.join(process.env.BVA_SALIDAS_PATH ||
    path.join(require('os').homedir(), 'Documents', 'BVA-salidas'), 'extractos');
}

// [base]/[N] - [Sociedad] - [CUIT]/[BANCO]
function carpetaSalida(base, soc, banco) {
  return { dir: path.join(base, path.basename(soc.dir), canonico(banco)), creada: false };
}

// ── carpeta del banco en la unidad (--unidad) ───────────────────────────────────────────────────────
// Reusa la que exista (sin distinguir mayusculas, "bco. X", "X cta ...",
// Frances = BBVA). Si hay mas de una candidata devuelve la lista para preguntar.
function candidatasBanco(dirBancos, banco) {
  const b = canonico(banco);
  if (!fs.existsSync(dirBancos)) return [];
  return fs.readdirSync(dirBancos, { withFileTypes: true })
    .filter(d => d.isDirectory())
    .map(d => d.name)
    .filter(n => {
      const x = norm(n).replace(/^BCO\s+/, '');
      const primera = canonico(x.split(' ')[0]);
      return canonico(x) === b || primera === b || (b === 'BBVA' && /^(FRANCES|BBVA)/.test(x));
    });
}

// Devuelve { dir, creada } o { candidatas } si hay que elegir.
function carpetaBanco(soc, banco, { elegida } = {}) {
  if (!/^3[034]/.test(soc.cuit) && !fs.existsSync(path.join(soc.dir, BANCOS_DIR))) {
    throw new Error(`${soc.nombre} (${soc.cuit}) no tiene "${BANCOS_DIR}" y es persona humana: ` +
                    'confirma a mano si corresponde crearla.');
  }
  const dirBancos = path.join(soc.dir, BANCOS_DIR);
  if (elegida) return { dir: path.join(dirBancos, elegida), creada: false };
  const cand = candidatasBanco(dirBancos, banco);
  if (cand.length === 1) return { dir: path.join(dirBancos, cand[0]), creada: false };
  if (cand.length > 1) return { candidatas: cand };
  return { dir: path.join(dirBancos, canonico(banco)), creada: true };
}

// ── archivado ───────────────────────────────────────────────────────────────
const sha = (buf) => crypto.createHash('sha256').update(buf).digest('hex');

function nombreArchivo({ soc, banco, periodo: p, semana, ext }) {
  const base = `${slug(soc.nombre)}_${canonico(banco).toLowerCase()}_${p.slice(0, 4)}-${p.slice(4)}`;
  return `${base}_${semana ? `semana${semana}` : 'mensual'}${ext}`;
}

// Copia `origen` a la carpeta del banco/ano. Devuelve { estado, destino }:
//   'guardado' | 'ya_estaba' (mismo contenido) — nunca pisa.
function archivar({ origen, dirBanco, periodo: p, nombre }) {
  const dirAno = path.join(dirBanco, p.slice(0, 4));
  const buf = fs.readFileSync(origen);
  if (!buf.length) throw new Error(`El archivo descargado esta vacio: ${path.basename(origen)}`);
  // Un "PDF" que no empieza con %PDF suele ser la pagina de error o de sesion
  // vencida del banco: no se archiva.
  if (/\.pdf$/i.test(nombre) && buf.slice(0, 1024).indexOf('%PDF') < 0) {
    throw new Error(`${path.basename(origen).replace(/^\d+-/, "")} no es un PDF valido (probablemente una pagina de error del banco).`);
  }
  const hash = sha(buf);

  // Si ya esta (con cualquier nombre) en la carpeta del ano, o en la vieja
  // [BANCO]/[AAAAMM]/ que usaba la version anterior, no se copia de nuevo.
  for (const dir of [dirAno, path.join(dirBanco, p)]) {
    if (!fs.existsSync(dir)) continue;
    for (const f of fs.readdirSync(dir)) {
      const p2 = path.join(dir, f);
      if (fs.statSync(p2).isFile() && sha(fs.readFileSync(p2)) === hash) {
        return { estado: 'ya_estaba', destino: p2 };
      }
    }
  }
  fs.mkdirSync(dirAno, { recursive: true });
  const ext = path.extname(nombre);
  const raiz = nombre.slice(0, -ext.length);
  let destino = path.join(dirAno, nombre);
  for (let n = 2; fs.existsSync(destino); n++) destino = path.join(dirAno, `${raiz}_${n}${ext}`);
  fs.copyFileSync(origen, destino, fs.constants.COPYFILE_EXCL);
  return { estado: 'guardado', destino };
}

module.exports = {
  BANCOS_DIR, periodo, rango, periodoDeArchivo, sociedad, slug, bancosDelMaestro,
  candidatasBanco, carpetaBanco, salidaBase, carpetaSalida, nombreArchivo, archivar,
};
