// ─────────────────────────────────────────────────────────────────────────────
// Vendorizado desde el plugin fisco-ar (MIT) — https://github.com/javiergradiche/fisco-ar-claude-plugin
// Commit de origen: 2ad06df · Copiado el 21/09/2026 al repo del Estudio BVA.
//
// Se copia en vez de depender del plugin para que una actualizacion del
// marketplace no pueda romper un cierre. Si ARCA cambia el login, el arreglo va
// ACA (y conviene mirar upstream por si ya lo resolvieron).
// ─────────────────────────────────────────────────────────────────────────────

// xlsx-min.js — lector mínimo de .xlsx (sin dependencias).
//
// Un .xlsx es un ZIP con XML adentro. Acá sólo necesitamos leer celdas como
// texto de una planilla chica, así que en vez de sumar SheetJS al plugin
// implementamos las dos piezas que hacen falta:
//
//   1. Un lector de ZIP (directorio central + inflate crudo con zlib).
//   2. Un parser de <sheetN.xml> + <sharedStrings.xml> que devuelve una matriz.
//
// Limitaciones asumidas a propósito: se lee todo en memoria, no se interpretan
// formatos de número ni fechas (todo vuelve como string, tal cual el XML) y no
// se resuelven fórmulas (se usa el último valor cacheado, que es lo que Excel
// guarda en <v>). Para el archivo de claves alcanza y sobra.

const zlib = require('zlib');

// ── ZIP ───────────────────────────────────────────────────────────────────────

// Devuelve un Map<nombre, Buffer> con el contenido del ZIP.
function unzip(buf) {
  // El End Of Central Directory está al final, pero puede tener comentario
  // atrás, así que lo buscamos hacia atrás desde el final.
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0 && i >= buf.length - 22 - 0xffff; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('no parece un archivo ZIP válido (falta el End Of Central Directory)');

  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16); // offset del directorio central

  const files = new Map();
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('entrada corrupta en el directorio central del ZIP');
    const method  = buf.readUInt16LE(p + 10);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const cmtLen  = buf.readUInt16LE(p + 32);
    const localAt = buf.readUInt32LE(p + 42);
    const name    = buf.toString('utf8', p + 46, p + 46 + nameLen);

    // El header local repite el nombre y trae SU propio largo de "extra",
    // que puede diferir del que figura en el directorio central.
    const lNameLen  = buf.readUInt16LE(localAt + 26);
    const lExtraLen = buf.readUInt16LE(localAt + 28);
    const dataAt    = localAt + 30 + lNameLen + lExtraLen;
    const compSize  = buf.readUInt32LE(p + 20);
    const raw       = buf.subarray(dataAt, dataAt + compSize);

    if (!name.endsWith('/')) {
      files.set(name, method === 0 ? Buffer.from(raw) : zlib.inflateRawSync(raw));
    }
    p += 46 + nameLen + extraLen + cmtLen;
  }
  return files;
}

// ── XML ───────────────────────────────────────────────────────────────────────

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

function decodeXml(s) {
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-z]+);/g, (all, ent) => {
    if (ent[0] === '#') {
      const code = ent[1] === 'x' ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : all;
    }
    return ENTITIES[ent] !== undefined ? ENTITIES[ent] : all;
  });
}

// Texto de todos los <t> de un fragmento (una <si> puede venir partida en runs).
function textOf(fragment) {
  let out = '';
  for (const m of fragment.matchAll(/<t\b[^>]*\/>|<t\b[^>]*>([\s\S]*?)<\/t>/g)) {
    out += decodeXml(m[1] || '');
  }
  return out;
}

// "BC" → 54 (índice 0-based de columna).
function colIndex(ref) {
  const letters = (ref.match(/^[A-Z]+/) || [''])[0];
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

// ── Planillas ─────────────────────────────────────────────────────────────────

function sharedStrings(files) {
  const xml = files.get('xl/sharedStrings.xml');
  if (!xml) return [];
  return [...xml.toString('utf8').matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => textOf(m[1]));
}

// Nombre de hoja → ruta dentro del zip, respetando el orden del workbook.
function sheetIndex(files) {
  const wb = files.get('xl/workbook.xml');
  if (!wb) return [];
  const relsXml = (files.get('xl/_rels/workbook.xml.rels') || Buffer.alloc(0)).toString('utf8');
  const rels = new Map(
    [...relsXml.matchAll(/<Relationship\b[^>]*>/g)].map(m => [
      (m[0].match(/Id="([^"]+)"/) || [])[1],
      (m[0].match(/Target="([^"]+)"/) || [])[1],
    ])
  );
  return [...wb.toString('utf8').matchAll(/<sheet\b[^>]*\/?>/g)].map((m, i) => {
    const name = decodeXml((m[0].match(/name="([^"]*)"/) || [, ''])[1]);
    const rid  = (m[0].match(/r:id="([^"]+)"/) || [])[1];
    let target = rels.get(rid) || `worksheets/sheet${i + 1}.xml`;
    target = target.replace(/^\/?xl\//, '').replace(/^\.\//, '');
    return { name, path: `xl/${target}` };
  });
}

// Matriz de strings de una hoja. Las filas y columnas vacías se rellenan con ''
// para que los índices coincidan con los de Excel (fila 4 → rows[3]).
function sheetRows(files, path) {
  const xml = (files.get(path) || Buffer.alloc(0)).toString('utf8');
  const strings = sharedStrings(files);
  const rows = [];

  for (const rm of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const rowNum = Number((rm[1].match(/\br="(\d+)"/) || [, 0])[1]) || rows.length + 1;
    const cells = [];
    for (const cm of rm[2].matchAll(/<c\b([^>]*)\/>|<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const attrs = cm[1] || cm[2] || '';
      const body  = cm[3] || '';
      const ref   = (attrs.match(/\br="([A-Z]+\d+)"/) || [, ''])[1];
      const type  = (attrs.match(/\bt="([^"]+)"/) || [, 'n'])[1];

      let value = '';
      if (type === 's') {
        const idx = Number(textOfV(body));
        value = strings[idx] !== undefined ? strings[idx] : '';
      } else if (type === 'inlineStr') {
        value = textOf(body);
      } else {
        value = decodeXml(textOfV(body));
      }

      const at = ref ? colIndex(ref) : cells.length;
      while (cells.length < at) cells.push('');
      cells[at] = value;
    }
    while (rows.length < rowNum - 1) rows.push([]);
    rows[rowNum - 1] = cells;
  }
  return rows;
}

// Contenido del <v> de una celda (o del <t> si vino como inline sin <v>).
function textOfV(body) {
  const m = body.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
  return m ? m[1] : textOf(body);
}

// API: abre un buffer .xlsx (ya descifrado) y devuelve { hojas, filas(nombre) }.
function readWorkbook(buf) {
  const files = unzip(buf);
  const sheets = sheetIndex(files);
  return {
    sheetNames: sheets.map(s => s.name),
    rows(name) {
      const sheet = name
        ? sheets.find(s => s.name.toLowerCase() === String(name).toLowerCase())
        : sheets[0];
      if (!sheet) throw new Error(`la planilla no tiene una hoja llamada "${name}"`);
      return sheetRows(files, sheet.path);
    },
  };
}

module.exports = { readWorkbook, unzip };
