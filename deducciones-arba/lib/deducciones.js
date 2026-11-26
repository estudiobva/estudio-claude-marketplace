// deducciones.js — cruce del ZIP de deducciones de ARBA (RP-CUIT-AAAAMMC.zip) contra lo que ARBA
// tomó en la DJ (Deducciones_DJ_<nro>.xls) y armado de los lotes de modificación (operación M).
//
// El problema: los bancos informan percepciones con tipo "O", sucursal 0 y número 0, y el SIRTAC
// a veces repite el número de liquidación. ARBA toma UNA por agente y descarta el resto como
// duplicadas. Cargarlas de alta con otro número da "Deducción duplicada". Lo que ARBA acepta es
// modificar (M) el certificado que tomó, poniéndole la suma de todas las de ese agente.
//
// Diseños de registro DDJJ web vigentes (DISENIODEREGISTROiibbddjjweb.pdf):
//   Percepciones (P), 53 car.: CUIT 99-99999999-9 | fecha dd/mm/aaaa | tipo | letra | sucursal 5 |
//                              emisión 8 | importe 14 con coma | operación
//   SIRTAC (T), 70 car.: CUIT | AAAA | MM | fecha | liquidación 20 | signo 0/1 | monto 19 con coma | operación
// Nombre del lote: CUIT-AAAAMMM-<tipo>-LOTE1_<MD5 del zip en mayúsculas>.zip, con un TXT adentro
// que se llama igual sin el MD5.

const zlib = require('zlib');
const crypto = require('crypto');

const r2 = (x) => Math.round(x * 100) / 100;
const num = (s) => r2(Number(String(s).replace(/\./g, '').replace(',', '.')));
const coma = (x) => r2(x).toFixed(2).replace('.', ',');
const cuitGuiones = (c) => { c = String(c).replace(/\D/g, ''); return `${c.slice(0, 2)}-${c.slice(2, 10)}-${c.slice(10)}`; };

// ── ZIP mínimo (leer y escribir, sin dependencias) ───────────────────────────
// Se lee por el directorio central: el ZIP de ARBA trae los tamaños en "data descriptors"
// (el encabezado local dice 0), así que recorrer los encabezados locales no sirve.
function leerZip(buf) {
  const out = {};
  let e = buf.length - 22; while (e >= 0 && buf.readUInt32LE(e) !== 0x06054b50) e--;
  if (e < 0) throw new Error('ZIP inválido');
  let i = buf.readUInt32LE(e + 16); const n = buf.readUInt16LE(e + 10);
  for (let k = 0; k < n; k++) {
    const met = buf.readUInt16LE(i + 10), csize = buf.readUInt32LE(i + 20);
    const nlen = buf.readUInt16LE(i + 28), xlen = buf.readUInt16LE(i + 30), clen = buf.readUInt16LE(i + 32);
    const off = buf.readUInt32LE(i + 42);
    const name = buf.slice(i + 46, i + 46 + nlen).toString('latin1');
    const ln = buf.readUInt16LE(off + 26), lx = buf.readUInt16LE(off + 28);
    const data = buf.slice(off + 30 + ln + lx, off + 30 + ln + lx + csize);
    out[name] = met === 8 ? zlib.inflateRawSync(data) : data;
    i += 46 + nlen + xlen + clen;
  }
  return out;
}
function crc32(b) { let c = ~0; for (const x of b) { c ^= x; for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xEDB88320 & -(c & 1)); } return ~c >>> 0; }
function escribirZip(nombre, contenido) {
  const n = Buffer.from(nombre, 'latin1'), d = zlib.deflateRawSync(contenido), crc = crc32(contenido);
  const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(8, 8);
  lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(d.length, 18); lh.writeUInt32LE(contenido.length, 22); lh.writeUInt16LE(n.length, 26);
  const ch = Buffer.alloc(46); ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(8, 10);
  ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(d.length, 20); ch.writeUInt32LE(contenido.length, 24); ch.writeUInt16LE(n.length, 28);
  const cd = Buffer.concat([ch, n]), off = 30 + n.length + d.length;
  const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(1, 8); end.writeUInt16LE(1, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(off, 16);
  return Buffer.concat([lh, n, d, cd, end]);
}

// ── TXT del ZIP de ARBA ──────────────────────────────────────────────────────
function lineas(buf) { return buf.toString('latin1').split(/\r?\n/).filter((l) => l.trim()); }

// CP: <prefijo><CUIT 13><fecha><sucursal><emisión 20><tipo><letra><importe>
function parseCP(buf) {
  return lineas(buf).map((l, i) => {
    const m = l.match(/^.*?(\d\d-\d{8}-\d)(\d\d\/\d\d\/\d{4})(\d+?)(\d{20})([A-Z ])([A-Z ])(-?\d+,\d\d)$/);
    if (!m) throw new Error(`CP línea ${i + 1} con formato desconocido: ${l}`);
    return { linea: i + 1, raw: l, cuit: m[1].replace(/-/g, ''), fecha: m[2], suc: +m[3], nro: +m[4], tipo: m[5], letra: m[6], imp: num(m[7]) };
  });
}
// CT (SIRTAC): <prefijo><CUIT><fecha><dígitos><tipo><letra><liquidación 20><importe>
function parseCT(buf) {
  return lineas(buf).map((l, i) => {
    const m = l.match(/^.*?(\d\d-\d{8}-\d)(\d\d\/\d\d\/\d{4})\d+[A-Z ][A-Z ](\d{20})(-?\d+,\d\d)$/);
    if (!m) throw new Error(`CT línea ${i + 1} con formato desconocido: ${l}`);
    return { linea: i + 1, raw: l, cuit: m[1].replace(/-/g, ''), fecha: m[2], liq: m[3], imp: num(m[4]) };
  });
}
// Importe al final de la línea, de ancho fijo: 15 en CP/CR/CB, 11 en CT (antes va pegada la liquidación).
const ANCHO = { CP: 15, CR: 15, CB: 15, CT: 11 };
const totalTxt = (buf, t) => r2(lineas(buf).reduce((s, l) => s + num(l.slice(-ANCHO[t])), 0));

// ── Deducciones_DJ_<nro>.xls (SpreadsheetML que baja ARBA) ───────────────────
function leerXlsDJ(buf) {
  const txt = buf.toString('utf8');
  if (!/<Workbook/i.test(txt.slice(0, 2000))) throw new Error('El Deducciones_DJ no es el XML que baja ARBA (¿lo abrieron y guardaron en Excel?). Bajalo de nuevo.');
  const hojas = {};
  const re = /<(?:ss:)?Worksheet[^>]*?Name="([^"]+)"[^>]*>([\s\S]*?)<\/(?:ss:)?Worksheet>/g; let m;
  while ((m = re.exec(txt))) {
    const filas = [...m[2].matchAll(/<(?:ss:)?Row[^>]*>([\s\S]*?)<\/(?:ss:)?Row>/g)].map((r) =>
      [...r[1].matchAll(/<(?:ss:)?Cell[^>]*?(?:\/>|>([\s\S]*?)<\/(?:ss:)?Cell>)/g)].map((c) => {
        const d = (c[1] || '').match(/<(?:ss:)?Data[^>]*>([\s\S]*?)<\/(?:ss:)?Data>/);
        return d ? d[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>') : '';
      }));
    hojas[m[1]] = filas.slice(1);
  }
  return hojas;
}

// Multiconjunto: cuenta claves y va descontando.
function contador(arr, key) { const c = new Map(); for (const x of arr) { const k = key(x); c.set(k, (c.get(k) || 0) + 1); } return c; }
function marcarTomadas(txt, xls, key) {
  const c = contador(xls, key);
  for (const t of txt) { const k = key(t); const n = c.get(k) || 0; t.tomada = n > 0; if (n > 0) c.set(k, n - 1); }
  return [...c.entries()].filter(([, n]) => n > 0).map(([k]) => k); // en la DJ y no en el TXT
}

// ── Cruce completo ───────────────────────────────────────────────────────────
function cruzar({ zip, xls, cuit, periodo }) {
  const files = leerZip(zip);
  const pick = (t) => { const k = Object.keys(files).find((n) => new RegExp(`\\d{11}${t}\\d`).test(n)); return k ? files[k] : Buffer.alloc(0); };
  const hojas = leerXlsDJ(xls);
  const res = { cuit, periodo, secciones: {}, percepciones: {}, sirtac: {}, lotes: [] };

  // Totales por sección
  const H = { CP: ['Percepciones', 8], CR: ['Retenciones', 6], CB: ['Retenciones bancarias', 8], CT: ['Retenciones Sirtac', 4] };
  for (const [t, [hoja, col]] of Object.entries(H)) {
    const b = pick(t); const filasX = hojas[hoja] || [];
    res.secciones[t] = { hoja, lineas_txt: lineas(b).length, total_txt: b.length ? totalTxt(b, t) : 0,
      filas_dj: filasX.length, total_dj: r2(filasX.reduce((s, f) => s + (f[col] ? num(f[col]) : 0), 0)) };
  }

  // Percepciones
  const cp = parseCP(pick('CP'));
  const xp = (hojas['Percepciones'] || []).map((f) => ({ tipo: (f[0] || '').trim() || 'O', letra: f[1] || ' ', suc: +(f[2] || 0), nro: +(f[3] || 0), cuit: String(f[4]).replace(/\D/g, ''), fecha: f[5], imp: num(f[6]) }));
  const kP = (d) => [d.cuit, d.tipo, d.suc, d.nro, d.fecha, d.imp].join('|');
  const extraP = marcarTomadas(cp, xp, kP);
  const bancos = new Map();
  for (const d of cp.filter((x) => x.tipo === 'O' && x.suc === 0 && x.nro === 0)) { if (!bancos.has(d.cuit)) bancos.set(d.cuit, []); bancos.get(d.cuit).push(d); }
  const lineasP = [], detalleBancos = [];
  for (const [c, ds] of [...bancos].sort((a, b) => a[0].localeCompare(b[0]))) {
    const tomada = ds.find((d) => d.tomada); const total = r2(ds.reduce((s, d) => s + d.imp, 0));
    detalleBancos.push({ cuit: c, lineas: ds.length, tomada: tomada && { fecha: tomada.fecha, imp: tomada.imp }, total, faltan: ds.filter((d) => !d.tomada).length });
    if (ds.length > 1 && ds.some((d) => !d.tomada)) {
      if (!tomada) { detalleBancos[detalleBancos.length - 1].error = 'ARBA no tomó ninguna de este banco: no hay certificado para modificar'; continue; }
      const l = cuitGuiones(c) + tomada.fecha + 'O' + ' ' + '00000' + '00000000' + coma(total).padStart(14, '0') + 'M';
      if (l.length !== 53) throw new Error('línea P de largo ' + l.length);
      lineasP.push(l);
    }
  }
  const noTomadasP = cp.filter((d) => !d.tomada);
  const sinNro = noTomadasP.filter((d) => !(d.tipo === 'O' && d.suc === 0 && d.nro === 0));
  res.percepciones = { no_tomadas: noTomadasP.length, importe_no_tomado: r2(noTomadasP.reduce((s, d) => s + d.imp, 0)),
    en_dj_y_no_en_txt: extraP.length, no_tomadas_con_numero: sinNro.length, bancos: detalleBancos };

  // SIRTAC
  const ct = parseCT(pick('CT'));
  const xt = (hojas['Retenciones Sirtac'] || []).map((f) => ({ cuit: String(f[0]).replace(/\D/g, ''), fecha: f[1], imp: num(f[2]) }));
  const kT = (d) => [d.cuit, d.fecha, d.imp].join('|');
  const extraT = marcarTomadas(ct, xt, kT);
  const grupos = new Map();
  for (const d of ct) { const k = d.cuit + '|' + d.liq; if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(d); }
  const lineasT = [], detalleT = [];
  for (const ds of grupos.values()) {
    if (ds.length < 2 || ds.every((d) => d.tomada)) continue;
    const tomada = ds.find((d) => d.tomada); const total = r2(ds.reduce((s, d) => s + d.imp, 0));
    const det = { cuit: ds[0].cuit, liquidacion: ds[0].liq, lineas: ds.map((d) => ({ fecha: d.fecha, imp: d.imp, tomada: d.tomada })), total };
    detalleT.push(det);
    if (!tomada) { det.error = 'ARBA no tomó ninguna de esta liquidación'; continue; }
    const l = cuitGuiones(ds[0].cuit) + periodo.slice(0, 4) + periodo.slice(4) + tomada.fecha + ds[0].liq + (total < 0 ? '1' : '0') + coma(Math.abs(total)).padStart(19, '0') + 'M';
    if (l.length !== 70) throw new Error('línea T de largo ' + l.length);
    lineasT.push(l);
  }
  const noTomadasT = ct.filter((d) => !d.tomada);
  res.sirtac = { no_tomadas: noTomadasT.length, importe_no_tomado: r2(noTomadasT.reduce((s, d) => s + d.imp, 0)), en_dj_y_no_en_txt: extraT.length, liquidaciones: detalleT };

  // Lotes
  for (const [tipo, ls] of [['P', lineasP], ['T', lineasT]]) {
    if (!ls.length) continue;
    const base = `${cuit}-${periodo}M-${tipo}-LOTE1`;
    const txt = Buffer.from(ls.join('\r\n') + '\r\n', 'ascii');
    const z = escribirZip(base + '.txt', txt);
    const md5 = crypto.createHash('md5').update(z).digest('hex').toUpperCase();
    res.lotes.push({ tipo, base, txt, zip: z, zipNombre: `${base}_${md5}.zip`, lineas: ls });
  }
  res.faltante = r2(res.percepciones.importe_no_tomado + res.sirtac.importe_no_tomado);
  res.otras_diferencias = ['CR', 'CB'].filter((t) => Math.abs(res.secciones[t].total_txt - res.secciones[t].total_dj) > 0.005)
    .map((t) => `${res.secciones[t].hoja}: TXT ${res.secciones[t].total_txt} vs DJ ${res.secciones[t].total_dj} (revisar a mano)`);
  if (sinNro.length) res.otras_diferencias.push(`${sinNro.length} percepciones con número que ARBA no tomó (no son del caso de los bancos, revisar a mano)`);
  res._detalle = { cp, ct };
  return res;
}

module.exports = { cruzar, leerZip, escribirZip, leerXlsDJ, parseCP, parseCT, cuitGuiones, coma };
