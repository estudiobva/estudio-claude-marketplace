// token-sociedad.js — el <TOKEN> que el equipo usa en los nombres de archivo de cada
// sociedad ("AAAAMM-IIBB-DJ-<TOKEN>.pdf", "AAAAMM-IVA-DDJJ-<TOKEN>.pdf").
//
// No se deduce del nombre de la carpeta (muchas sociedades usan una sigla que no
// sale de ahi, por ejemplo S65): se toma el que ya tienen los archivos de esa sociedad en
// 01-Impuestos Mensuales, primero los del mismo tipo y si no los del otro. Si la
// sociedad no tiene ninguno, no se inventa: se devuelve null y el script pide --token.

const fs = require('fs');
const path = require('path');
const rutas = require('./rutas-bva');

const PATRONES = {
  'iibb-dj': /^\d{6}\s*-\s*IIBB\s*-\s*DJ\s*-\s*(.+?)\.pdf$/i,
  'iva-ddjj': /^\d{6}\s*-\s*IVA\s*-\s*DDJJ\s*-\s*(.+?)\.pdf$/i,
};

function contar(cuit, re) {
  const base = path.join(rutas.carpetaSociedad(cuit), '01-Impuestos Mensuales');
  const cuenta = new Map();
  const recorrer = (dir, nivel) => {
    let entradas;
    try { entradas = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entradas) {
      if (e.isDirectory()) { if (nivel < 3) recorrer(path.join(dir, e.name), nivel + 1); continue; }
      const m = e.name.match(re);
      if (m) cuenta.set(m[1].trim(), (cuenta.get(m[1].trim()) || 0) + 1);
    }
  };
  recorrer(base, 0);
  return [...cuenta.entries()].sort((a, b) => b[1] - a[1]);
}

// tipo: 'iibb-dj' | 'iva-ddjj'. Devuelve el token mas usado o null.
function tokenSociedad(cuit, tipo) {
  const orden = tipo === 'iva-ddjj' ? ['iva-ddjj', 'iibb-dj'] : ['iibb-dj', 'iva-ddjj'];
  for (const t of orden) {
    const r = contar(cuit, PATRONES[t]);
    if (r.length) return r[0][0];
  }
  return null;
}

module.exports = { tokenSociedad };
