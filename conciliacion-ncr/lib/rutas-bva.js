// rutas-bva.js — ubica las carpetas de la unidad compartida del estudio.
//
// La unidad esta montada como I:\ en las maquinas del estudio, asi que se
// trabaja con fs normal. No hace falta el conector de Drive (que era lo que
// obligaba a pasar por Cowork).
//
// Estructura:
//   [N] - [Sociedad] - [CUIT]/        ej. "3 - Farmacia Ejemplo - 30712345678"
//     01-Impuestos Mensuales/
//       AAAAMM/                        ej. "202608"
//         IVA/                         <- aca va el archivo

const fs   = require('fs');
const path = require('path');

const RAIZ_DEFECTO = 'I:\\Unidades compartidas\\BVA - Sociedades Farmaceuticas';

function raiz() {
  const r = process.env.BVA_UNIDAD_PATH && process.env.BVA_UNIDAD_PATH.trim()
    ? process.env.BVA_UNIDAD_PATH.trim()
    : RAIZ_DEFECTO;
  if (!fs.existsSync(r)) {
    throw new Error(
      `No encuentro la unidad compartida en "${r}". ` +
      `Verifica que Google Drive este montado, o defini BVA_UNIDAD_PATH en el .env.`
    );
  }
  return r;
}

// Carpeta de la sociedad, buscada por CUIT (que es el desempate definitivo:
// varios nombres se parecen entre si — dos sociedades distintas
// pueden compartir parte del nombre).
function carpetaSociedad(cuit) {
  const c = String(cuit).replace(/\D/g, '');
  const base = raiz();
  const candidatas = fs.readdirSync(base, { withFileTypes: true })
    .filter(d => d.isDirectory() && d.name.replace(/\D/g, '').includes(c));

  if (candidatas.length === 0) {
    throw new Error(`No hay ninguna carpeta en la unidad compartida cuyo nombre contenga el CUIT ${c}.`);
  }
  if (candidatas.length > 1) {
    throw new Error(
      `El CUIT ${c} matchea ${candidatas.length} carpetas: ${candidatas.map(d => d.name).join(' | ')}. ` +
      `Resolvelo a mano antes de seguir.`
    );
  }
  return path.join(base, candidatas[0].name);
}

// Carpeta IVA del periodo, creandola si falta.
//
// Ojo: en algunas sociedades la subcarpeta existe como "AAAAMM-IVA" en vez de
// "AAAAMM/IVA" (hay sociedades armadas asi). Si ya existe con ese nombre se
// usa esa, en vez de crear una segunda que despues nadie sabe cual es la buena.
function carpetaIva(cuit, periodoAAAAMM, { crear = true } = {}) {
  const soc = carpetaSociedad(cuit);
  const mensuales = path.join(soc, '01-Impuestos Mensuales');
  if (!fs.existsSync(mensuales)) {
    if (!crear) throw new Error(`Falta "01-Impuestos Mensuales" en ${soc}`);
    fs.mkdirSync(mensuales, { recursive: true });
  }

  // Variantes que ya existen en la unidad, en orden de preferencia.
  for (const variante of [
    path.join(mensuales, periodoAAAAMM, `${periodoAAAAMM}-IVA`),   // ej. 202608/202608-IVA
    path.join(mensuales, `${periodoAAAAMM}-IVA`),
  ]) {
    if (fs.existsSync(variante)) return variante;
  }

  const destino = path.join(mensuales, periodoAAAAMM, 'IVA');
  if (!fs.existsSync(destino)) {
    if (!crear) throw new Error(`Falta ${destino}`);
    fs.mkdirSync(destino, { recursive: true });
  }
  return destino;
}

// Carpeta ncr/ del periodo (NCR de Suizo/Monroe + salida de conciliacion-ncr).
//
// Conviven varias ubicaciones porque se movieron a mano: [IVA]/ncr (la de
// carpetaIva), y tambien AAAAMM/ncr suelta en el mes (algunas sociedades,
// 08/2026). Se usa la que exista; si hay mas de una, se corta antes que elegir.
// Si no existe ninguna, [carpetaIva]/ncr (creandola si crear=true).
function carpetaNcr(cuit, periodoAAAAMM, { crear = true } = {}) {
  const mensuales = path.join(carpetaSociedad(cuit), '01-Impuestos Mensuales');
  const candidatas = [
    path.join(mensuales, periodoAAAAMM, `${periodoAAAAMM}-IVA`, 'ncr'),
    path.join(mensuales, `${periodoAAAAMM}-IVA`, 'ncr'),
    path.join(mensuales, periodoAAAAMM, 'IVA', 'ncr'),
    path.join(mensuales, periodoAAAAMM, 'ncr'),
  ];
  const hay = candidatas.filter((d) => fs.existsSync(d));
  if (hay.length > 1) {
    throw new Error(`Hay ${hay.length} carpetas ncr/ para ${periodoAAAAMM}: ${hay.join(' | ')}. Deja una sola.`);
  }
  if (hay.length === 1) return hay[0];
  const destino = path.join(carpetaIva(cuit, periodoAAAAMM, { crear }), 'ncr');
  if (crear) fs.mkdirSync(destino, { recursive: true });
  return destino;
}

// Nombre final del archivo, segun la convencion del estudio.
const nombreArchivo = (periodoAAAAMM, libro, ext) =>
  `${periodoAAAAMM} - PORTAL IVA - ${libro.toUpperCase()}.${ext}`;

module.exports = { raiz, carpetaSociedad, carpetaIva, carpetaNcr, nombreArchivo, RAIZ_DEFECTO };
