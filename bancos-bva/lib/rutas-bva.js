// rutas-bva.js — ubica las carpetas de la unidad compartida del estudio.
//
// La unidad esta montada como I:\ en las maquinas del estudio, asi que se
// trabaja con fs normal. No hace falta el conector de Drive (que era lo que
// obligaba a pasar por Cowork).
//
// Estructura:
//   [N] - [Sociedad] - [CUIT]/        ej. "3 - Sociedad Ejemplo - 30000000007"
//     01-Impuestos Mensuales/
//       AAAAMM/                        ej. "202608"
//         IVA/                         <- aca va el archivo

const fs   = require('fs');
const path = require('path');

const RAIZ_DEFECTO = 'I:\\Unidades compartidas\\BVA - Sociedades Farmaceuticas';
const UNIDAD = path.join('Unidades compartidas', 'BVA - Sociedades Farmaceuticas');

// Donde puede estar montada la unidad: BVA_UNIDAD_PATH, I:\ (PCs del estudio) o
// Google Drive para escritorio en Mac (~/Library/CloudStorage/GoogleDrive-<cuenta>/).
function candidatasRaiz() {
  const out = [];
  if (process.env.BVA_UNIDAD_PATH && process.env.BVA_UNIDAD_PATH.trim()) {
    out.push(process.env.BVA_UNIDAD_PATH.trim().replace(/^~/, require('os').homedir()));
  }
  out.push(RAIZ_DEFECTO);
  const cs = path.join(require('os').homedir(), 'Library', 'CloudStorage');
  try {
    // Puede haber varias cuentas de Drive montadas: primero las del estudio.
    const cuentas = fs.readdirSync(cs).filter(d => /^GoogleDrive-/i.test(d))
      .sort((a, b) => /estudiobva/i.test(b) - /estudiobva/i.test(a));
    for (const d of cuentas) out.push(path.join(cs, d, UNIDAD));
  } catch { /* no es Mac o no hay Drive */ }
  return out;
}

function raiz() {
  const cand = candidatasRaiz();
  const r = cand.find(c => fs.existsSync(c));
  if (!r) {
    throw new Error(
      `No encuentro la unidad compartida (busque en: ${cand.join(' | ')}). ` +
      `Verifica que Google Drive este montado, o defini BVA_UNIDAD_PATH en el .env.`
    );
  }
  return r;
}

// Carpeta de la sociedad, buscada por CUIT (que es el desempate definitivo:
// varios nombres de carpeta se parecen entre si y un nombre corto puede estar
// contenido en el de dos sociedades distintas).
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
// "AAAAMM/IVA" (hay al menos un caso conocido). Si ya existe con ese nombre se
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
    path.join(mensuales, periodoAAAAMM, `${periodoAAAAMM}-IVA`),   // variante vista: 202608/202608-IVA
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

// Nombre final del archivo, segun la convencion del estudio.
const nombreArchivo = (periodoAAAAMM, libro, ext) =>
  `${periodoAAAAMM} - PORTAL IVA - ${libro.toUpperCase()}.${ext}`;

module.exports = { raiz, candidatasRaiz, carpetaSociedad, carpetaIva, nombreArchivo, RAIZ_DEFECTO };
