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
//         TSH/                         <- aca va el recibo

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

// Carpeta del mes dentro de "01-Impuestos Mensuales".
//
// Conviven dos esquemas: el mes suelto ("01-Impuestos Mensuales/202608") y el
// mes dentro de una carpeta de año ("01-Impuestos Mensuales/2026/202608", unas
// 20 sociedades en 09/2026). Si existe la carpeta del año y no hay un mes suelto
// con ese nombre, el mes va adentro del año: crear "202608" al lado de "2026"
// deja dos lugares para el mismo periodo.
function carpetaMes(mensuales, periodoAAAAMM) {
  const suelto = path.join(mensuales, periodoAAAAMM);
  const anio = path.join(mensuales, periodoAAAAMM.slice(0, 4));
  if (!fs.existsSync(suelto) && fs.existsSync(anio) && fs.statSync(anio).isDirectory()) {
    return path.join(anio, periodoAAAAMM);
  }
  return suelto;
}

// Carpeta TSH del periodo (tasas municipales: Seguridad e Higiene y las
// boletas de marquesina / Publicidad y Propaganda), creandola si falta.
// Es donde el equipo guarda "AAAAMM-boleta marquesina- <Sociedad>.pdf".
function carpetaTsh(cuit, periodoAAAAMM, { crear = true } = {}) {
  const soc = carpetaSociedad(cuit);
  const mensuales = path.join(soc, '01-Impuestos Mensuales');
  if (!fs.existsSync(mensuales)) {
    if (!crear) throw new Error(`Falta "01-Impuestos Mensuales" en ${soc}`);
    fs.mkdirSync(mensuales, { recursive: true });
  }
  const destino = path.join(carpetaMes(mensuales, periodoAAAAMM), 'TSH');
  if (!fs.existsSync(destino)) {
    if (!crear) throw new Error(`Falta ${destino}`);
    fs.mkdirSync(destino, { recursive: true });
  }
  return destino;
}

// Nombre corto de la sociedad segun su carpeta: "3 - Farmacia Ejemplo - 30712345678" -> "Farmacia Ejemplo".
function nombreSociedad(cuit) {
  const base = path.basename(carpetaSociedad(cuit));
  const m = base.match(/^\s*\d+\s*-\s*(.+?)\s*-\s*\d{11}\s*$/);
  return (m ? m[1] : base).normalize('NFC');
}

module.exports = { raiz, carpetaSociedad, carpetaMes, carpetaTsh, nombreSociedad, RAIZ_DEFECTO };
