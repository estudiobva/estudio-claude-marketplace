// Prueba offline: rehace el cruce con los archivos ya archivados en la unidad y compara los lotes
// contra los que se armaron a mano (y que ARBA aceptó en Argentina).
const fs = require('fs'), path = require('path');
const { cruzar } = require('../lib/deducciones');
const B = 'G:/Unidades compartidas/BVA - Sociedades Farmaceuticas';
let ok = true;
for (const cuit of process.argv.slice(2)) {
  const soc = fs.readdirSync(B).find((d) => d.includes(cuit));
  const dir = path.join(B, soc, '01-Impuestos Mensuales/2026/202610/ARBA 202610 DEDUCCIONES');
  const fs_ = fs.readdirSync(dir);
  const zip = fs.readFileSync(path.join(dir, fs_.find((f) => /^RP-.*\.zip$/.test(f))));
  const xlsF = fs_.find((f) => /^Deducciones_DJ_.*\.xls$/.test(f)) ;
  const xls = fs.readFileSync(xlsF ? path.join(dir, xlsF) : path.join(dir, '..', fs.readdirSync(path.join(dir, '..')).find((f) => /^Deducciones_DJ_/.test(f))));
  let r; try { r = cruzar({ zip, xls, cuit, periodo: '202610' }); } catch (e) { console.log(soc, 'ERROR', e.message); ok = false; continue; }
  console.log(`${soc}: faltante ${r.faltante} (perc ${r.percepciones.importe_no_tomado} / sirtac ${r.sirtac.importe_no_tomado}) otras: ${r.otras_diferencias.join('; ') || '-'}`);
  for (const l of r.lotes) {
    const manual = fs_.find((f) => f === l.base + '.txt') || (cuit === '30683756071' && l.tipo === 'P' ? '30683756071-202610M-P-LOTE5.txt' : null);
    const igual = manual && fs.readFileSync(path.join(dir, manual)).equals(l.txt);
    console.log(`   lote ${l.tipo}: ${l.lineas.length} líneas · igual al manual (${manual}): ${igual}`);
    if (!igual) { ok = false; console.log(l.lineas.join('\n')); }
  }
}
console.log(ok ? 'TODO OK' : 'HAY DIFERENCIAS');
