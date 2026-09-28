#!/usr/bin/env node
/**
 * planes-reporte.js — arma el REPORTE en PDF (y HTML) de planes de facilidades
 * a partir de lo que dejo planes-facilidades.js. No entra a ARCA.
 *
 *   node scripts/planes-reporte.js                    (relevamiento de hoy)
 *   node scripts/planes-reporte.js --fecha=2026-09-24
 *
 * Orden del reporte, fijado a pedido del estudio:
 *   1. Sociedades que NO se pudieron revisar (y por que).
 *   2. Sociedades con cuotas IMPAGAS: plan, cuota, vencimiento, dias de atraso,
 *      debitos rechazados, CBU y obligaciones del plan.
 *   3. Todas las sociedades: cada plan vigente con su numero, CBU declarado,
 *      obligaciones incluidas (solo el concepto) y estado de cuotas.
 */

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const { parseArgs, fail } = require('../lib/args');

require('../lib/env').loadEnv();
const args = parseArgs();
const hoy = new Date();
const iso = args['fecha'] || [hoy.getFullYear(), String(hoy.getMonth() + 1).padStart(2, '0'),
                              String(hoy.getDate()).padStart(2, '0')].join('-');
const DIR = path.join(args['salida'] || path.join((process.env.BVA_SALIDAS_PATH || path.join(require('os').homedir(), 'Documents', 'BVA-salidas')), 'planes'), iso);
const RESUMEN = path.join(DIR, 'lote-resumen.json');
if (!fs.existsSync(RESUMEN)) fail(`No hay relevamiento en ${DIR}. Corre primero planes-facilidades.js.`);

const estado = JSON.parse(fs.readFileSync(RESUMEN, 'utf-8'));
const lote = fs.existsSync(path.join(DIR, '_crudo', 'lote.json'))
  ? JSON.parse(fs.readFileSync(path.join(DIR, '_crudo', 'lote.json'), 'utf-8')) : { omitidos: [] };
const fechaTxt = iso.split('-').reverse().join('/');

const BANCOS = { '007': 'Galicia', '011': 'Nación', '014': 'Provincia', '015': 'ICBC', '017': 'BBVA',
  '027': 'Supervielle', '029': 'Ciudad', '034': 'Patagonia', '044': 'Hipotecario', '072': 'Santander',
  '150': 'HSBC / Galicia Más', '191': 'Credicoop', '285': 'Macro', '299': 'Comafi', '322': 'Industrial',
  '389': 'Columbia' };
const banco = (c) => BANCOS[String(c || '').slice(0, 3)] || '';
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const pesos = (n) => (n == null ? '' : '$ ' + (Math.abs(n) < 0.005 ? 0 : Number(n)).toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const cuitF = (c) => (String(c || '').length === 11 ? `${c.slice(0, 2)}-${c.slice(2, 10)}-${c.slice(10)}` : esc(c));
const cbuF = (p) => {
  if (!('cbuDeclarado' in p)) return '<span class="pend">pendiente</span>';
  const d = p.cbuDeclarado || '';
  if (!d) return '<span class="muted">sin dato</span>';
  let t = `<span class="mono">${esc(d)}</span>${banco(d) ? ` <span class="muted">${banco(d)}</span>` : ''}`;
  if (p.cbuVigente && p.cbuVigente !== d) {
    t += `<br><span class="muted">vigente hoy:</span> <span class="mono">${esc(p.cbuVigente)}</span> ${banco(p.cbuVigente)}`;
  }
  return t;
};
const conceptosF = (p) => (!('conceptos' in p) ? '<span class="pend">pendiente</span>' : (p.conceptos || []).length
  ? p.conceptos.map(esc).join('<br>') : '<span class="muted">sin dato</span>');

// ── Datos ───────────────────────────────────────────────────────────────────
const soc = Object.values(estado.sociedades).sort((a, b) => Number(a.pos) - Number(b.pos));
const datos = soc.map(s => {
  let d = null;
  if (s.ok && s.datos && fs.existsSync(s.datos)) d = JSON.parse(fs.readFileSync(s.datos, 'utf-8'));
  const vig = d ? d.planes.filter(p => /vigente/i.test(p.situacion)) : [];
  return { s, d, vig, impagas: vig.filter(p => p.cuotasImpagas) };
});

const QUE_HACER = {
  cuit_no_disponible: 'La CUIT no figura entre las asociadas del apoderado en Mis Facilidades: delegar el servicio en el Administrador de Relaciones.',
  login: 'El login se traba después de ingresar la CUIT (falló en dos intentos). Revisar la clave fiscal a mano en ARCA.',
  servicio_no_disponible: 'No tiene adherido el servicio Mis Facilidades.',
  credenciales_invalidas: 'ARCA rechazó la clave: revisar Claves_Organismos.xlsx.',
};
const noRevisadas = [
  ...datos.filter(x => !x.s.ok).map(x => ({ pos: x.s.pos, sociedad: x.s.sociedad, cuit: x.s.cuit,
    motivo: QUE_HACER[x.s.code] || x.s.error })),
  ...(lote.omitidos || []).filter(o => !/repetido/i.test(o.motivo))
    .map(o => ({ pos: o.pos, sociedad: o.titular, cuit: '', motivo: 'Sin CUIT o sin clave de ARCA en el archivo de claves.' })),
];
const conImpagas = datos.filter(x => x.impagas.length);
const relevadas = datos.filter(x => x.s.ok);
const totPlanes = relevadas.reduce((a, x) => a + x.vig.length, 0);
const totCuotasImp = conImpagas.reduce((a, x) => a + x.impagas.reduce((b, p) => b + p.cuotasImpagas, 0), 0);
const saldo = relevadas.reduce((a, x) => a + x.vig.reduce((b, p) => b + (p.saldoCapital || 0), 0), 0);
const sinVig = relevadas.filter(x => !x.vig.length).length;
const incompletos = relevadas.flatMap(x => x.vig.filter(p => (p.lecturaIncompleta || []).length)
  .map(p => ({ sociedad: x.s.sociedad, plan: p.nroPlan, falta: p.lecturaIncompleta })));
const planesPend = relevadas.reduce((a, x) => a + x.vig.filter(p => !('cbuDeclarado' in p)).length, 0);

// ── HTML ────────────────────────────────────────────────────────────────────
let h = `<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Planes de facilidades ${fechaTxt}</title><style>
  @page { size: A4 landscape; margin: 11mm 10mm; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 9pt; color: #1b1b1b; margin: 0; }
  h1 { font-size: 16pt; margin: 0 0 2px; color: #1F4E79; }
  h2 { font-size: 12pt; margin: 18px 0 6px; padding: 4px 8px; color: #fff; background: #1F4E79; }
  h2.rojo { background: #9C0006; } h2.gris { background: #595959; }
  .sub { color: #555; margin-bottom: 10px; }
  .kpis { display: flex; gap: 8px; margin: 8px 0 4px; }
  .kpi { border: 1px solid #ccc; border-radius: 4px; padding: 6px 10px; min-width: 110px; }
  .kpi b { display: block; font-size: 14pt; } .kpi.r { border-color: #9C0006; color: #9C0006; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 6px; page-break-inside: auto; }
  tr { page-break-inside: avoid; }
  th { background: #DDE7F0; text-align: left; font-size: 8pt; padding: 3px 5px; border: 1px solid #b7c7d6; }
  td { padding: 3px 5px; border: 1px solid #d0d0d0; vertical-align: top; }
  tr.imp td { background: #FFE3E6; } td.num { text-align: right; white-space: nowrap; }
  .mono { font-family: Menlo, Consolas, monospace; font-size: 8pt; }
  .muted { color: #777; } .tag { font-weight: bold; }
  .tag.r { color: #9C0006; } .tag.v { color: #1e6b30; } .tag.g { color: #666; }
  .socbloque { page-break-inside: avoid; margin-top: 10px; }
  .soctit { font-weight: bold; font-size: 10pt; border-bottom: 2px solid #1F4E79; padding: 2px 0; margin-bottom: 3px; }
  .soctit .tag { float: right; font-size: 9pt; }
  .alerta { border: 2px solid #9C0006; background: #FFF4F5; padding: 6px 8px; margin: 6px 0; page-break-inside: avoid; }
  .nota { font-size: 8pt; color: #555; margin-top: 12px; }
  .pend { color: #8a6d00; font-style: italic; }
  .aviso { border: 1px solid #d4b106; background: #fffbe6; padding: 6px 8px; margin: 8px 0; }
</style></head><body>
<h1>Planes de facilidades de pago — relevamiento de todas las sociedades</h1>
<div class="sub">Fuente: ARCA · Mis Facilidades · relevado el ${fechaTxt}. Cuota IMPAGA = sin pago con 1° vencimiento anterior a la fecha del relevamiento.</div>
<div class="kpis">
  <div class="kpi"><b>${datos.length + (lote.omitidos || []).filter(o => !/repetido/i.test(o.motivo)).length}</b>sociedades</div>
  <div class="kpi"><b>${relevadas.length}</b>revisadas</div>
  <div class="kpi ${noRevisadas.length ? 'r' : ''}"><b>${noRevisadas.length}</b>no se pudieron revisar</div>
  <div class="kpi ${conImpagas.length ? 'r' : ''}"><b>${conImpagas.length}</b>con cuotas impagas (${totCuotasImp} cuota${totCuotasImp === 1 ? '' : 's'})</div>
  <div class="kpi"><b>${totPlanes}</b>planes vigentes</div>
  <div class="kpi"><b>${pesos(saldo).replace('$ ', '$')}</b>saldo de capital</div>
</div>
${planesPend ? `<div class="aviso"><b>CBU y obligaciones:</b> faltan en ${planesPend} de ${totPlanes} planes vigentes (marcados <span class="pend">pendiente</span>).
Mis Facilidades de ARCA dejó de responder esas pantallas durante el relevamiento. </div>` : ''}
${incompletos.length ? `<div class="aviso"><b>Planes con lectura incompleta</b> (ARCA no respondió la pantalla; no se afirma que estén impagos):
${incompletos.map(i => `${esc(i.sociedad)} · plan ${esc(i.plan)} (falta ${i.falta.join(' y ')})`).join('; ')}.</div>` : ''}`;

// 1. No revisadas
h += `<h2 class="gris">1. Sociedades que no se pudieron revisar (${noRevisadas.length})</h2>`;
if (!noRevisadas.length) h += '<p>Todas las sociedades se pudieron revisar.</p>';
else {
  h += '<table><tr><th style="width:4%">#</th><th style="width:22%">Sociedad</th><th style="width:12%">CUIT</th><th>Motivo / qué hacer</th></tr>';
  for (const n of noRevisadas) h += `<tr><td>${esc(n.pos)}</td><td><b>${esc(n.sociedad)}</b></td><td class="mono">${cuitF(n.cuit)}</td><td>${esc(n.motivo)}</td></tr>`;
  h += '</table>';
}

// 2. Con impagas
h += `<h2 class="rojo">2. Sociedades con cuotas impagas (${conImpagas.length}) — riesgo de caída del plan</h2>`;
if (!conImpagas.length) h += '<p>Ningún plan vigente tiene cuotas vencidas sin pagar.</p>';
for (const x of conImpagas) {
  h += `<div class="alerta"><div class="soctit">${esc(x.s.sociedad)} · CUIT ${cuitF(x.s.cuit)}
        <span class="tag r">${x.impagas.length} plan${x.impagas.length > 1 ? 'es' : ''} con cuotas impagas</span></div>
    <table><tr><th>Nro plan</th><th>Cuota</th><th>Vencimiento</th><th>Días de atraso</th><th>Importe</th>
    <th>Débitos rechazados</th><th>CBU declarado</th><th>Obligaciones incluidas (concepto)</th></tr>`;
  for (const p of x.impagas) {
    p.impagas.forEach((c, k) => {
      const seg = c.segundoVencPendiente ? `<br><span class="muted">2° vto ${esc(c.segundoVenc)}: se reintenta el débito</span>` : '';
      h += `<tr class="imp">${k === 0 ? `<td rowspan="${p.impagas.length}"><b>${esc(p.nroPlan)}</b><br><span class="muted">${esc(p.tipo)}</span><br>${p.cuotasImpagas} impaga(s) · ${p.cuotasPagadas} pagada(s) · ${p.cuotasAVencer} a vencer</td>` : ''}
        <td><b>Cuota ${esc(c.cuota)}</b></td><td><b>${esc(c.vencimiento)}</b>${seg}</td>
        <td class="num"><b>${c.diasAtraso}</b></td><td class="num">${pesos(c.importe)}</td>
        <td>${(c.intentosFallidos || []).map(esc).join('<br>') || '—'}</td>
        ${k === 0 ? `<td rowspan="${p.impagas.length}">${cbuF(p)}</td><td rowspan="${p.impagas.length}">${conceptosF(p)}</td>` : ''}</tr>`;
    });
  }
  h += '</table></div>';
}

// 3. Todas
h += `<h2>3. Todas las sociedades — planes vigentes, CBU y obligaciones</h2>`;
for (const x of datos) {
  const { s, vig } = x;
  let tag;
  if (!s.ok) tag = '<span class="tag g">NO REVISADA</span>';
  else if (x.impagas.length) tag = '<span class="tag r">CUOTAS IMPAGAS</span>';
  else if (vig.some(p => p.alerta === 'SIN DATO DE PAGOS')) tag = '<span class="tag g">LECTURA INCOMPLETA</span>';
  else if (vig.length) tag = '<span class="tag v">AL DÍA</span>';
  else tag = '<span class="tag g">SIN PLANES VIGENTES</span>';
  h += `<div class="socbloque"><div class="soctit">#${esc(s.pos)} · ${esc(s.sociedad)} · CUIT ${cuitF(s.cuit)} ${tag}</div>`;
  if (!s.ok) { h += `<div class="muted">No se pudo revisar: ${esc(QUE_HACER[s.code] || s.error)}</div></div>`; continue; }
  if (!vig.length) { h += `<div class="muted">Sin planes vigentes (${x.d.presentaciones.length} presentaciones históricas).</div></div>`; continue; }
  h += `<table><tr><th style="width:8%">Nro plan</th><th style="width:7%">Presentación</th><th style="width:17%">CBU declarado</th>
        <th style="width:30%">Obligaciones incluidas (concepto)</th><th style="width:9%">Cuotas</th>
        <th style="width:10%">Saldo capital</th><th style="width:10%">Próx. vencimiento</th><th>Estado</th></tr>`;
  for (const p of vig) {
    const imp = p.cuotasImpagas > 0;
    const pr = p.proxima;
    const est = imp ? `<span class="tag r">${p.cuotasImpagas} IMPAGA(S)</span><br>${p.impagas.map(c => `cuota ${esc(c.cuota)} vto ${esc(c.vencimiento)}`).join('<br>')}`
      : p.alerta === 'SIN DATO DE PAGOS' ? '<span class="pend">sin dato de pagos</span>'
      : (p.cuotasAVencer ? 'Al día' : 'Al día · todas pagadas');
    h += `<tr class="${imp ? 'imp' : ''}"><td><b>${esc(p.nroPlan)}</b></td><td>${esc(p.presentacion)}</td><td>${cbuF(p)}</td>
      <td>${conceptosF(p)}</td><td>${p.cuotasPagadas}/${p.cantCuotas} pagadas</td>
      <td class="num">${pesos(p.saldoCapital)}</td>
      <td>${pr ? `${esc(pr.vencimiento)}<br><span class="muted">cuota ${esc(pr.cuota)} · ${pesos(pr.importe)}</span>` : '—'}</td>
      <td>${est}</td></tr>`;
  }
  h += '</table></div>';
}
h += `<div class="nota">Sin planes vigentes: ${sinVig} sociedades. Cuotas = pagadas / total del plan (incluye el pago a cuenta).
El detalle cuota por cuota y cada obligación con su período están en INFORME_PLANES_FACILIDADES_${iso.replace(/-/g, '')}.xlsx.
La causal de caducidad depende del régimen de cada plan: confirmarla ante cualquier cuota impaga.</div>
</body></html>`;

(async () => {
  const base = path.join(DIR, `REPORTE_PLANES_FACILIDADES_${iso.replace(/-/g, '')}`);
  fs.writeFileSync(base + '.html', h, 'utf-8');
  const b = await chromium.launch();
  const pg = await b.newPage();
  await pg.setContent(h, { waitUntil: 'load' });
  await pg.pdf({ path: base + '.pdf', format: 'A4', landscape: true, printBackground: true,
                 displayHeaderFooter: true, headerTemplate: '<span></span>',
                 footerTemplate: `<div style="font-size:7pt;color:#777;width:100%;text-align:center">Planes de facilidades · ${fechaTxt} · página <span class="pageNumber"></span> de <span class="totalPages"></span></div>`,
                 margin: { top: '11mm', bottom: '13mm', left: '10mm', right: '10mm' } });
  await b.close();
  console.log(JSON.stringify({ ok: true, pdf: base + '.pdf', html: base + '.html',
    noRevisadas: noRevisadas.length, conImpagas: conImpagas.length, planesVigentes: totPlanes }));
})().catch(e => fail(e.message));
