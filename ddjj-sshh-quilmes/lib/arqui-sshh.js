// arqui-sshh.js — Tasa de Seguridad e Higiene (Regimen General) en ARQui
// (https://arquivirtual.quilmes.gov.ar). El login y el manejo del 403 estan en
// arqui.js; aca van las pantallas de SSHH.
//
// Lo aprendido (22/10/2026):
//   - Consultar = ConsultarMensual.aspx: <select> #ContentPlaceHolder1_cmbCta con
//     las cuentas de SSHH vinculadas al usuario (value = numero SIN digito
//     verificador) y #ContentPlaceHolder1_cmbTipoOp (anio). "Buscar" hace un POST
//     JSON a ConsultarMensual.aspx/WsListar {v_tipo: anio, v_valor: cuenta} que
//     devuelve una fila por DDJJ del anio: ID, COD_SEG, Origen (Mensual/Anual),
//     MESstr, Cuit, CARACTER_PRESENT (ORIGINAL/RECTIFICATIVA), COMPLEMENTO (Nº)
//     y la fecha de presentacion en CTA_NRO_Y_DIGITO (la pagina la muestra como
//     FECHA). Puede venir la misma DDJJ repetida.
//   - En esa pantalla CargarVer(ID, COD_SEG) baja la Declaracion ("DDJJ.pdf":
//     actividades, alicuota, ingreso, tasa) y CargarVer2(ID, COD_SEG) el Recibo
//     de Pago ("Recibo.pdf", la boleta que el equipo guarda en TSH). Son
//     postbacks que devuelven el archivo: la pagina queda igual.
//   - Liquidacion Mensual = DDJJAddSeleccionM.aspx -> DDJJAddSeleccionM_AF.aspx:
//     cmbCuentas (value = id interno, texto = cuenta sin digito; postback),
//     cmbAnio (postback) y cmbMes (texto 1..12, solo hasta el ultimo mes
//     cerrado), boton "Nueva" -> ConsultarPeriodo.aspx (DDJJ del periodo con el
//     INGRESO $ declarado) y ahi "Nuevo" = RgAdd.aspx, el formulario de carga.

const { ArquiError } = require('./arqui');

const BASE = 'https://arquivirtual.quilmes.gov.ar/municipal/App/DDJJ/Sshh';
const URL_CONSULTA = `${BASE}/ConsultarMensual.aspx`;
const SEL_CTA = '#ContentPlaceHolder1_cmbCta';

const esBloqueo = (resp) => Boolean(resp && resp.status() === 403);
const errBloqueo = () => new ArquiError('bloqueo',
  'ARQui contesto 403 (Forbidden): el firewall del municipio bloqueo temporalmente los pedidos por ser muchos seguidos. Hay que esperar unos minutos.');

async function abrirConsulta(page) {
  if (esBloqueo(await page.goto(URL_CONSULTA, { waitUntil: 'domcontentloaded', timeout: 60000 }))) throw errBloqueo();
  if (/Login_arqui/i.test(page.url())) throw new ArquiError('sesion', 'La sesion de ARQui vencio.');
  const sel = await page.waitForSelector(SEL_CTA, { timeout: 30000 }).catch(() => null);
  if (!sel) {
    const texto = (await page.innerText('body').catch(() => '')).replace(/\s+/g, ' ');
    if (/Forbidden/i.test(texto)) throw errBloqueo();
    throw new ArquiError('pantalla', `No aparecio el desplegable de cuentas de Seguridad e Higiene (${texto.slice(0, 200)}).`);
  }
}

const enConsulta = async (page) => page.url().startsWith(URL_CONSULTA) && Boolean(await page.$(SEL_CTA));

// Cuentas de SSHH vinculadas al usuario (sin digito verificador).
async function cuentas(page) {
  if (!(await enConsulta(page))) await abrirConsulta(page);
  return page.$$eval(`${SEL_CTA} option`, (os) => os.map((o) => o.value.trim()).filter(Boolean));
}

// DDJJ de una cuenta en un anio: [{ id, seg, origen, mes, cuit, fecha, caracter, nro }]
async function ddjjDelAnio(page, cuenta, anio) {
  if (!(await enConsulta(page))) await abrirConsulta(page);
  const r = await page.evaluate(async ([c, a]) => {
    const res = await fetch('ConsultarMensual.aspx/WsListar', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: JSON.stringify({ v_tipo: String(a), v_valor: String(c) }),
    });
    return { status: res.status, url: res.url, body: await res.text() };
  }, [cuenta, anio]);
  if (r.status === 403) throw errBloqueo();
  if (/Login_arqui/i.test(r.url)) throw new ArquiError('sesion', 'La sesion de ARQui vencio.');
  let d;
  try { d = JSON.parse(r.body).d; } catch {
    throw new ArquiError('pantalla', `ARQui no devolvio la lista de DDJJ de la cuenta ${cuenta} (${r.status}: ${r.body.replace(/\s+/g, ' ').slice(0, 150)}).`);
  }
  const fechaDe = (x) => [x.CTA_NRO_Y_DIGITO, x.Fecha_aud_str, x.FECHA].find((v) => /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(String(v || '').trim()));
  return (d || []).map((x) => ({
    id: x.ID,
    seg: x.COD_SEG,
    origen: String(x.Origen || '').trim(),
    mes: Number(String(x.MESstr || x.MES || '').trim()) || null,
    cuit: String(x.Cuit || '').replace(/\D/g, ''),
    fecha: fechaDe(x) || null,
    caracter: String(x.CARACTER_PRESENT || '').trim(),
    nro: Number(x.COMPLEMENTO) || 0,
  }));
}

// Baja la Declaracion (cual = 'ddjj') o el Recibo de Pago (cual = 'recibo') de
// una DDJJ de la lista y lo guarda en `destino`.
async function descargar(page, fila, cual, destino) {
  if (!(await enConsulta(page))) await abrirConsulta(page);
  const fn = fila.origen === 'Anual' ? (cual === 'ddjj' ? 'CargarVer3' : 'CargarVer4') : (cual === 'ddjj' ? 'CargarVer' : 'CargarVer2');
  let download;
  try {
    [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 45000 }),
      page.evaluate(([f, id, seg]) => window[f](id, seg), [fn, fila.id, fila.seg]),
    ]);
  } catch {
    if (/Login_arqui/i.test(page.url())) throw new ArquiError('sesion', 'La sesion de ARQui vencio.');
    const texto = (await page.innerText('body').catch(() => '')).replace(/\s+/g, ' ');
    if (/Forbidden/i.test(texto)) throw errBloqueo();
    throw new ArquiError('sin_pdf', `ARQui no devolvio ${cual === 'ddjj' ? 'la Declaracion' : 'el Recibo de Pago'} (${texto.replace(/.*CARACTER_PRESENT/, '').slice(0, 150)}).`);
  }
  await download.saveAs(destino);
  return destino;
}

module.exports = { abrirConsulta, cuentas, ddjjDelAnio, descargar, URL_CONSULTA };
