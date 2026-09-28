// facilidades.js — Mis Facilidades de ARCA: planes de pago, cuotas y pagos.
//
// Circuito (verificado contra ARCA el 24/09/2026):
//   portal → tile "Mis Facilidades" (pestaña nueva, serviciossegsoc…/MisFacilidadesNet)
//   → login/IndexContribuyente.aspx: select #ContentPlaceHolder1_ddlCUIT + Aceptar
//   → contribuyente/seguimiento_presentacion.aspx: "Presentaciones Enviadas"
//   → boton Detalle de la fila → seguimiento/nuevos_planes.aspx
//       → "Ver Pagos"   → seguimiento/detalle_pagos.aspx      (lo pagado / intentado)
//       → "Plan de Pago"→ seguimiento/detalle_plan_pago.aspx  (cronograma completo)
//
// Cosas de ARCA que no son obvias:
// - Para cambiar de CUIT NO sirve volver a IndexContribuyente.aspx por URL:
//   responde ErrorPage "No puede acceder a la pagina solicitada" (403). Hay que
//   cerrar la pestaña y volver a abrir el servicio desde el portal.
// - El paginado de las tablas es del lado del cliente (numperpage): TODAS las
//   filas estan en el DOM, ocultas con display:none. Se leen por DOM y no hace
//   falta clickear "Siguiente".
// - En "Ver Pagos" una cuota ocupa varias filas: la primera trae Cuota N°,
//   Capital y Estado (con rowspan) y las siguientes son los reintentos de debito
//   con interes resarcitorio. La columna Pago dice "Pago" en el intento que se
//   cobro; en los fallidos hay un boton cuyo tooltip es el motivo ("falta de
//   fondos", "pago de vep no registrado").
// - En "Plan de Pago" el capital viene sin separador de miles ("3044635,95").
// - El CBU declarado esta en "Datos del Plan" (#ContentPlaceHolder1_tdCbu). El
//   "Historial de Cambios de CBU" trae los cambios con su fecha de vigencia.
//   OJO: esa pantalla tiene un boton "Cambiar CBU" — no se toca nunca.
// - Las obligaciones estan en "Obligaciones Impositivas" (btnObligImp) o en
//   "Obligaciones Previsionales" (btnObligPrev) segun el plan; puede haber ambas.
//
// REGLA DURA: esto solo LEE. Nunca presenta, modifica, reformula ni paga nada.

const log = (m) => process.stderr.write(`[facilidades] ${m}\n`);
const BASE = 'https://serviciossegsoc.afip.gob.ar/tramites_con_clave_fiscal/MisFacilidadesNet/app';

const soloDig = (s) => String(s || '').replace(/\D/g, '');

// '3.044.635,95' | '3044635,95' | '$ 29.816.587,57' | '-' -> numero o null.
function aNumero(s) {
  const t = String(s || '').replace(/[^\d,.-]/g, '');
  if (!t || t === '-') return null;
  const n = Number(t.replace(/\./g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

function aFecha(s) {
  const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}

async function esperar(p) {
  await p.waitForLoadState('domcontentloaded', { timeout: 40000 }).catch(() => {});
  await p.waitForLoadState('networkidle', { timeout: 40000 }).catch(() => {});
  await p.waitForTimeout(800);
}

function errorPagina(txt) {
  return /No puede acceder a la p.gina solicitada/i.test(txt);
}

// Clickea algo que dispara un postback / navegacion y espera la pagina nueva.
// Se clickea por DOM: los botones de filas paginadas estan ocultos y Playwright
// no los clickea con .click() normal.
async function navegar(mf, locator, urlEsperada) {
  if (process.env.FAC_DEBUG) log(`   → ${urlEsperada} (desde ${mf.url().split('/').pop()})`);
  if (!(await locator.count())) {
    const e = new Error(`No encontre el boton/enlace para ir a ${urlEsperada} en ${mf.url()}`);
    e.code = 'boton_ausente';
    throw e;
  }
  await Promise.all([
    mf.waitForURL(urlEsperada, { timeout: 45000 }),
    locator.evaluate(e => e.click()),
  ]);
  if (process.env.FAC_DEBUG) log(`   navegado a ${mf.url().split('/').pop()}; espero carga`);
  await esperar(mf);
  const txt = await mf.innerText('body').catch(() => '');
  if (errorPagina(txt)) {
    const e = new Error(`ARCA devolvio "No puede acceder a la pagina solicitada" en ${mf.url()}`);
    e.code = 'pagina_error';
    throw e;
  }
  return txt;
}

// Que contribuyente muestra el encabezado ("Contribuyente:30-00000000-0 - RAZON SOCIAL").
async function contribuyenteActivo(mf) {
  const txt = await mf.innerText('body').catch(() => '');
  const m = txt.match(/Contribuyente:\s*([\d-]{11,13})\s*-\s*([^\n]+)/i);
  return m ? { cuit: soloDig(m[1]), razon: m[2].trim() } : { cuit: null, razon: null };
}

// Abre Mis Facilidades en pestaña nueva y la deja parada en el CUIT pedido,
// en la pantalla de Presentaciones Enviadas. Verifica el CUIT del encabezado.
async function abrirEnCuit(context, portal, cuit, { intentos = 2 } = {}) {
  const { abrirServicio } = require('./arca-login');
  const objetivo = soloDig(cuit);
  let ultimo;
  for (let i = 1; i <= intentos; i++) {
    let mf = null;
    try {
      mf = await abrirServicio(context, portal, /Mis Facilidades/i);
      await esperar(mf);
      if (/auth\.afip\.gob\.ar/i.test(mf.url())) {
        const e = new Error('Mis Facilidades rebota al login'); e.code = 'rebota_login'; throw e;
      }

      const ddl = mf.locator('#ContentPlaceHolder1_ddlCUIT');
      if (await ddl.isVisible({ timeout: 8000 }).catch(() => false)) {
        const valor = await ddl.evaluate((s, c) => {
          const o = [...s.options].find(x => (x.value + ' ' + x.text).replace(/\D/g, '').includes(c));
          return o ? o.value : null;
        }, objetivo);
        if (valor == null) {
          const e = new Error(`El CUIT ${objetivo} no figura entre las CUITs asociadas de Mis Facilidades.`);
          e.code = 'cuit_no_disponible';
          throw e;
        }
        await ddl.selectOption({ value: valor });
        await navegar(mf, mf.locator('#ContentPlaceHolder1_btnAceptar'), /seguimiento_presentacion\.aspx/i);
      } else if (!/seguimiento_presentacion/i.test(mf.url())) {
        // Login con CUIT propio: puede entrar directo o tener otro paso.
        log(`sin selector de CUIT; url ${mf.url()}`);
        await mf.goto(`${BASE}/contribuyente/seguimiento_presentacion.aspx`).catch(() => {});
        await esperar(mf);
      }

      const activo = await contribuyenteActivo(mf);
      if (activo.cuit !== objetivo) {
        const e = new Error(`Mis Facilidades muestra ${activo.cuit || '?'} (${activo.razon || 'sin encabezado'}) ` +
                            `en vez de ${objetivo}. Corto para no relevar otra sociedad.`);
        e.code = 'contribuyente_incorrecto';
        throw e;
      }
      log(`contribuyente confirmado: ${activo.cuit} - ${activo.razon}`);
      return { mf, razon: activo.razon };
    } catch (e) {
      ultimo = e;
      if (mf) await mf.close().catch(() => {});
      if (e.code === 'cuit_no_disponible' || i === intentos) throw e;
      log(`intento ${i}: ${e.message.split('\n')[0]} — vuelvo al portal y reintento`);
      await portal.goto('https://portalcf.cloud.afip.gob.ar/portal/app/', { waitUntil: 'domcontentloaded' }).catch(() => {});
      await portal.waitForTimeout(3000);
    }
  }
  throw ultimo;
}

// Tabla de Presentaciones Enviadas. Todas las filas, esten o no en la pagina visible.
async function leerPresentaciones(mf) {
  const filas = await mf.evaluate(() => {
    const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
    const tabla = [...document.querySelectorAll('table')]
      .find(t => /Consolidado/i.test(t.tHead ? t.tHead.innerText : '') && /Situaci/i.test(t.tHead ? t.tHead.innerText : ''));
    if (!tabla) return null;
    return [...tabla.tBodies].flatMap(b => [...b.rows]).map(r => {
      const c = [...r.cells].map(x => norm(x.textContent));
      const btn = r.querySelector('input[id*="detallePlan"]');
      return {
        presentacion: c[0], numero: c[1], cuotas: c[2], tipo: c[3],
        consolidado: c[4], estado: c[5], situacion: c[6],
        botonId: btn ? btn.id : null,
      };
    }).filter(f => /\d{2}\/\d{2}\/\d{4}/.test(f.presentacion));
  });
  if (filas === null) {
    const e = new Error('No encontre la tabla "Presentaciones Enviadas".');
    e.code = 'estructura_desconocida';
    throw e;
  }
  return filas;
}

async function datosDelPlan(mf) {
  const txt = await mf.innerText('body').catch(() => '');
  const cap = (re) => { const m = txt.match(re); return m ? m[1].trim() : ''; };
  return {
    descripcion: cap(/Descripci.n:\s*([^\n]+)/i),
    nroPlan: cap(/Nro de plan:\s*([^\n]+)/i),
    fechaConsolidacion: cap(/Fecha de consolidaci.n:\s*([^\n]+)/i),
    tipoPlan: cap(/Tipo de plan:\s*([^\n]+)/i),
    perfil: cap(/Perfil:\s*([^\n]+)/i),
  };
}

// "Ver Pagos": cuotas con sus intentos de cobro.
async function leerPagos(mf) {
  return mf.evaluate(() => {
    const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
    const tabla = document.querySelector('#tableDetallePagos') ||
      [...document.querySelectorAll('table')].find(t => /Estado de Cuota/i.test(t.innerText));
    if (!tabla) return null;
    const cuotas = [];
    let totalPagado = null;
    for (const r of tabla.querySelectorAll('tbody tr')) {
      const tds = [...r.cells];
      if (/Total Pagado/i.test(r.innerText)) {
        const v = tds.map(x => norm(x.textContent));
        totalPagado = { capital: v[1], intFinanciero: v[2], intResarcitorio: v[3], total: v[4] };
        continue;
      }
      const tdCuota = r.querySelector('td[id*="tdCuotaNro"]');
      // Celdas visibles de datos del intento (los reintentos traen 2 celdas
      // ocultas al principio y 1 al final).
      const vis = tds.filter(x => x.style.display !== 'none');
      let cuota = null, capital = null, estado = null, datos;
      if (tdCuota) {
        cuota = norm(tdCuota.textContent);
        capital = norm(vis[1].textContent);
        const tdEst = r.querySelector('td[id*="tdEstadoCuota"]');
        estado = tdEst ? norm(tdEst.textContent) : '';
        datos = vis.slice(2, 7);
      } else {
        datos = vis.slice(0, 5);
      }
      if (datos.length < 5) continue;
      const tdPago = datos[4];
      const motivoEl = tdPago.querySelector('[title], [data-original-title]');
      const intento = {
        intFinanciero: norm(datos[0].textContent),
        intResarcitorio: norm(datos[1].textContent),
        total: norm(datos[2].textContent),
        fecha: norm(datos[3].textContent),
        pagado: /^pago$/i.test(norm(tdPago.textContent)),
        motivo: motivoEl ? (motivoEl.getAttribute('title') || motivoEl.getAttribute('data-original-title') || '') : norm(tdPago.textContent),
      };
      if (cuota !== null) cuotas.push({ cuota, capital, estado, intentos: [intento] });
      else if (cuotas.length) cuotas[cuotas.length - 1].intentos.push(intento);
    }
    return { cuotas, totalPagado };
  });
}

// "Plan de Pago": cronograma completo (1° vencimiento y, si hay, 2°).
async function leerCronograma(mf) {
  return mf.evaluate(() => {
    const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
    const tabla = document.querySelector('#tableDetallePlanPago') ||
      [...document.querySelectorAll('table')].find(t => /Fecha Vencimiento/i.test(t.innerText));
    if (!tabla) return null;
    const cuotas = [];
    for (const r of tabla.querySelectorAll('tr[id*="rptDetallePlanPago_trRpt_"]')) {
      const q = (k) => r.querySelector(`td[id*="${k}"]`);
      const tdCuota = q('tdCuota_');
      const principal = tdCuota && tdCuota.style.display !== 'none';
      const v = {
        intFinanciero: norm((q('tdInteresFinanciero_') || {}).textContent),
        intResarcitorio: norm((q('tdInteresResarcitorio_') || {}).textContent),
        total: norm((q('tdTotal_') || {}).textContent),
        vencimiento: norm((q('tdFechaVencimiento_') || {}).textContent),
      };
      if (principal) {
        cuotas.push({ cuota: norm(tdCuota.textContent), capital: norm((q('tdCapital_') || {}).textContent),
                      venc1: v, venc2: null });
      } else if (cuotas.length) {
        cuotas[cuotas.length - 1].venc2 = v;
      }
    }
    const t = (id) => { const e = document.getElementById(id); return e ? norm(e.textContent) : ''; };
    return { cuotas, totales: { capital: t('ContentPlaceHolder1_tdCapitalTotal'),
                                intFinanciero: t('ContentPlaceHolder1_tdInteresFinancieroTotal1'),
                                total: t('ContentPlaceHolder1_tdTotalTotal1') } };
  });
}

// "Datos del Plan": CBU declarado.
async function leerCbu(mf) {
  return mf.evaluate(() => {
    const e = document.getElementById('ContentPlaceHolder1_tdCbu');
    return e ? e.textContent.replace(/\s+/g, '') : '';
  });
}

// "Historial de Cambios de CBU".
async function leerHistorialCbu(mf) {
  return mf.evaluate(() => {
    const t = document.getElementById('tableHistorial_CBU');
    if (!t) return null;
    return [...t.querySelectorAll('tbody tr')].map(r => {
      const q = (k) => { const e = r.querySelector(`td[id*="${k}"]`); return e ? e.textContent.trim() : ''; };
      return { origen: q('tdOrigen'), presentacion: q('tdFechaPresentacion'),
               vigencia: q('tdFechaVigencia'), cbu: q('tdCBU').replace(/\s+/g, '') };
    }).filter(x => x.cbu);
  });
}

// "Obligaciones Impositivas / Previsionales": se mapea por encabezado porque
// las dos pantallas no tienen exactamente las mismas columnas.
async function leerObligaciones(mf) {
  return mf.evaluate(() => {
    const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
    const out = [], vistos = new Set();
    for (const t of document.querySelectorAll('table')) {
      if (!t.tHead) continue;
      const hs = [...t.tHead.querySelectorAll('th')].map(x => norm(x.textContent));
      if (!hs.some(h => /^(Cpto|Concepto)/i.test(h))) continue;
      const idx = (re) => hs.findIndex(h => re.test(h));
      const iImp = idx(/^Imp(\.|uesto)/i), iCpto = idx(/^(Cpto|Concepto)/i), iSub = idx(/^Sub/i);
      const iAnio = idx(/^A.o/i), iMes = idx(/^Mes/i), iVto = idx(/Vto|Venc/i), iPer = idx(/Per.odo/i);
      for (const r of t.querySelectorAll('tbody tr')) {
        const c = [...r.cells].map(x => norm(x.textContent));
        if (c.length < hs.length - 1) continue;
        const g = (i) => (i >= 0 ? c[i] || '' : '');
        const o = { impuesto: g(iImp), concepto: g(iCpto), subconcepto: g(iSub),
                    anio: g(iAnio), mes: g(iMes), periodo: g(iPer), vencimiento: g(iVto) };
        const k = JSON.stringify(o);
        if ((o.impuesto || o.concepto) && !vistos.has(k)) { vistos.add(k); out.push(o); }
      }
    }
    return out;
  });
}

// Entra a una pantalla del detalle, lee y vuelve a nuevos_planes.aspx.
async function leerSubpantalla(mf, botonId, urlRe, lector) {
  await navegar(mf, mf.locator(`#${botonId}`), urlRe);
  if (process.env.FAC_DEBUG) log(`   llegue a ${mf.url().split('/').pop()}; leo`);
  const r = await lector(mf);
  if (process.env.FAC_DEBUG) log(`   leido; vuelvo`);
  await navegar(mf, mf.locator('a[href="nuevos_planes.aspx"]').first(), /nuevos_planes\.aspx/i);
  return r;
}

// Entra al Detalle de una presentacion, lee Pagos y Plan de Pago, y vuelve a
// la lista. Deja la pestaña en seguimiento_presentacion.aspx.
async function leerPlan(mf, fila) {
  const obs = [];
  if (!fila.botonId) throw Object.assign(new Error(`La fila ${fila.numero} no tiene boton Detalle`), { code: 'sin_detalle' });

  await navegar(mf, mf.locator(`#${fila.botonId}`), /nuevos_planes\.aspx/i);
  const datos = await datosDelPlan(mf);
  if (datos.nroPlan && datos.nroPlan !== fila.numero) {
    throw Object.assign(new Error(`Abri el plan ${datos.nroPlan} pero pedi ${fila.numero}`), { code: 'plan_incorrecto' });
  }

  let pagos = null, cronograma = null;
  try {
    await navegar(mf, mf.locator('#ContentPlaceHolder1_btnVerPagos'), /detalle_pagos\.aspx/i);
    pagos = await leerPagos(mf);
    if (!pagos) obs.push('No encontre la tabla de "Ver Pagos".');
    await navegar(mf, mf.locator('a[href="nuevos_planes.aspx"]').first(), /nuevos_planes\.aspx/i);
  } catch (e) {
    obs.push(`Ver Pagos no cargo: ${e.message.split('\n')[0]}`);
    if (!/nuevos_planes/i.test(mf.url())) await mf.goBack().catch(() => {});
    await esperar(mf);
  }

  try {
    await navegar(mf, mf.locator('#ContentPlaceHolder1_btnPlanPago'), /detalle_plan_pago\.aspx/i);
    cronograma = await leerCronograma(mf);
    if (!cronograma) obs.push('No encontre la tabla de "Plan de Pago".');
    await navegar(mf, mf.locator('a[href="nuevos_planes.aspx"]').first(), /nuevos_planes\.aspx/i);
  } catch (e) {
    obs.push(`Plan de Pago no cargo: ${e.message.split('\n')[0]}`);
  }

  // CBU y obligaciones. Si una pantalla falla, se anota y se sigue: no es
  // motivo para descartar la lectura de cuotas.
  let cbu = '', historialCbu = null, obligaciones = [];
  const intentar = async (nombre, fn) => {
    try { return await fn(); }
    catch (e) {
      obs.push(`${nombre} no cargo: ${e.message.split('\n')[0]}`);
      if (!/nuevos_planes/i.test(mf.url())) {
        await mf.locator('a[href="nuevos_planes.aspx"]').first().evaluate(x => x.click()).catch(() => mf.goBack());
        await esperar(mf);
      }
      return null;
    }
  };
  cbu = (await intentar('Datos del Plan', () =>
    leerSubpantalla(mf, 'ContentPlaceHolder1_btnDatosPlan', /datos_plan\.aspx/i, leerCbu))) || '';
  if (await mf.locator('#ContentPlaceHolder1_btnHisCBU').count()) {
    historialCbu = await intentar('Historial de CBU', () =>
      leerSubpantalla(mf, 'ContentPlaceHolder1_btnHisCBU', /modifica_forma_pago\.aspx/i, leerHistorialCbu));
  }
  let hayOblig = false;
  for (const [id, tipo] of [['ContentPlaceHolder1_btnObligImp', 'Impositiva'], ['ContentPlaceHolder1_btnObligPrev', 'Previsional']]) {
    if (!(await mf.locator(`#${id}`).count())) continue;
    hayOblig = true;
    const r = await intentar(`Obligaciones ${tipo}s`, () =>
      leerSubpantalla(mf, id, /^(?!.*nuevos_planes).*\.aspx/i, leerObligaciones));
    if (r) obligaciones.push(...r.map(o => ({ tipo, ...o })));
  }
  if (!hayOblig) obs.push('El detalle del plan no tiene boton de Obligaciones.');
  else if (!obligaciones.length) obs.push('No pude leer las obligaciones incluidas en el plan.');
  if (!cbu) obs.push('No pude leer el CBU declarado.');

  // Volver a la lista de presentaciones.
  const volver = mf.locator('a[href*="seguimiento_presentacion.aspx"]').first();
  if (await volver.count()) {
    await navegar(mf, volver, /seguimiento_presentacion\.aspx/i);
  } else {
    await mf.goto(`${BASE}/contribuyente/seguimiento_presentacion.aspx`);
    await esperar(mf);
  }
  return { datos, pagos, cronograma, cbu, historialCbu, obligaciones, observaciones: obs };
}

// ── Analisis: cruza cronograma y pagos, y marca las cuotas impagas ─────────
//
// Cuota pagada: el Estado de Cuota dice "Cancelada" o algun intento dice "Pago".
// Cuota IMPAGA: no esta pagada y su 1° vencimiento es anterior a hoy.
// Cuota a vencer: 1° vencimiento hoy o despues.
function analizar(plan, hoy = new Date()) {
  const inicioHoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  const obs = [...(plan.observaciones || [])];
  const pagosPor = new Map((plan.pagos ? plan.pagos.cuotas : []).map(c => [c.cuota, c]));
  const cronoPor = new Map((plan.cronograma ? plan.cronograma.cuotas : []).map(c => [c.cuota, c]));
  const claves = [...new Set([...cronoPor.keys(), ...pagosPor.keys()])];
  const orden = (k) => (/cuenta/i.test(k) ? -1 : Number(k) || 9999);
  claves.sort((a, b) => orden(a) - orden(b));

  const cuotas = claves.map(k => {
    const cr = cronoPor.get(k), pg = pagosPor.get(k);
    const venc1 = (cr && cr.venc1.vencimiento) || (pg && pg.intentos[0] && pg.intentos[0].fecha) || '';
    const venc2 = cr && cr.venc2 ? cr.venc2.vencimiento : '';
    const pagoOk = pg && pg.intentos.find(x => x.pagado);
    const cancelada = pg && /cancelad/i.test(pg.estado || '');
    // Solo cuentan los debitos cuya fecha ya paso: el 2° vencimiento futuro
    // aparece en la tabla sin "Pago" pero todavia no se intento.
    const fallidos = pg ? pg.intentos.filter(x => !x.pagado && aFecha(x.fecha) && aFecha(x.fecha) < inicioHoy) : [];
    const fv1 = aFecha(venc1), fv2 = aFecha(venc2);
    let estado;
    if (pagoOk || cancelada) estado = 'Pagada';
    // Sin la pantalla Ver Pagos no se puede afirmar que una cuota vencida este
    // impaga: se deja sin dato (caso real: un timeout de Ver Pagos daba un falso "impaga").
    else if (!plan.pagos && fv1 && fv1 < inicioHoy) estado = 'Sin dato';
    else if (fv1 && fv1 < inicioHoy) estado = 'IMPAGA';
    else if (fv1) estado = 'A vencer';
    else estado = 'Sin dato';
    const fechaPago = pagoOk ? pagoOk.fecha : '';
    const fp = aFecha(fechaPago);
    const diasAtraso = estado === 'IMPAGA' ? Math.round((inicioHoy - fv1) / 86400000) : null;
    return {
      cuota: k,
      capital: aNumero((cr && cr.capital) || (pg && pg.capital)),
      intFinanciero: aNumero(cr ? cr.venc1.intFinanciero : pg && pg.intentos[0].intFinanciero),
      total: aNumero(cr ? cr.venc1.total : pg && pg.intentos[0].total),
      venc1, venc2,
      totalVenc2: cr && cr.venc2 ? aNumero(cr.venc2.total) : null,
      estado,
      estadoArca: pg ? pg.estado : '',
      fechaPago,
      pagadoTotal: pagoOk ? aNumero(pagoOk.total) : null,
      intResarcitorioPagado: pagoOk ? aNumero(pagoOk.intResarcitorio) : null,
      fueraDeTermino: !!(fp && fv1 && fp > fv1),
      intentosFallidos: fallidos.map(x => `${x.fecha}${x.motivo ? ' (' + x.motivo + ')' : ''}`),
      diasAtraso,
      segundoVencPendiente: estado === 'IMPAGA' && !!(fv2 && fv2 >= inicioHoy),
      soloEnPagos: !cr && !!pg, soloEnCronograma: !!cr && !pg,
    };
  });

  const pagadas = cuotas.filter(c => c.estado === 'Pagada');
  const impagas = cuotas.filter(c => c.estado === 'IMPAGA');
  const aVencer = cuotas.filter(c => c.estado === 'A vencer');
  const sum = (arr, k) => arr.reduce((a, c) => a + (c[k] || 0), 0);
  const consolidado = typeof plan.fila.consolidado === 'number' ? plan.fila.consolidado : aNumero(plan.fila.consolidado);
  const capitalPagado = sum(pagadas, 'capital');

  // Controles
  if (plan.cronograma && consolidado != null) {
    const capCrono = sum(cuotas.filter(c => !c.soloEnPagos), 'capital');
    if (Math.abs(capCrono - consolidado) > 1) {
      obs.push(`La suma del capital del cronograma (${capCrono.toFixed(2)}) no iguala el consolidado (${consolidado.toFixed(2)}).`);
    }
  }
  if (plan.pagos && plan.pagos.totalPagado) {
    const tp = aNumero(plan.pagos.totalPagado.total);
    const suma = sum(pagadas, 'pagadoTotal');
    if (tp != null && Math.abs(tp - suma) > 1) {
      obs.push(`La suma de las cuotas pagadas (${suma.toFixed(2)}) no coincide con el Total Pagado de ARCA (${tp.toFixed(2)}).`);
    }
  }
  if (!plan.cronograma) obs.push('Sin cronograma: las cuotas futuras pueden faltar.');
  for (const c of cuotas) {
    if (c.soloEnCronograma && c.estado !== 'A vencer') obs.push(`Cuota ${c.cuota}: vencida y sin registro en Ver Pagos.`);
    if (c.fueraDeTermino) obs.push(`Cuota ${c.cuota}: pagada fuera de termino el ${c.fechaPago} (vencia ${c.venc1}).`);
  }
  if (!plan.pagos) obs.push('No se pudo leer Ver Pagos: el estado de las cuotas vencidas queda sin dato.');
  else if (!pagadas.length) obs.push('El plan no registra pagos todavia.');

  let alerta = 'OK';
  if (impagas.length >= 2) alerta = 'RIESGO DE CADUCIDAD';
  else if (impagas.length === 1) alerta = 'CUOTA IMPAGA';
  else if (!plan.pagos) alerta = 'SIN DATO DE PAGOS';
  const lecturaIncompleta = [!plan.pagos && 'pagos', !plan.cronograma && 'cronograma'].filter(Boolean);
  const proxima = aVencer[0] || null;

  return {
    nroPlan: plan.fila.numero,
    presentacion: plan.fila.presentacion,
    tipo: plan.fila.tipo,
    tipoPlan: plan.datos.tipoPlan,
    fechaConsolidacion: plan.datos.fechaConsolidacion,
    cantCuotas: Number(plan.fila.cuotas) || cuotas.length,
    consolidado,
    situacion: plan.fila.situacion,
    cuotasPagadas: pagadas.length,
    cuotasImpagas: impagas.length,
    cuotasAVencer: aVencer.length,
    capitalPagado,
    interesesPagados: pagadas.reduce((a, c) => a + ((c.pagadoTotal || 0) - (c.capital || 0)), 0),
    totalPagado: sum(pagadas, 'pagadoTotal'),
    totalPagadoArca: plan.pagos && plan.pagos.totalPagado ? aNumero(plan.pagos.totalPagado.total) : null,
    saldoCapital: consolidado != null ? consolidado - capitalPagado : null,
    deudaImpaga: sum(impagas, 'total'),
    impagas: impagas.map(c => ({ cuota: c.cuota, vencimiento: c.venc1, segundoVenc: c.venc2,
                                 importe: c.total, diasAtraso: c.diasAtraso,
                                 intentosFallidos: c.intentosFallidos,
                                 segundoVencPendiente: c.segundoVencPendiente })),
    proxima: proxima ? { cuota: proxima.cuota, vencimiento: proxima.venc1, importe: proxima.total } : null,
    fueraDeTermino: cuotas.filter(c => c.fueraDeTermino).length,
    alerta,
    lecturaIncompleta,
    // Solo si se leyeron: un relevamiento viejo sin estas pantallas tiene que
    // verse como "pendiente", no como "sin CBU".
    ...(plan.cbu !== undefined ? {
      cbuDeclarado: plan.cbu || '',
      cbuVigente: cbuVigente(plan.historialCbu, plan.cbu, inicioHoy),
      historialCbu: plan.historialCbu || [],
    } : {}),
    ...(plan.obligaciones !== undefined ? {
      obligaciones: plan.obligaciones,
      conceptos: conceptos(plan.obligaciones),
    } : {}),
    cuotas,
    observaciones: obs,
  };
}

// CBU que rige hoy segun el historial (ultimo con vigencia <= hoy). Si todos
// rigen a futuro o no hay historial, el declarado.
function cbuVigente(historial, declarado, hoy) {
  const vig = (historial || []).filter(h => aFecha(h.vigencia) && aFecha(h.vigencia) <= hoy)
    .sort((a, b) => aFecha(a.vigencia) - aFecha(b.vigencia));
  return vig.length ? vig[vig.length - 1].cbu : (declarado || '');
}

// "Ganancias Personas Fisicas - Declaración Jurada (2024)": una linea por
// impuesto + concepto, con los periodos juntos. Sin importes.
function conceptos(obligaciones) {
  const m = new Map();
  for (const o of obligaciones) {
    const k = [o.impuesto, o.concepto].filter(Boolean).join(' - ');
    if (!m.has(k)) m.set(k, new Set());
    const per = o.periodo || [o.anio, o.mes].filter(Boolean).join('/') ;
    if (per) m.get(k).add(per);
  }
  return [...m].map(([k, pers]) => {
    const p = [...pers];
    return p.length ? `${k} (${p.length > 6 ? `${p.length} períodos, ${p[0]} a ${p[p.length - 1]}` : p.join(', ')})` : k;
  });
}

module.exports = { BASE, abrirEnCuit, leerPresentaciones, leerPlan, analizar, aNumero, aFecha, contribuyenteActivo };
