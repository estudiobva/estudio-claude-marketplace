// mis-comprobantes.js — "Mis Comprobantes" de ARCA (fes.afip.gob.ar/mcmp).
//
// Basado en scripts/lib/mis-comprobantes.js del plugin fisco-ar (MIT), adaptado
// al login por apoderado del estudio: la mayoria de las sociedades entran con la
// clave de un apoderado y Mis Comprobantes abre con "Elegí una persona para
// ingresar" en vez de ir directo al menu.
//
// El servicio no devuelve los comprobantes en el HTML: ajax.do?f=generarConsulta
// encola una consulta que responde un idConsulta, el server la procesa
// (estado PE -> PR) y recien ahi descargarComprobantes.do baja un ZIP con el CSV.
//
// Desde 01/01/2026 trae tambien los tiques de los controladores fiscales de
// nueva tecnologia, uno por uno (tipo 83 / 81 / 82 / 110...), igual que el
// Libro IVA Ventas del Portal IVA.
//
// Solo lectura: no hay nada en este servicio que modifique comprobantes.

const MC_URL = 'https://fes.afip.gob.ar/mcmp/jsp/';

const log = (m) => process.stderr.write(`[mis-comprobantes] ${m}\n`);

const fmtCuit = (c) => String(c).replace(/\D/g, '').replace(/^(\d{2})(\d{8})(\d)$/, '$1-$2-$3');

// CUIT del encabezado: "REPRESENTANDO A: RAZON SOCIAL [30-12345678-9]" o, si se
// entro con la clave propia, el de la linea del usuario.
async function cuitEnEncabezado(page) {
  const txt = await page.innerText('body').catch(() => '');
  const rep = txt.split('\n').map((s) => s.trim()).find((l) => /REPRESENTANDO A/i.test(l));
  const linea = rep || txt.split('\n').map((s) => s.trim()).find((l) => /\[\d{2}-\d{8}-\d\]/.test(l)) || '';
  const m = linea.match(/\[(\d{2})-(\d{8})-(\d)\]/);
  return { linea, cuit: m ? m.slice(1).join('') : null };
}

// Abre Mis Comprobantes desde el portal (ya logueado) y elige la sociedad.
// Devuelve la pagina del servicio, parada en el menu principal de la sociedad.
async function abrir(context, portalPage, cuit) {
  const { abrirServicio } = require('./arca-login');
  const mc = await abrirServicio(context, portalPage, /Mis Comprobantes/i);
  await mc.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});

  const objetivo = String(cuit).replace(/\D/g, '');
  const txt = await mc.innerText('body').catch(() => '');
  if (/Eleg[ií] una persona/i.test(txt)) {
    const tarjeta = mc.locator('a.panel', { has: mc.locator('small', { hasText: fmtCuit(objetivo) }) });
    if (!(await tarjeta.count())) {
      const e = new Error(`El usuario no representa al CUIT ${fmtCuit(objetivo)} en Mis Comprobantes (no esta en "Elegí una persona").`);
      e.code = 'sin_representacion';
      throw e;
    }
    await Promise.all([
      mc.waitForNavigation({ timeout: 30000 }).catch(() => {}),
      tarjeta.first().click(),
    ]);
    await mc.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
  }

  const enc = await cuitEnEncabezado(mc);
  if (enc.cuit !== objetivo) {
    const e = new Error(`Mis Comprobantes quedo en otro contribuyente: "${enc.linea}" (esperaba ${fmtCuit(objetivo)}).`);
    e.code = 'representacion_incorrecta';
    throw e;
  }
  log(`${enc.linea}`);
  return mc;
}

// Ultimo dia del mes, para el rango de la consulta.
function rangoMes(periodoAAAAMM) {
  const a = Number(periodoAAAAMM.slice(0, 4)), m = Number(periodoAAAAMM.slice(4));
  const ultimo = new Date(a, m, 0).getDate();
  const mm = String(m).padStart(2, '0');
  return { desde: `01/${mm}/${a}`, hasta: `${String(ultimo).padStart(2, '0')}/${mm}/${a}` };
}

// Consulta emitidos ('E') o recibidos ('R') de un rango DD/MM/AAAA y baja el ZIP.
// Devuelve { zip, idConsulta, cantidad }.
async function descargar(mc, { tipo = 'E', desde, hasta, destinoZip }) {
  const pagina = tipo === 'R' ? 'comprobantesRecibidos.do' : 'comprobantesEmitidos.do';
  await mc.goto(MC_URL + pagina, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await mc.waitForSelector('#fechaEmision', { state: 'visible', timeout: 20000 });

  // El handler de Buscar lee #fechaEmision como "DD/MM/AAAA - DD/MM/AAAA".
  await mc.evaluate(({ d, h }) => {
    const $el = window.$ && window.$('#fechaEmision');
    const dp = $el && $el.data && $el.data('daterangepicker');
    if (dp) { dp.setStartDate(d); dp.setEndDate(h); }
    document.querySelector('#fechaEmision').value = `${d} - ${h}`;
  }, { d: desde, h: hasta });

  let idConsulta = null, cantidad = null, estado = null, errorConsulta = null;
  const escuchar = async (resp) => {
    if (!resp.url().includes('ajax.do')) return;
    const d = await resp.json().catch(() => null);
    if (!d) return;
    if (d.estado && d.estado !== 'ok') errorConsulta = JSON.stringify(d).slice(0, 300);
    const datos = d.datos || {};
    if (datos.idConsulta) idConsulta = String(datos.idConsulta);
    if (datos.consulta) {
      if (typeof datos.consulta.cantidadResultados === 'number') cantidad = datos.consulta.cantidadResultados;
      if (datos.consulta.estado) estado = datos.consulta.estado;
      if (datos.consulta.error) errorConsulta = String(datos.consulta.error);
    }
  };
  mc.on('response', escuchar);
  try {
    log(`consultando ${tipo === 'R' ? 'recibidos' : 'emitidos'} ${desde} a ${hasta}…`);
    await mc.locator('#buscarComprobantes').click();

    // generarConsulta -> idConsulta; listaResultados -> cantidad y estado PR.
    const limite = Date.now() + 180000;
    while (Date.now() < limite && !(idConsulta && cantidad !== null && estado === 'PR')) {
      if (errorConsulta) throw new Error(`Mis Comprobantes devolvio un error: ${errorConsulta}`);
      const sinResultados = await mc.locator('text=/no se encontraron|sin resultados|no hay comprobantes/i').count().catch(() => 0);
      if (idConsulta && sinResultados) { cantidad = 0; break; }
      await mc.waitForTimeout(2000);
    }
    if (!idConsulta) throw new Error('Mis Comprobantes no devolvio el id de la consulta (timeout).');
    if (cantidad === null) throw new Error(`La consulta ${idConsulta} no termino de procesarse en 3 minutos.`);
  } finally {
    mc.off('response', escuchar);
  }
  log(`consulta ${idConsulta}: ${cantidad} comprobante(s)`);
  if (cantidad === 0) return { zip: null, idConsulta, cantidad };

  const url = `${MC_URL}descargarComprobantes.do?id=${idConsulta}&tc=${tipo}&tf=csv`;
  const espera = mc.waitForEvent('download', { timeout: 120000 });
  await mc.evaluate((u) => { window.location.href = u; }, url);
  const download = await espera;
  await download.saveAs(destinoZip);
  return { zip: destinoZip, idConsulta, cantidad };
}

module.exports = { abrir, descargar, rangoMes, cuitEnEncabezado, fmtCuit, MC_URL };
