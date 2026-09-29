// ─────────────────────────────────────────────────────────────────────────────
// Base vendorizada desde el plugin fisco-ar (MIT) — https://github.com/javiergradiche/fisco-ar-claude-plugin
// (scripts/lib/rcel.js). Copiada el 29/09/2026 al repo del Estudio BVA y
// ampliada para facturacion-arca.js: elegir la empresa por CUIT/nombre en vez
// de la primera, fecha del comprobante, condición de venta, alícuota de IVA,
// esperar la razón social del padrón y controlar la pantalla de confirmación.
// ─────────────────────────────────────────────────────────────────────────────

// rcel.js — wizard de "Comprobantes en Línea" (RCEL) de ARCA.
//
// El flujo es el mismo para todos los tipos de comprobante:
//   abrir → elegir empresa → Generar Comprobantes → PV + tipo → 4 pasos → CAE
//
// Ojo: el tipo de comprobante NO es libre. RCEL puebla los tipos por AJAX recién
// después de elegir el punto de venta, y cada PV habilita los suyos.

const fs = require('fs');

const MENU_URL = 'https://fe.afip.gob.ar/rcel/jsp/menu_ppal.jsp';

const log = (m) => process.stderr.write(`[rcel] ${m}\n`);

const norm = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
  .toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();

// Setea un <select> por id, eligiendo la opción cuyo texto matchea `re`.
// Devuelve { ok, picked } o { ok:false, opts } con las opciones disponibles,
// para que el error diga qué SÍ había.
async function elegirSelect(page, id, re) {
  return page.evaluate(({ id, src, flags }) => {
    const rx = new RegExp(src, flags);
    const s = document.getElementById(id);
    if (!s) return { ok: false, reason: `no existe el select #${id}` };
    const opts = Array.from(s.options).filter(o => o.value);
    // Preferir la coincidencia EXACTA sobre la parcial.
    const texto = (o) => (o.textContent || '').trim();
    const o = opts.find(o => rx.test(texto(o)) && rx.source.replace(/\\/g, '').toUpperCase() === texto(o).toUpperCase())
           || opts.find(o => rx.test(texto(o)));
    if (!o) return { ok: false, reason: 'ninguna opción matchea', opts: opts.map(texto).filter(Boolean) };
    s.value = o.value;
    s.dispatchEvent(new Event('change', { bubbles: true }));
    return { ok: true, picked: (o.textContent || '').trim() };
  }, { id, src: re.source, flags: re.flags });
}

// Opciones (con valor) de un select, para decidir antes de elegir.
async function opcionesSelect(page, id) {
  return page.evaluate((id) => {
    const s = document.getElementById(id);
    if (!s) return null;
    return Array.from(s.options).filter(o => o.value).map(o => (o.textContent || '').trim());
  }, id);
}

// Escribe en un input/textarea por id, disparando los eventos que RCEL escucha
// (algunos campos recalculan en keyup o blur). Ignora los que no existen: cada
// tipo de comprobante muestra un subconjunto distinto.
async function setCampos(page, valores) {
  return page.evaluate((vals) => {
    const puestos = [], faltantes = [];
    for (const [id, val] of Object.entries(vals)) {
      if (val === undefined || val === null || val === '') continue;
      const el = document.getElementById(id);
      if (!el) { faltantes.push(id); continue; }
      if (el.type === 'checkbox') { if (el.checked !== !!val) el.click(); }
      else {
        el.value = String(val);
        for (const ev of ['input', 'keyup', 'change', 'blur']) el.dispatchEvent(new Event(ev, { bubbles: true }));
      }
      puestos.push(id);
    }
    return { puestos, faltantes };
  }, valores);
}

// Valor actual de un input por id (o null si no existe).
async function valorCampo(page, id) {
  return page.evaluate((id) => { const e = document.getElementById(id); return e ? e.value : null; }, id);
}

// Lee en qué paso del wizard está la página ("PASO 2 DE 4" → 2). Reintenta si
// la página está navegando justo en ese momento.
async function pasoActual(page, intentos = 3) {
  for (let i = 0; i < intentos; i++) {
    try {
      return await page.evaluate(() => {
        const m = document.body.innerText.replace(/\s+/g, ' ').match(/PASO\s+(\d)\s+DE\s+4/i);
        return m ? Number(m[1]) : null;
      });
    } catch (e) {
      if (!/Execution context was destroyed|navigation/i.test(e.message) || i === intentos - 1) throw e;
      await page.waitForLoadState('domcontentloaded', { timeout: 20000 }).catch(() => {});
      await page.waitForTimeout(800);
    }
  }
  return null;
}

// RCEL valida con alert() del browser. Playwright los descarta solo, así que sin
// esto los mensajes de "falta tal campo" se pierden y el script cree que avanzó.
// Llamar una vez por página; los alerts quedan en el array devuelto.
function capturarAlertas(page) {
  const alertas = [];
  page.on('dialog', async (d) => { alertas.push(d.message().replace(/\s+/g, ' ').trim()); await d.accept().catch(() => {}); });
  return alertas;
}

// Avanza al paso siguiente y VERIFICA que haya avanzado de verdad: sin esto una
// validación fallida deja la página donde estaba y el wizard sigue "avanzando"
// sobre la pantalla vieja, hasta creer que llegó a la confirmación.
async function continuar(page, etiqueta, { alertas = [], esperado = null } = {}) {
  const antes = await pasoActual(page);
  const clickeado = await page.evaluate(() => {
    const b = document.getElementById('btn_continuar') ||
      Array.from(document.querySelectorAll('input[type="button"], input[type="submit"], button'))
        .find(x => /continuar|siguiente/i.test(x.value || x.textContent || ''));
    if (!b) return false;
    b.click();
    return true;
  });
  if (!clickeado) throw new Error(`No encontré el botón para continuar en "${etiqueta}"`);
  await page.waitForLoadState('domcontentloaded', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(900);

  let ahora = await pasoActual(page);
  // La skill manual lo decía: después de "Continuar" la página a veces tarda.
  if (esperado && ahora !== esperado) {
    await page.waitForTimeout(2500);
    ahora = await pasoActual(page);
  }
  const avanzo = esperado ? ahora === esperado : (antes === null || ahora === null || ahora > antes);
  if (!avanzo) {
    const msgs = [...alertas.splice(0), ...(await erroresDeValidacion(page))];
    throw new Error(
      `El wizard no avanzó a "${etiqueta}" (sigue en el paso ${ahora ?? '?'}). ` +
      (msgs.length ? `ARCA dijo: ${msgs.join(' | ')}` : 'ARCA no dio un mensaje; revisá el dump de ese paso.'));
  }
  log(`→ ${etiqueta} (paso ${ahora})`);
}

// Junta los mensajes de validación que RCEL muestra al no poder avanzar.
async function erroresDeValidacion(page) {
  return page.evaluate(() => Array.from(
    document.querySelectorAll('.error, .msg_error, #errores, .alert-danger, span[style*="color: red"], font[color="red"]'))
    .filter(e => e.offsetParent !== null)
    .map(e => (e.textContent || '').replace(/\s+/g, ' ').trim())
    .filter(t => t.length > 3));
}

// Abre RCEL desde el portal y elige la empresa a representar.
//
// A diferencia de la versión de fisco-ar (que toma la primera), acá se elige la
// del emisor pedido — por CUIT o por nombre — y se verifica en el encabezado.
// Facturar a nombre de otro sería un comprobante real imposible de deshacer.
async function abrir(context, portalPage, abrirServicio, { cuit, nombre } = {}) {
  const rcel = await abrirServicio(context, portalPage, /comprobantes en l[ií]nea/i);
  await rcel.waitForLoadState('domcontentloaded', { timeout: 30000 }).catch(() => {});
  await rcel.waitForTimeout(1500);
  await elegirEmpresa(rcel, { cuit, nombre });
  return rcel;
}

async function elegirEmpresa(rcel, { cuit, nombre } = {}) {
  const c = String(cuit || '').replace(/\D/g, '');
  const empresas = await rcel.evaluate(() =>
    Array.from(document.querySelectorAll('input.btn_empresa, .btn_empresa')).map((b, i) => ({
      i,
      texto: (b.value || b.textContent || '').replace(/\s+/g, ' ').trim(),
      contexto: ((b.closest('tr, li, div, form') || b).innerText || '').replace(/\s+/g, ' ').trim(),
    })));

  if (empresas.length) {
    const n = norm(nombre);
    let elegida = null;
    if (c) elegida = empresas.find(e => e.texto.replace(/\D/g, '').includes(c) || e.contexto.replace(/\D/g, '').includes(c));
    if (!elegida && n) elegida = empresas.find(e => norm(e.texto) === n)
                            || empresas.find(e => norm(e.texto).includes(n) || (norm(e.texto) && n.includes(norm(e.texto))));
    if (!elegida && empresas.length === 1) elegida = empresas[0];
    if (!elegida) {
      const e = new Error(
        `RCEL ofrece ${empresas.length} empresas y ninguna coincide con el emisor ` +
        `(${nombre || '?'} · CUIT ${c || '?'}): ${empresas.map(x => x.texto).join(' | ')}. ` +
        'Pasá --emisor-nombre con el nombre tal como figura en ARCA.');
      e.code = 'emisor_no_encontrado';
      throw e;
    }
    log(`empresa: ${elegida.texto}`);
    await Promise.all([
      rcel.waitForLoadState('domcontentloaded', { timeout: 20000 }).catch(() => {}),
      rcel.locator('input.btn_empresa, .btn_empresa').nth(elegida.i).click(),
    ]);
    await rcel.waitForTimeout(1500);
  }

  // Verificación: el menú muestra a quién se representa.
  const cuerpo = await rcel.evaluate(() => document.body.innerText.replace(/\s+/g, ' '));
  const okCuit = !!c && (cuerpo.match(/\d[\d-]{9,14}\d/g) || []).some(t => t.replace(/\D/g, '') === c);
  const okNombre = nombre && norm(cuerpo).includes(norm(nombre));
  if (!okCuit && !okNombre) {
    const e = new Error(
      `No pude confirmar en RCEL que se está representando al emisor ${nombre || ''} (CUIT ${c}). ` +
      `Encabezado: "${cuerpo.slice(0, 200)}". Corto antes de generar nada.`);
    e.code = 'emisor_incorrecto';
    throw e;
  }
  log(`emisor confirmado en RCEL (${okCuit ? 'por CUIT' : 'por nombre'})`);
}

// Menú principal → Generar Comprobantes. Se usa para cada factura nueva: la
// skill manual lo pedía así ("nunca uses < Volver").
async function irAGenerar(rcel) {
  if (!/menu_ppal/.test(rcel.url())) {
    await rcel.goto(MENU_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await rcel.waitForTimeout(800);
  }
  const ok = await rcel.evaluate(() => {
    const el = Array.from(document.querySelectorAll('a, input, button'))
      .find(x => /generar comprobantes/i.test(x.textContent || x.value || ''));
    if (el) el.click();
    return !!el;
  });
  if (!ok) throw new Error('No encontré "Generar Comprobantes" en el menú de RCEL');
  await rcel.waitForLoadState('domcontentloaded', { timeout: 30000 }).catch(() => {});
}

// Elige punto de venta y tipo de comprobante. El tipo se puebla por AJAX
// después del PV, así que hay que esperarlo.
async function elegirPvYTipo(rcel, { pv, tipoRe, tipoNombre }) {
  await rcel.waitForSelector('#puntodeventa', { state: 'attached', timeout: 30000 });

  const pvRes = await rcel.evaluate((pv) => {
    const s = document.getElementById('puntodeventa');
    const padded = String(pv).padStart(5, '0');
    const o = Array.from(s.options).find(o =>
      String(o.value) === String(pv) || (o.textContent || '').replace(/\s/g, '').includes(padded));
    if (!o) return { ok: false, opts: Array.from(s.options).map(o => (o.textContent || '').trim()).filter(Boolean) };
    s.value = o.value;
    s.dispatchEvent(new Event('change', { bubbles: true }));
    return { ok: true, picked: (o.textContent || '').trim() };
  }, pv);
  if (!pvRes.ok) {
    throw new Error(`El punto de venta ${pv} no está habilitado para Comprobantes en Línea. Disponibles: ${JSON.stringify(pvRes.opts)}`);
  }
  log(`PV: ${pvRes.picked}`);

  await rcel.waitForFunction(() => {
    const s = document.getElementById('universocomprobante');
    return s && Array.from(s.options).some(o => o.value);
  }, { timeout: 15000 }).catch(() => {});

  const tipoRes = await elegirSelect(rcel, 'universocomprobante', tipoRe);
  if (!tipoRes.ok) {
    throw new Error(
      `El punto de venta ${pv} no habilita ${tipoNombre}. ` +
      `Tipos habilitados: ${JSON.stringify(tipoRes.opts || [])}`);
  }
  log(`tipo: ${tipoRes.picked}`);
  return { pv: pvRes.picked, tipo: tipoRes.picked };
}

// Fecha del comprobante (RCEL pone hoy por defecto; hay que pisarla). El id
// conocido es "fc"; si cambia, se busca el input por la etiqueta de su fila.
async function setFechaComprobante(rcel, fecha) {
  const id = await rcel.evaluate(() => {
    if (document.getElementById('fc')) return 'fc';
    const celda = Array.from(document.querySelectorAll('td, label, span'))
      .find(x => /fecha del comprobante/i.test(x.textContent || '') && (x.textContent || '').length < 60);
    const fila = celda && (celda.closest('tr') || celda.parentElement);
    const inp = fila && fila.querySelector('input[type="text"], input:not([type])');
    if (!inp) return null;
    if (!inp.id) inp.id = '__fecha_comprobante';
    return inp.id;
  });
  if (!id) throw new Error('No encontré el campo "Fecha del Comprobante" en el paso 1');
  await setCampos(rcel, { [id]: fecha });
  const puesto = await valorCampo(rcel, id);
  if (puesto !== fecha) throw new Error(`La fecha del comprobante quedó en "${puesto}" en vez de ${fecha}`);
}

// Después de escribir el CUIT del receptor, RCEL consulta el padrón y completa
// la razón social. Se espera a que aparezca: si no llega, el CUIT no existe o
// el padrón no respondió, y seguir sería facturarle a nadie.
async function esperarRazonSocial(rcel, { timeout = 15000 } = {}) {
  const fin = Date.now() + timeout;
  while (Date.now() < fin) {
    const r = await rcel.evaluate(() => {
      const val = (id) => { const e = document.getElementById(id); if (!e) return ''; return e.tagName === 'SELECT' ? ((e.options[e.selectedIndex] || {}).text || '') : e.value; };
      return { razon: (val('razonsocialreceptor') || '').trim(), domicilio: (val('domicilioreceptor') || '').trim() };
    });
    if (r.razon && !/recuperando/i.test(r.razon)) return r;
    await rcel.waitForTimeout(700);
  }
  return { razon: '', domicilio: '' };
}

// Tilda la condición de venta (checkbox con su <label>). Destilda las otras
// que hayan quedado marcadas, así la factura lleva sólo la pedida.
async function elegirCondicionVenta(rcel, re) {
  return rcel.evaluate(({ src, flags }) => {
    const rx = new RegExp(src, flags);
    const pares = Array.from(document.querySelectorAll('input[type="checkbox"]')).map(i => {
      const l = (i.id && document.querySelector(`label[for="${i.id}"]`)) || i.closest('label');
      const texto = ((l && l.textContent) || (i.nextSibling && i.nextSibling.textContent) || '').replace(/\s+/g, ' ').trim();
      return { i, texto };
    }).filter(p => /contado|tarjeta|cuenta corriente|cheque|transferencia|otr[ao]s?\b|medios? de pago/i.test(p.texto));
    const objetivo = pares.find(p => rx.test(p.texto));
    if (!objetivo) return { ok: false, opts: pares.map(p => p.texto) };
    for (const p of pares) if (p !== objetivo && p.i.checked) p.i.click();
    if (!objetivo.i.checked) objetivo.i.click();
    return { ok: objetivo.i.checked, picked: objetivo.texto };
  }, { src: re.source, flags: re.flags });
}

// Texto visible de la pantalla de confirmación (paso 4), para controlarlo
// contra lo aprobado antes de apretar "Confirmar Datos".
async function textoPantalla(rcel) {
  return rcel.evaluate(() => document.body.innerText.replace(/[ \t]+/g, ' ').replace(/\n\s*\n+/g, '\n').trim());
}

// ¿Aparece `importe` en el texto? RCEL a veces muestra 1234.50 y a veces
// 1.234,50: se prueban las dos lecturas de cada número.
function contieneImporte(texto, importe) {
  const tokens = String(texto).match(/\d[\d.,]*\d|\d/g) || [];
  return tokens.some(t => {
    const a = Number(t.replace(/\./g, '').replace(',', '.'));
    const b = Number(t.replace(/,/g, ''));
    return Math.abs(a - importe) < 0.005 || Math.abs(b - importe) < 0.005;
  });
}

// Consulta los comprobantes ya generados de un punto de venta.
//
// Es la fuente autoritativa: la pantalla que queda después de confirmar mezcla
// divs de "Generando Comprobante…", "Comprobante Generado" y "Error!" que están
// siempre en el DOM, así que leer el CAE de ahí es poco fiable.
async function consultarGenerados(rcel, { pv, desde, hasta } = {}) {
  if (!/menu_ppal/.test(rcel.url())) {
    await rcel.goto(MENU_URL, { waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => {});
    await rcel.waitForTimeout(800);
  }
  await rcel.getByText('Consultas', { exact: true }).first().click({ timeout: 20000 });
  await rcel.waitForLoadState('domcontentloaded').catch(() => {});
  await rcel.waitForTimeout(1500);

  await rcel.evaluate(({ pv, desde, hasta }) => {
    const set = (id, v) => { const e = document.getElementById(id); if (e && v) { e.value = v; e.dispatchEvent(new Event('change', { bubbles: true })); } };
    set('fed', desde); set('feh', hasta);
    const s = document.getElementById('puntodeventa');
    if (s && pv) {
      const o = [...s.options].find(o => (o.textContent || '').replace(/\s/g, '').includes(String(pv).padStart(5, '0')));
      if (o) { s.value = o.value; s.dispatchEvent(new Event('change', { bubbles: true })); }
    }
  }, { pv, desde, hasta });
  await rcel.waitForTimeout(600);
  await rcel.evaluate(() => {
    const b = [...document.querySelectorAll('input[type=button],input[type=submit],button')]
      .find(x => /buscar|consultar/i.test(x.value || x.textContent || ''));
    if (b) b.click();
  });
  await rcel.waitForTimeout(4000);

  return rcel.evaluate(() => [...document.querySelectorAll('table tr')]
    .filter(tr => [...tr.querySelectorAll('td')].length >= 6)
    .map(tr => {
      const c = [...tr.querySelectorAll('td')].map(td => td.innerText.replace(/\s+/g, ' ').trim());
      if (!/^\d{2}\/\d{2}\/\d{4}$/.test(c[0])) return null;
      // El link "Imprimir" lleva el id interno del comprobante: imprimirComprobante.do?c=<id>
      const acc = [...tr.querySelectorAll('a, input, button')]
        .map(e => e.getAttribute('onclick') || e.getAttribute('href') || '').join(' ');
      const m = acc.match(/imprimirComprobante\.do\?c=(\d+)/);
      return { fecha: c[0], tipo: c[1], comprobante: c[2], tipoDoc: c[3], nroDoc: c[4],
               cae: c[5], total: c[6], id: m ? m[1] : null };
    })
    .filter(Boolean));
}

// Baja el PDF de un comprobante ya generado, a partir del id que trae
// consultarGenerados(). Devuelve el path, o null si no vino un PDF.
async function descargarPdf(rcel, id, destinoPath) {
  const base = new URL(rcel.url());
  const url = `${base.protocol}//${base.host}${base.pathname.replace(/[^/]+$/, '')}imprimirComprobante.do?c=${id}`;
  const resp = await rcel.context().request.get(url, { timeout: 60000 });
  if (!resp.ok()) { log(`PDF: HTTP ${resp.status()} bajando el comprobante ${id}`); return null; }
  const body = await resp.body();
  if (!body.subarray(0, 5).toString().startsWith('%PDF')) {
    log(`PDF: la respuesta del comprobante ${id} no es un PDF`);
    return null;
  }
  fs.writeFileSync(destinoPath, body);
  return destinoPath;
}

module.exports = {
  MENU_URL, abrir, elegirEmpresa, irAGenerar, elegirPvYTipo, elegirSelect, opcionesSelect,
  setCampos, valorCampo, setFechaComprobante, esperarRazonSocial, elegirCondicionVenta,
  continuar, erroresDeValidacion, pasoActual, capturarAlertas, textoPantalla, contieneImporte,
  consultarGenerados, descargarPdf, log, norm,
};
