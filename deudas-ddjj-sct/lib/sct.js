// sct.js — Sistema de Cuentas Tributarias de ARCA: leer la solapa DEUDAS.
//
// Los anticipos de Ganancias y Bienes Personales se leen de la solapa DEUDAS
// de la portada del SCT. NO de la solapa Vencimientos: ahi no esta el detalle
// de importe ni el numero de anticipo.
//
// La solapa "DDJJ pendientes de presentación" (leerPendientes) la usa
// deudas-ddjj-sct.js, junto con las deudas de Deudas que no son anticipos.
//
// REGLA DURA: esto solo LEE. Nunca modifica, elimina, presenta ni paga nada.

const SCT_HOST = 'https://ctacte.cloud.afip.gob.ar';
const log = (m) => process.stderr.write(`[sct] ${m}\n`);

const esLogin = (u) => /auth\.afip\.gob\.ar.*login/i.test(u);

// Abre el SCT y deja la sesion parada en la sociedad pedida.
//
// El servicio rebota al login cada tanto (URL auth.afip.gob.ar/...?system=
// cuenta_corriente_contrib). No es que las credenciales esten mal: hay que
// volver al portal y entrar de nuevo.
//
// CRITICO: el SCT recuerda el ultimo contribuyente elegido y puede abrir
// directo en la portada, SIN mostrar el selector. Si uno asume que eso ya es
// la sociedad pedida, termina relevando los anticipos de OTRA empresa y el
// Excel sale plausible pero equivocado. Por eso aca nunca se sigue sin
// verificar el CUIT contra el encabezado.
const URL_SCT = `${SCT_HOST}/contribuyente/externo`;

// Lee el CUIT del bloque de contribuyente de arriba a la derecha (el que dice
// ESTABLECIMIENTO). No sirve buscar el primer CUIT del body: en la pantalla de
// seleccion eso devuelve la primera opcion del desplegable.
// Que contribuyente tiene cargado el SCT.
//
// NO se lee del texto del encabezado: el bloque que contiene "ESTABLECIMIENTO"
// arrastra tambien las ~60 opciones del desplegable, asi que cualquier intento
// de sacar "el CUIT del encabezado" termina agarrando el primero de la lista.
// La fuente confiable es la opcion seleccionada del propio <select>, que es
// justamente lo que el SCT esta mostrando.
//
// Ojo: "el primer <select>" no sirve. En el iframe de la portada el primero es
// el "Mostrar 10/20/50" de la tabla. Se toma el select cuyas opciones son CUITs.
//
// Si el login no tiene representados (persona humana que entra con su propio
// CUIT) no hay select: el SCT muestra el titular como texto entre CONTRIBUYENTE
// y ESTABLECIMIENTO. Solo en ese caso se lee de ahi, y solo si hay UN CUIT.
async function contribuyenteActivo(sct) {
  for (const f of [sct.mainFrame(), ...sct.frames()]) {
    const r = await f.evaluate(() => {
      const esCuit = (t) => (t || '').replace(/\D/g, '').length === 11;
      const s = [...document.querySelectorAll('select')]
        .find(x => [...x.options].some(o => esCuit(o.text)));
      if (!s) return null;
      const op = s.options[s.selectedIndex];
      if (!op) return null;
      const cuit = (op.text || op.value || '').replace(/\D/g, '');
      if (cuit.length !== 11) return null;
      // Razon social: la linea que sigue a "CONTRIBUYENTE" en el encabezado.
      const lineas = document.body.innerText.split('\n').map(x => x.trim()).filter(Boolean);
      const i = lineas.findIndex(x => /^CONTRIBUYENTE$/i.test(x));
      const razon = i >= 0 ? (lineas.slice(i + 1).find(x => /[A-Za-z]{3}/.test(x)) || '') : '';
      return { cuit, texto: `${razon} [${cuit}]`.trim() };
    }).catch(() => null);
    if (r && r.cuit) return r;
  }
  // Sin select de contribuyentes: leer el encabezado del frame principal.
  const r = await sct.mainFrame().evaluate(() => {
    const lineas = document.body.innerText.split('\n').map(x => x.trim()).filter(Boolean);
    const i = lineas.findIndex(x => /^CONTRIBUYENTE$/i.test(x));
    const j = lineas.findIndex((x, k) => k > i && /^ESTABLECIMIENTO/i.test(x));
    if (i < 0 || j < 0) return null;
    const bloque = lineas.slice(i + 1, j);
    const cuits = [...new Set(bloque.map(x => x.replace(/\D/g, '')).filter(x => x.length === 11))];
    if (cuits.length !== 1) return null;
    const razon = bloque.find(x => /[A-Za-z]{3}/.test(x)) || '';
    return { cuit: cuits[0], texto: `${razon} [${cuits[0]}]`.trim() };
  }).catch(() => null);
  return r || { cuit: null, texto: null };
}

// Espera a que cargue la portada (las solapas Vencimientos/Deudas).
async function esperarPortada(sct, ms = 30000) {
  const hasta = Date.now() + ms;
  while (Date.now() < hasta) {
    for (const f of [sct.mainFrame(), ...sct.frames()]) {
      const ok = await f.evaluate(() =>
        /Vencimientos/i.test(document.body.innerText) &&
        /Deudas/i.test(document.body.innerText)).catch(() => false);
      if (ok) return true;
    }
    await sct.waitForTimeout(1500);
  }
  return false;
}

// Elige el contribuyente en el desplegable de la pantalla de seleccion.
async function elegirContribuyente(sct, cuit) {
  const sel = sct.locator('select').first();
  if (!(await sel.isVisible({ timeout: 8000 }).catch(() => false))) return false;

  // Ojo: el `value` de las opciones NO es el CUIT (es un indice: "5"); el
  // CUIT esta en el texto. selectOption con el CUIT suelto compara contra value
  // o label y en la practica no siempre cambiaba (paso con dos sociedades):
  // se busca la opcion por texto y se elige por su value exacto.
  const valor = await sel.evaluate((s, c) => {
    const o = [...s.options].find(x => x.text.replace(/\D/g, '') === c);
    return o ? o.value : null;
  }, String(cuit));
  if (valor == null) {
    const e = new Error(`El CUIT ${cuit} no figura entre los contribuyentes del SCT.`);
    e.code = 'cuit_no_disponible';
    throw e;
  }
  await sel.selectOption({ value: valor });
  await sct.waitForTimeout(2000);
  await esperarPortada(sct);
  await sct.waitForTimeout(1500);
  return true;
}

async function abrirSct(context, page, cuit, { intentos = 3 } = {}) {
  const { abrirServicio } = require('./arca-login');
  const objetivo = String(cuit).replace(/\D/g, '');
  let sct = null;

  for (let i = 1; i <= intentos; i++) {
    sct = await abrirServicio(
      context, page, /Sistema de Cuentas Tributarias|Cuentas Tributarias/i);
    await sct.waitForLoadState('networkidle', { timeout: 40000 }).catch(() => {});
    await sct.waitForTimeout(4000);
    if (!esLogin(sct.url())) break;

    log(`intento ${i}/${intentos}: el SCT reboto al login; vuelvo al portal`);
    await sct.close().catch(() => {});
    await page.goto('https://portalcf.cloud.afip.gob.ar/portal/app/',
                    { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(4000);
  }
  if (esLogin(sct.url())) {
    const e = new Error('El SCT rebota al login. Suele pasar con varias sesiones ' +
                        'abiertas seguidas: espera un rato y reintenta.');
    e.code = 'sct_rebota';
    throw e;
  }

  return posicionar(sct, objetivo);
}

// Deja un SCT ya abierto parado en el CUIT pedido, verificandolo. Sirve tanto
// al abrirlo como para pasar a otro representado sin volver a loguearse.
async function posicionar(sct, cuit) {
  const objetivo = String(cuit).replace(/\D/g, '');
  await elegirContribuyente(sct, objetivo);

  // Verificar y, si hace falta, forzar la pantalla de seleccion volviendo a
  // la URL base del SCT.
  let activo = await contribuyenteActivo(sct);
  if (activo.cuit !== objetivo) {
    log(`el SCT abrio en ${activo.cuit || '?'}; fuerzo la pantalla de seleccion`);
    await sct.goto(URL_SCT, { waitUntil: 'networkidle', timeout: 40000 }).catch(() => {});
    await sct.waitForTimeout(4000);
    await elegirContribuyente(sct, objetivo);
    activo = await contribuyenteActivo(sct);
  }

  if (activo.cuit !== objetivo) {
    const e = new Error(
      `No pude posicionarme en el CUIT ${objetivo}: el SCT sigue mostrando ` +
      `${activo.cuit || 'otro contribuyente'} (${activo.texto || 'sin encabezado'}). ` +
      `Corto antes de relevar, para no entregar los anticipos de otra sociedad.`);
    e.code = 'contribuyente_incorrecto';
    throw e;
  }
  log(`contribuyente confirmado: ${activo.texto}`);

  // El contenido vive en un iframe Vue (…/scripts/vue/homeContrib).
  const marco = sct.frames().find(f => /homeContrib/.test(f.url())) || sct.mainFrame();
  return { sct, marco };
}

// Numero que muestra el badge de una solapa ("Deudas 11", "DDJJ pendientes de
// presentación 0"). null si no se encuentra.
async function contadorSolapa(marco, re) {
  return marco.evaluate((src) => {
    const re = new RegExp(src, 'i');
    const t = [...document.querySelectorAll('a,button,li')]
      .map(e => (e.innerText || '').trim().replace(/\s+/g, ' '))
      .find(x => re.test(x));
    const m = t && t.match(/(\d+)\s*$/);
    return m ? Number(m[1]) : null;
  }, re.source);
}

// Subir el tamano de pagina. El control es un <select> numerico de
// BootstrapVue (en este frame no hay jQuery, asi que la API de DataTables no
// aplica). Se acepta cualquier forma: se elige "Todos" si existe, si no el
// numero mas grande.
async function mostrarTodos(sct, marco) {
  const selects = marco.locator('select');
  const cantSel = await selects.count();
  for (let i = 0; i < cantSel; i++) {
    const sl = selects.nth(i);
    if (!(await sl.isVisible().catch(() => false))) continue;
    const ops = (await sl.locator('option').allTextContents()).map(x => x.trim());
    const nums = ops.filter(o => o !== '' && !isNaN(Number(o)));
    if (!nums.length) continue;
    const todos = ops.find(o => /todos|all/i.test(o));
    const elegida = todos || nums.reduce((a, b) => (Number(b) > Number(a) ? b : a), nums[0]);
    await sl.selectOption(elegida).catch(() => {});
    log(`paginacion: "Mostrar" en ${elegida}`);
    break;
  }
  await sct.waitForTimeout(2500);
}

// Lee la tabla de la solapa ACTIVA. Si se tomara "la tabla mas grande de la
// pagina" se leeria Vencimientos cuando Deudas esta vacia.
const LEER_TABLA_ACTIVA = () => {
  const norm = (s) => (s || '').trim().replace(/\s+/g, ' ');
  const panel = document.querySelector('.tab-pane.active, [role="tabpanel"].active')
    || document.querySelector('.tab-pane:not([style*="display: none"])');
  const ambito = panel || document;
  const tabla = [...ambito.querySelectorAll('table')]
    .filter(t => t.offsetParent)
    .sort((a, b) => b.rows.length - a.rows.length)[0];
  if (!tabla) return { encabezados: [], filas: [] };
  const encabezados = [...tabla.querySelectorAll('thead th, thead td')]
    .map(x => norm(x.innerText).replace(/\s*\(Click to sort.*?\)\s*/i, '').trim());
  // Una tabla vacia trae una sola fila "No se encontraron resultados": no es un
  // registro (en DDJJ pendientes salia como 1 pendiente).
  const vacia = (r) => [...r.cells].filter(x => norm(x.innerText)).length <= 1 &&
    /no se encontraron|no hay (datos|registros)|sin resultados/i.test(r.innerText || '');
  const filas = [...tabla.querySelectorAll('tbody tr')].filter(r => !vacia(r)).map(r => {
    const celdas = [...r.cells].map(x => norm(x.innerText));
    const o = {};
    encabezados.forEach((h, i) => { if (h) o[h] = celdas[i]; });
    const cls = (r.className || '') + ' ' + (r.getAttribute('style') || '');
    o.__rojo = /danger|vencid|red/i.test(cls);
    o.__clave = celdas.join('|');
    return o;
  }).filter(f => f.__clave.replace(/[|\s]/g, '').length > 0);
  return { encabezados, filas };
};

// Acumular pagina por pagina. Es lo unico que no depende de que el selector
// de tamano haya funcionado: si quedo en 10 por pagina, igual se juntan
// todas las filas clickeando SIGUIENTE.
async function acumularPaginas(sct, marco, esperadas) {
  let encabezados = [], filas = [];
  const vistas = new Set();
  for (let pag = 1; pag <= 30; pag++) {
    const parcial = await marco.evaluate(LEER_TABLA_ACTIVA);
    if (parcial.encabezados.length) encabezados = parcial.encabezados;
    let nuevas = 0;
    for (const f of parcial.filas) {
      if (vistas.has(f.__clave)) continue;
      vistas.add(f.__clave); filas.push(f); nuevas++;
    }
    if (esperadas && filas.length >= esperadas) break;
    const sig = marco.locator('a, button, li').filter({ hasText: /^\s*SIGUIENTE\s*$/i }).first();
    const hay = await sig.isVisible().catch(() => false);
    const deshabilitado = hay && await sig.evaluate(
      e => e.disabled || /disabled/i.test(e.className) ||
           /disabled/i.test((e.parentElement && e.parentElement.className) || '')).catch(() => true);
    if (!hay || deshabilitado || nuevas === 0) break;
    await sig.click().catch(() => {});
    await sct.waitForTimeout(2000);
  }
  return { encabezados, filas };
}

// Abre la solapa DEUDAS y devuelve todas sus filas, ya mapeadas por encabezado.
async function leerDeudas(sct, marco, { confirmando = false } = {}) {
  const tab = marco.locator('a, button, li', { hasText: /^\s*\$?\s*Deudas/i }).first();
  await tab.waitFor({ state: 'visible', timeout: 20000 });
  await tab.click();
  await sct.waitForTimeout(4000);

  // Cuantas deudas dice la solapa. Sirve de control: si despues leo menos
  // filas que esto, me quedo paginando y estaria perdiendo registros.
  let esperadas = await contadorSolapa(marco, /^\$?\s*Deudas\s+\d+$/);

  await mostrarTodos(sct, marco);
  const { encabezados, filas } = await acumularPaginas(sct, marco, esperadas);

  const resumen = await marco.evaluate(() => {
    const txt = document.body.innerText;
    const cap = (re) => { const m = txt.match(re); return m ? m[1].trim() : null; };
    return {
      saldoTotal:       cap(/Saldo total a pagar:\s*\$?\s*([\d.,]+)/i),
      capitalVencido:   cap(/Capital vencido\s*\$?\s*([\d.,]+)/i),
      capitalNoVencido: cap(/Capital no vencido\s*\$?\s*([\d.,]+)/i),
      intResarcitorios: cap(/Int\.?\s*Resarcitorios\s*\$?\s*([\d.,]+)/i),
      intPunitorios:    cap(/Int\.?\s*Punitorios\s*\$?\s*([\d.,]+)/i),
    };
  });
  const datos = { encabezados, filas, resumen };

  // Guardarrail: la tabla de Deudas SIEMPRE tiene columna Saldo; la de
  // Vencimientos no. Si lo que quedo visible no tiene Saldo, estamos leyendo
  // Vencimientos — que no trae importe ni sirve para este relevamiento. Pasa
  // cuando Deudas esta en 0: la solapa no se abre y queda Vencimientos.
  //
  // Pero un "Deudas 0" recien cambiado el contribuyente NO alcanza para decir
  // que no hay deudas: el contador carga despues que la tabla (una sociedad leyo 0
  // teniendo 12 deudas). Antes de concluir "sin deudas" se espera y se vuelve a
  // probar una vez; solo si el contador sigue en 0 se da por vacia.
  const tieneSaldo = datos.encabezados.some(h => /saldo|importe/i.test(h));
  if (!tieneSaldo) {
    if (!confirmando) {
      log(`no aparecio la tabla de Deudas (contador: ${esperadas}); espero y reintento`);
      await sct.waitForTimeout(6000);
      return leerDeudas(sct, marco, { confirmando: true });
    }
    if (esperadas === 0) {
      log('la solapa Deudas sigue en 0 y no se abre: no hay deudas que relevar');
      return { encabezados: [], filas: [], resumen: datos.resumen, esperadas: 0 };
    }
    const e = new Error(
      `La tabla que quedo visible no tiene columna Saldo (encabezados: ` +
      `${datos.encabezados.filter(Boolean).join(', ')}). Eso es Vencimientos, no Deudas, ` +
      `y no trae importes. No sigo para no reportar datos incompletos.`);
    e.code = 'solapa_incorrecta';
    throw e;
  }

  // Con la tabla de Deudas a la vista, un contador en 0 es el que todavia no
  // cargo: se ignora (la paginacion igual se recorrio hasta SIGUIENTE).
  if (esperadas === 0 && datos.filas.length) {
    log(`el contador de Deudas dice 0 pero la tabla tiene ${datos.filas.length} filas: lo ignoro`);
    esperadas = null;
  }

  log(`solapa Deudas: ${datos.filas.length} filas` +
      (esperadas != null ? ` (la solapa declara ${esperadas})` : ''));
  if (esperadas) {
    if (datos.filas.length < esperadas) {
    const e = new Error(
      `Lei ${datos.filas.length} filas pero la solapa Deudas declara ${esperadas}. ` +
      `Falta paginacion: no sigo, porque podria estar perdiendo anticipos.`);
      e.code = 'lectura_incompleta';
      throw e;
    }
  }
  return { ...datos, esperadas };
}

// Abre la solapa "DDJJ pendientes de presentación" y devuelve sus filas.
//
// Columnas (28/09/2026): Establecimiento | Impuesto | Concepto | Subconcepto |
// Período | Ant/Cuota | Vencimiento. No trae importes: son obligaciones de
// presentacion (DDJJ y regimenes de informacion) que el SCT da por omitidas.
//
// Mismo cuidado que con Deudas: si la solapa esta en 0 no se abre y queda a la
// vista Vencimientos, que tambien es una tabla sin Saldo. Se distinguen porque
// Vencimientos tiene "Detalle" (PAGO / PRESENTACION) y "Fecha Vencimiento", y
// la de pendientes no. Un contador en 0 recien cambiado el contribuyente puede
// ser uno que todavia no cargo: se espera y se reintenta una vez.
async function leerPendientes(sct, marco, { confirmando = false } = {}) {
  const RE_TAB = /DDJJ pendientes/i;
  const tab = marco.locator('a, button, li', { hasText: RE_TAB }).first();
  await tab.waitFor({ state: 'visible', timeout: 20000 });
  await tab.click();
  await sct.waitForTimeout(4000);

  let esperadas = await contadorSolapa(marco, /^DDJJ pendientes.*?\d+$/);
  await mostrarTodos(sct, marco);
  const { encabezados, filas } = await acumularPaginas(sct, marco, esperadas);

  const esVencimientos = encabezados.some(h => /^detalle$/i.test(h) || /fecha vencimiento/i.test(h));
  const esDeudas = encabezados.some(h => /saldo|importe/i.test(h));
  const esPendientes = !esVencimientos && !esDeudas &&
    encabezados.some(h => /impuesto/i.test(h)) && encabezados.some(h => /^vencimiento$/i.test(h));

  if (!esPendientes) {
    if (!confirmando) {
      log(`no aparecio la tabla de DDJJ pendientes (contador: ${esperadas}); espero y reintento`);
      await sct.waitForTimeout(6000);
      return leerPendientes(sct, marco, { confirmando: true });
    }
    if (esperadas === 0) {
      log('la solapa DDJJ pendientes sigue en 0: no hay presentaciones pendientes');
      return { encabezados: [], filas: [], esperadas: 0 };
    }
    const e = new Error(
      `La tabla visible no es la de DDJJ pendientes (encabezados: ` +
      `${encabezados.filter(Boolean).join(', ') || 'ninguno'}; contador ${esperadas}). ` +
      `No sigo para no reportar que no hay pendientes cuando no se pudo leer.`);
    e.code = 'solapa_incorrecta';
    throw e;
  }

  if (esperadas === 0 && filas.length) {
    log(`el contador de DDJJ pendientes dice 0 pero la tabla tiene ${filas.length} filas: lo ignoro`);
    esperadas = null;
  }
  log(`solapa DDJJ pendientes: ${filas.length} filas` +
      (esperadas != null ? ` (la solapa declara ${esperadas})` : ''));
  if (esperadas && filas.length < esperadas) {
    const e = new Error(
      `Lei ${filas.length} filas pero la solapa DDJJ pendientes declara ${esperadas}. ` +
      `Falta paginacion: no sigo, porque podria estar perdiendo presentaciones.`);
    e.code = 'lectura_incompleta';
    throw e;
  }
  return { encabezados, filas, esperadas };
}

const col = (fila, ...nombres) => {
  for (const n of nombres) {
    const k = Object.keys(fila).find(x => x.toLowerCase().includes(n.toLowerCase()));
    if (k && fila[k]) return fila[k];
  }
  return '';
};

// "12/07/2012" -> Date. El SCT usa DD/MM/AAAA.
function aFecha(s) {
  const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}

// Vencido o no, por fecha. Es mas confiable que el color de la fila, que
// depende de como el SCT pinte el <tr> ese dia.
//
// Se compara contra el INICIO de hoy: un anticipo que vence hoy todavia no
// esta vencido (comparando contra `new Date()` salia Vencido desde las 00:01).
function estadoDe(fila, hoy = new Date()) {
  const f = aFecha(col(fila, 'Vencimiento'));
  if (!f) return 'Sin fecha';
  const inicioHoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  return f < inicioHoy ? 'Vencido' : 'No Vencido';
}

// Separa los anticipos del resto y, dentro de esos, los de Ganancias y Bienes
// Personales de cualquier otro impuesto.
//
// Normalmente los unicos impuestos con anticipos son esos dos. Si aparece otro
// NO se descarta: es una novedad real y se reporta aparte para que el usuario
// la confirme.
const esAnticipo = (f) =>
  /anticipo/i.test(col(f, 'Concepto') + ' ' + col(f, 'Subconcepto')) ||
  (Number(col(f, 'Ant/Cuota', 'Ant.')) > 0 &&
   !/declaraci.n jurada/i.test(col(f, 'Concepto')));

function clasificar(filas, hoy = new Date()) {

  // Ojo: "217 - SICORE-IMPTO.A LAS GANANCIAS" y "787 - RET ART 79 LEY GCIAS"
  // tambien dicen Ganancias pero son retenciones, no el impuesto propio. Un
  // anticipo de esos seria una novedad y tiene que salir en inesperados.
  const esperado = (f) => {
    const imp = col(f, 'Impuesto');
    return /ganancias|bienes\s*personales/i.test(imp) &&
           !/sicore|ret(\.|enc|\s)|percep|art\.?\s*79|beneficiarios del exterior/i.test(imp);
  };

  const mapear = (f) => ({
    impuesto: col(f, 'Impuesto'),
    concepto: col(f, 'Concepto'),
    subconcepto: col(f, 'Subconcepto'),
    periodo: col(f, 'Período', 'Periodo'),
    nroAnticipo: col(f, 'Ant/Cuota', 'Ant.'),
    importe: col(f, 'Saldo', 'Importe'),
    // Un anticipo pagado fuera de termino queda con Saldo $ 0 y "Vencido",
    // pero sigue en Deudas por los intereses: sin estas columnas parece un
    // error del SCT.
    intResarcitorio: col(f, 'Resarc'),
    intPunitorio: col(f, 'Punit'),
    vencimiento: col(f, 'Vencimiento'),
    estado: estadoDe(f, hoy),
  });

  const anticipos = filas.filter(esAnticipo);
  return {
    anticipos:  anticipos.filter(esperado).map(mapear),
    inesperados: anticipos.filter(f => !esperado(f)).map(mapear),
    otrasDeudas: filas.filter(f => !esAnticipo(f)).length,
  };
}

// Deudas de la solapa Deudas que NO son anticipos (saldos de DDJJ, multas,
// intereses, retenciones), con el mismo criterio de "anticipo" que clasificar():
// lo que un script cuenta como anticipo, el otro no lo cuenta como deuda.
function otrasDeudas(filas, hoy = new Date()) {
  return filas.filter(f => !esAnticipo(f)).map(f => ({
    establecimiento: col(f, 'Establecimiento'),
    impuesto: col(f, 'Impuesto'),
    concepto: col(f, 'Concepto'),
    subconcepto: col(f, 'Subconcepto'),
    periodo: col(f, 'Período', 'Periodo'),
    cuota: col(f, 'Ant/Cuota', 'Ant.'),
    capital: col(f, 'Saldo', 'Importe'),
    intResarcitorio: col(f, 'Resarc'),
    intPunitorio: col(f, 'Punit'),
    vencimiento: col(f, 'Vencimiento'),
    estado: estadoDe(f, hoy),
  }));
}

// Filas de la solapa DDJJ pendientes, con los dias de atraso a hoy.
function pendientes(filas, hoy = new Date()) {
  const inicioHoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  return filas.map(f => {
    const v = aFecha(col(f, 'Vencimiento'));
    return {
      establecimiento: col(f, 'Establecimiento'),
      impuesto: col(f, 'Impuesto'),
      concepto: col(f, 'Concepto'),
      subconcepto: col(f, 'Subconcepto'),
      periodo: col(f, 'Período', 'Periodo'),
      cuota: col(f, 'Ant/Cuota', 'Ant.'),
      vencimiento: col(f, 'Vencimiento'),
      diasAtraso: v ? Math.max(0, Math.round((inicioHoy - v) / 86400000)) : null,
    };
  });
}

module.exports = { SCT_HOST, abrirSct, posicionar, contribuyenteActivo, esperarPortada, leerDeudas, leerPendientes,
                   clasificar, otrasDeudas, pendientes, estadoDe, col };
