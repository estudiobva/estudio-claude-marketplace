// portal-iva.js — flujo del Libro IVA de ARCA: importar desde ARCA y descargar.
//
// Traduce a Playwright el circuito que hasta ahora se hacia a mano con Claude
// in Chrome. Es el mismo para Compras y Ventas salvo la tarjeta del libro.
//
// REGLA DURA: esto solo deja el borrador cargado y se lleva el listado.
// NUNCA toca Presentar / Confirmar / Generar DJ / "ELIMINAR TODOS" / Rectificar.
// Ver `SELECTORES_PROHIBIDOS` abajo.

const IVA_URL = 'https://siapweb.cloud.afip.gob.ar/iva/';

// Textos que este modulo no debe clickear jamas. Se chequea en clickSeguro().
const SELECTORES_PROHIBIDOS = /presentar|confirmar|generar\s+dj|eliminar\s+todos|descartar|rectific/i;

const log = (m) => process.stderr.write(`[portal-iva] ${m}\n`);

// Click con red de seguridad: si el texto del elemento cae en la lista negra,
// aborta en vez de clickear. Barato, y evita que un cambio de layout de ARCA
// termine presentando una DDJJ.
async function clickSeguro(locator, descripcion) {
  const txt = (await locator.innerText().catch(() => '')) || '';
  if (SELECTORES_PROHIBIDOS.test(txt)) {
    throw new Error(
      `ABORTADO: iba a clickear "${txt.trim().slice(0, 60)}" (${descripcion}), ` +
      `que esta en la lista de acciones prohibidas. Reviso los selectores antes de seguir.`
    );
  }
  await locator.click();
}

// Entra al Portal IVA clickeando el tile del portal de clave fiscal.
//
// HAY que pasar por el tile: navegar directo a IVA_URL devuelve
// "401 - No existe token/sign en la peticion". ARCA emite el token/sign recien
// cuando se entra al servicio desde el portal, y sin eso Jetty rechaza todo.
//
// (La documentacion de la skill vieja decia lo contrario — "ir directo, no
// clickear el tile" — pero eso valia para Claude in Chrome, donde el tile abria
// una pestaña fuera del grupo manejable. Con Playwright el popup se captura
// sin problema y es el unico camino que autentica.)
//
// Devuelve la PAGINA DEL SERVICIO, que no es la misma que la del portal.
async function abrirPortalIva(context, page) {
  const { abrirServicio } = require('./arca-login');
  const ivaPage = await abrirServicio(context, page, /Portal IVA/i);
  await ivaPage.waitForLoadState('domcontentloaded', { timeout: 60000 });
  await ivaPage.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await ivaPage.waitForTimeout(2000);

  const txt = await ivaPage.innerText('body').catch(() => '');
  if (/No existe token|401/i.test(txt)) {
    throw new Error('El Portal IVA devolvio 401: ARCA no emitio el token al abrir el servicio.');
  }
  return ivaPage;
}

// Lee el encabezado "REPRESENTANDO A: [sociedad] [CUIT]".
async function representadoActual(page) {
  const linea = await page.evaluate(() => {
    const l = document.body.innerText.split('\n').map(s => s.trim())
      .find(t => /REPRESENTANDO/i.test(t));
    return l || null;
  });
  return { linea, cuit: linea ? (linea.replace(/\D/g, '').match(/\d{11}/) || [null])[0] : null };
}

// Cambia la representada si hace falta. La eleccion se conserva entre pantallas.
async function asegurarRepresentacion(page, cuitObjetivo) {
  const c = String(cuitObjetivo).replace(/\D/g, '');
  let actual = await representadoActual(page);
  if (actual.cuit === c) {
    log(`representando a ${actual.linea}`);
    return actual;
  }

  log(`cambiando representada -> ${c} (estaba: ${actual.cuit || 'ninguna'})`);
  await page.goto(`${IVA_URL}#/changeRelation`, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('text=/REPRESENTAR A/i', { timeout: 20000 });

  // Los cards son <a class="panel">. Se matchea por CUIT, no por nombre:
  // varias sociedades comparten parte del nombre y el CUIT desempata.
  const ok = await page.evaluate((cuit) => {
    const panels = Array.from(document.querySelectorAll('a.panel, a[href*="#/"], .panel'));
    const hit = panels.find(p => p.textContent.replace(/\D/g, '').includes(cuit));
    if (!hit) return false;
    hit.scrollIntoView({ block: 'center' });
    hit.click();
    return true;
  }, c);

  if (!ok) {
    throw new Error(
      `El CUIT ${c} no aparece entre las representadas. ` +
      `Falta la relacion en Administrador de Relaciones — eso se resuelve a mano.`
    );
  }

  await page.waitForTimeout(2500);
  actual = await representadoActual(page);
  if (actual.cuit !== c) {
    throw new Error(`Intente cambiar a ${c} pero el encabezado sigue diciendo: ${actual.linea}`);
  }
  log(`ahora representando a ${actual.linea}`);
  return actual;
}

// Clickea el boton `textoBoton` que pertenece a la tarjeta cuyo titulo matchea.
//
// Filtrar `div` por hasText no sirve: matchea decenas de contenedores anidados y
// `.last()` termina eligiendo el mas interno, que no contiene al boton. Aca se
// hace al reves — se parte de los botones y se sube por los ancestros hasta
// encontrar uno que contenga el titulo. Eso identifica la tarjeta sin depender
// de como ARCA anide los divs.
//
// Si hay un solo boton con ese texto en toda la pantalla, se usa directamente.
async function clickBotonDeTarjeta(page, tituloRegex, textoBoton, descripcion) {
  const resultado = await page.evaluate(({ titulo, texto }) => {
    const reT = new RegExp(titulo, 'i');
    const reB = new RegExp(texto, 'i');
    const prohibido = /presentar|confirmar|generar\s+dj|eliminar\s+todos|descartar|rectific/i;

    const botones = Array.from(document.querySelectorAll('a, button, input[type=submit]'))
      .filter(e => e.offsetParent !== null)
      .filter(e => reB.test((e.innerText || e.value || '').trim()));

    if (!botones.length) return { ok: false, motivo: 'no hay ningun boton con ese texto' };

    const elegir = (el) => {
      const t = (el.innerText || el.value || '').trim();
      if (prohibido.test(t)) return { ok: false, motivo: `boton prohibido: "${t}"` };
      el.scrollIntoView({ block: 'center' });
      el.click();
      return { ok: true, texto: t };
    };

    // Un solo candidato: no hay ambiguedad posible.
    if (botones.length === 1) return elegir(botones[0]);

    // Varios: subir por los ancestros hasta dar con el titulo de la tarjeta.
    for (const b of botones) {
      let cur = b;
      for (let i = 0; i < 8 && cur && cur.tagName !== 'BODY'; i++) {
        if (reT.test(cur.textContent || '')) return elegir(b);
        cur = cur.parentElement;
      }
    }
    return { ok: false, motivo: `hay ${botones.length} botones "${texto}" pero ninguno bajo una tarjeta que diga "${titulo}"` };
  }, { titulo: tituloRegex.source, texto: textoBoton.source });

  if (!resultado.ok) {
    throw new Error(`No pude clickear ${descripcion}: ${resultado.motivo}.`);
  }
  log(`click en "${resultado.texto}" (${descripcion})`);
}

// ARCA admite un solo borrador por vez. Si hay otro periodo abierto (el aviso sale
// despues del INGRESAR de Registracion y declaracion), la unica salida que ofrece
// es "DESCARTAR BORRADOR", que borra el trabajo de ese mes: eso lo decide una
// persona, nunca el script.
async function verificarOtroBorrador(page, periodoAAAAMM) {
  const cuerpo = await page.innerText('body').catch(() => '');
  const otro = cuerpo.match(/borrador sin presentar, correspondiente a otro per.odo \((\d{2}\/\d{4})\)/i);
  if (otro) {
    const e = new Error(
      `ARCA tiene abierto un borrador sin presentar del periodo ${otro[1]}. ` +
      `Hay que presentarlo (o descartarlo a mano) antes de abrir ${periodoAAAAMM}. No toco nada.`
    );
    e.code = 'otro_borrador_abierto';
    throw e;
  }

}

// Pasos 4-6: nueva DDJJ, periodo, y entrar a "Registracion y declaracion".
async function abrirBorrador(page, periodoAAAAMM) {
  log(`abriendo borrador del periodo ${periodoAAAAMM}`);

  // La tarjeta "Nueva declaracion jurada" recien aparece DESPUES de elegir la
  // representada: sin eso ARCA muestra "Tu CUIT no posee activa la
  // caracterizacion requerida" y el unico boton es el CONSULTAR de presentadas.
  await page.waitForSelector('button:has-text("INGRESAR"), a:has-text("INGRESAR")', { timeout: 30000 });
  await clickBotonDeTarjeta(page, /Nueva declaraci/i, /INGRESAR/i, 'INGRESAR de Nueva declaracion jurada');

  // El <select> de periodo usa AAAAMM como value.
  const select = page.locator('select').first();
  await select.waitFor({ state: 'visible', timeout: 30000 });
  await select.selectOption(periodoAAAAMM);

  const continuar = page.locator('a, button').filter({ hasText: /CONTINUAR/i }).first();
  await clickSeguro(continuar, 'CONTINUAR del periodo');
  await page.waitForTimeout(2000);

  await verificarOtroBorrador(page, periodoAAAAMM);

  // De las dos tarjetas, la nuestra es "Registracion y declaracion".
  // La otra ("Retenciones, percepciones y pagos a cuenta") no se toca.
  await page.waitForSelector('button:has-text("INGRESAR"), a:has-text("INGRESAR")', { timeout: 30000 });
  await clickBotonDeTarjeta(page, /Registraci.n y declaraci/i, /INGRESAR/i,
                            'INGRESAR de Registracion y declaracion');
  await page.waitForTimeout(3000);
  await verificarOtroBorrador(page, periodoAAAAMM);
}

// Paso 7: datos iniciales. Puede no aparecer si el borrador ya estaba iniciado,
// y en ese caso NO hay que reconfigurar nada.
async function datosIniciales(page) {
  if (!/verDatosInicialesPresentacion/i.test(page.url())) {
    log('datos iniciales: ARCA la salteo (borrador ya iniciado)');
    return false;
  }
  log('configurando datos iniciales');

  // El toggle ya viene en CON MOVIMIENTOS: no se toca.

  // Marcar "Operaciones No Gravadas o Exentas". El select de prorrateo recien
  // aparece despues de marcarlo.
  const chk = page.locator('input[type="checkbox"]').filter({ hasNot: page.locator('[disabled]') });
  const noGravadas = page.locator('label', { hasText: /No Gravadas o Exentas/i }).first();
  if (await noGravadas.isVisible({ timeout: 5000 }).catch(() => false)) {
    await noGravadas.click();
  } else {
    await chk.first().check().catch(() => {});
  }
  await page.waitForTimeout(1500);

  // "Forma de Computo del Credito Fiscal" -> CON PRORRATEO GLOBAL.
  // Es un dropdown de Bootstrap, no un <select> nativo: hay que abrir el
  // .dropdown-toggle y despues clickear la opcion de la lista. El default de
  // ARCA es "por asignacion directa" y hay que cambiarlo si o si.
  const toggle = page.locator('.dropdown-toggle').first();
  if (await toggle.isVisible({ timeout: 5000 }).catch(() => false)) {
    await toggle.click();
    const opcion = page.locator('a, li, span').filter({ hasText: /PRORRATEO GLOBAL/i }).first();
    await opcion.click({ timeout: 10000 });
  } else {
    const nativo = page.locator('select').filter({ hasText: /PRORRATEO|ASIGNACION/i }).first();
    if (await nativo.isVisible({ timeout: 3000 }).catch(() => false)) {
      await nativo.selectOption({ label: /PRORRATEO GLOBAL/i }).catch(() => {});
    } else {
      throw new Error('No encontre el selector de Forma de Computo del Credito Fiscal.');
    }
  }

  const continuar = page.locator('a, button').filter({ hasText: /CONTINUAR/i }).first();
  await clickSeguro(continuar, 'CONTINUAR de datos iniciales');
  await page.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
  await page.waitForTimeout(2000);
  return true;
}

// Paso 8: entrar al Libro Compras o Libro Ventas desde menuPresentacion.do.
async function abrirLibro(page, libro) {
  const nombre = libro.toLowerCase() === 'ventas' ? 'Ventas' : 'Compras';
  const re = new RegExp(`Libro\\s+${nombre}`, 'i');
  log(`abriendo Libro ${nombre}`);

  if (!/verCompras|verVentas/i.test(page.url())) {
    const tarjeta = page.locator('a.panel, a, .panel').filter({ hasText: re }).first();
    await tarjeta.waitFor({ state: 'visible', timeout: 30000 });
    await clickSeguro(tarjeta, `tarjeta Libro ${nombre}`);
    await page.waitForTimeout(3000);
  }
  return nombre;
}

// Tareas de importacion del libro abierto (el "Historial de Importaciones").
// Mismo pedido que hace compras.js de ARCA: ajax.do?f=listaTareas&c=<gTipoOperacion>.
async function listaTareas(page) {
  return page.evaluate(() => new Promise((res) => {
    jQuery.ajax({ url: 'ajax.do', data: { f: 'listaTareas', c: window.gTipoOperacion }, dataType: 'json',
      timeout: 30000, success: (r) => res((r && r.datos) || []), error: () => res([]) });
  }));
}

// Paso 9: IMPORTAR -> "Importar desde ARCA...".
//
// El modal trae el filtro de fechas precargado con el primer y ultimo dia del
// periodo: NO se modifica. Y el desplegable de duplicados se deja en
// "CONSERVAR EL COMPROBANTE PREVIAMENTE INCLUIDO" (su default).
async function importarDesdeArca(page) {
  log('importando comprobantes desde ARCA');

  const btnImportar = page.locator('a, button').filter({ hasText: /^\s*IMPORTAR\s*$/i }).first();
  await btnImportar.waitFor({ state: 'visible', timeout: 30000 });
  await clickSeguro(btnImportar, 'boton IMPORTAR');

  const opcion = page.locator('a, li, button').filter({ hasText: /Importar desde ARCA/i }).first();
  await opcion.waitFor({ state: 'visible', timeout: 15000 });
  await clickSeguro(opcion, 'Importar desde ARCA...');
  await page.waitForTimeout(2000);

  // Tareas que ya existian: la nuestra es la que aparezca despues del click.
  const previas = new Set((await listaTareas(page)).map((t) => t.codigo));

  // Confirmar el modal. El boton IMPORTAR de adentro es el ultimo visible.
  const confirmar = page.locator('.modal, [role="dialog"]').locator('button, a')
    .filter({ hasText: /IMPORTAR/i }).last();
  await confirmar.waitFor({ state: 'visible', timeout: 15000 });
  await clickSeguro(confirmar, 'IMPORTAR del modal');

  // La importacion es ASINCRONA: ARCA crea una tarea (PE pendiente -> PR
  // procesando -> TE procesada) y la procesa en su servidor. Si se baja el CSV
  // antes de que termine, sale cortado (paso en 08/2026: 3.500
  // de 5.600 comprobantes, hasta el dia 20). Hay que esperar a la tarea.
  const limite = Date.now() + 15 * 60 * 1000;
  let tarea = null, ultimoLog = 0;
  while (Date.now() < limite) {
    await page.waitForTimeout(5000);
    const nuevas = (await listaTareas(page)).filter((t) => !previas.has(t.codigo) && t.tipoTarea == 1);
    tarea = nuevas[0] || null;
    if (tarea && ['TE', 'ER', 'CA'].includes(tarea.estado)) break;
    if (Date.now() - ultimoLog > 30000) {
      ultimoLog = Date.now();
      log(`importacion en curso: ${tarea ? `${tarea.estado} ${tarea.progreso ?? ''}%` : 'esperando que ARCA cree la tarea'}`);
    }
  }
  if (!tarea) throw new Error('ARCA nunca creo la tarea de importacion.');
  if (tarea.estado !== 'TE') {
    const e = new Error(tarea.estado === 'ER' || tarea.estado === 'CA'
      ? `La importacion desde ARCA termino en estado ${tarea.estado}: ${tarea.errores || 'sin detalle'}.`
      : 'La importacion desde ARCA no termino en 15 minutos: no bajo un listado incompleto.');
    e.code = 'importacion_incompleta';
    throw e;
  }
  const resumen = {
    registros: tarea.cantidadRegistros, agregados: tarea.cantidadAgregados,
    modificados: tarea.cantidadModificados, ignorados: tarea.cantidadIgnorados,
    erroneos: tarea.cantidadErroneos,
  };
  log(`importacion terminada: registros=${resumen.registros} agregados=${resumen.agregados} ` +
      `ignorados=${resumen.ignorados} erroneos=${resumen.erroneos}`);

  // La grilla (y por lo tanto el CSV, que se arma en el navegador) se cargo
  // antes de importar: hay que recargar la pagina para que traiga todo.
  await page.reload({ waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
  await page.waitForTimeout(3000);

  await cerrarModales(page);
  return resumen;
}

// Cierra cualquier modal de Bootstrap que haya quedado abierto.
//
// Importa mas de lo que parece: el modal de Historial de Importaciones
// (#modalTareas) queda con un .modal-backdrop que INTERCEPTA los clicks, y el
// boton CSV se vuelve inclickeable aunque Playwright lo vea "visible, enabled
// and stable". El sintoma es un timeout con "intercepts pointer events".
async function cerrarModales(page, intentos = 6) {
  for (let i = 0; i < intentos; i++) {
    const abierto = page.locator('.modal.in, .modal.show').first();
    if (!(await abierto.isVisible().catch(() => false))) break;

    // Primero el boton propio del modal; despues la X; por ultimo Escape.
    const cerrar = abierto.locator('button, a')
      .filter({ hasText: /^\s*(CERRAR|CLOSE|ACEPTAR|VOLVER)\s*$/i }).last();
    if (await cerrar.isVisible().catch(() => false)) {
      await cerrar.click({ timeout: 5000 }).catch(() => {});
    } else {
      const x = abierto.locator('[data-dismiss="modal"], .close').last();
      if (await x.isVisible().catch(() => false)) await x.click({ timeout: 5000 }).catch(() => {});
      else await page.keyboard.press('Escape').catch(() => {});
    }
    await page.waitForTimeout(1200);
  }

  // El backdrop puede sobrevivir al modal y seguir tapando todo.
  await page.waitForFunction(
    () => !document.querySelector('.modal-backdrop') &&
          !document.querySelector('.modal.in, .modal.show'),
    { timeout: 15000 }
  ).catch(async () => {
    log('aviso: quedo un backdrop abierto; lo saco por DOM');
    await page.evaluate(() => {
      document.querySelectorAll('.modal-backdrop').forEach(e => e.remove());
      document.querySelectorAll('.modal.in, .modal.show').forEach(e => {
        e.classList.remove('in', 'show'); e.style.display = 'none';
      });
      document.body.classList.remove('modal-open');
      document.body.style.overflow = '';
    });
  });
  await page.waitForTimeout(500);
}

// Paso 10: boton CSV de la fila "CSV | Excel | PDF".
//
// El de "Excel" baja un .xls viejo y sucio: no se usa. ARCA suele entregar un
// ZIP con el CSV adentro, pero a veces el CSV pelado — el que llama resuelve.
async function descargarCsv(page, destinoDir) {
  log('descargando listado');
  await cerrarModales(page);   // el backdrop de #modalTareas bloquea el click

  // El boton real es <button title="Exportar como CSV">. Matchear por title es
  // mas estable que por texto (el texto puede venir de un <span> con iconos).
  const btn = page.locator('button[title*="CSV" i], a[title*="CSV" i]').first()
    .or(page.locator('a, button').filter({ hasText: /^\s*CSV\s*$/i }).first());
  await btn.scrollIntoViewIfNeeded().catch(() => {});
  await btn.waitFor({ state: 'visible', timeout: 30000 });

  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 120000 }),
    clickSeguro(btn, 'boton CSV'),
  ]);

  const path = require('path');
  const sugerido = download.suggestedFilename();
  const destino = path.join(destinoDir, sugerido);
  await download.saveAs(destino);
  log(`bajado: ${sugerido}`);
  return destino;
}

// ── DDJJ ya presentadas (solo lectura) ─────────────────────────────────────
//
// Camino: inicio -> "Declaraciones juradas presentadas" CONSULTAR -> tarjeta
// "Libro IVA" CONSULTAR (#/livas.anteriores) -> fila del periodo, boton "Ver"
// (title="Ver") -> vista previa en liva (verVistaPrevia.do) -> menuPresentacion.do
// -> Libro Compras / Ventas, que se abren SIN el boton IMPORTAR.
// Cada fila tiene tambien "Rectificar": esta en SELECTORES_PROHIBIDOS y aca se
// clickea solo el boton con title="Ver".

const aNumero = (s) => Number(String(s || '0').replace(/\./g, '').replace(',', '.'));

// Filas de la lista de presentadas para un periodo AAAAMM, en el orden de ARCA
// (la primera es la ultima secuencia: la rectificativa mas reciente).
async function listarPresentadas(page, periodoAAAAMM) {
  const mmaaaa = `${periodoAAAAMM.slice(4, 6)}/${periodoAAAAMM.slice(0, 4)}`;
  return page.evaluate((per) => Array.from(document.querySelectorAll('tr'))
    .map((tr) => Array.from(tr.querySelectorAll('td')).map((td) => td.innerText.trim()))
    .filter((c) => c.length >= 5 && c[2] === per)
    .map((c) => ({ formulario: c[1], periodo: c[2], secuencia: c[3], presentada: c[4] })), mmaaaa);
}

// Abre la vista previa de la ultima presentacion del periodo. Devuelve la
// pagina de liva (puede ser una pestaña nueva), la fila elegida, todas las
// secuencias del periodo y los totales de la vista previa para controlar el CSV.
async function abrirPresentada(context, page, periodoAAAAMM) {
  log(`buscando la DDJJ presentada de ${periodoAAAAMM}`);
  await page.waitForSelector('button:has-text("CONSULTAR"), a:has-text("CONSULTAR")', { timeout: 30000 });
  await clickBotonDeTarjeta(page, /presentad/i, /CONSULTAR/i, 'CONSULTAR de Declaraciones juradas presentadas');
  await page.waitForTimeout(2000);
  await clickBotonDeTarjeta(page, /Libro IVA/i, /CONSULTAR/i, 'CONSULTAR de Libro IVA presentado');
  await page.waitForSelector('button[title="Ver"]', { timeout: 30000 });
  await page.waitForTimeout(1500);

  const filas = await listarPresentadas(page, periodoAAAAMM);
  if (!filas.length) {
    const e = new Error(`ARCA no tiene ninguna DDJJ de IVA presentada para ${periodoAAAAMM}.`);
    e.code = 'sin_presentacion';
    throw e;
  }
  const elegida = filas[0];
  log(`presentada: ${elegida.formulario} ${elegida.periodo} · ${elegida.secuencia} · ${elegida.presentada}` +
      (filas.length > 1 ? ` (hay ${filas.length} secuencias, uso la ultima)` : ''));

  const mmaaaa = elegida.periodo;
  const fila = page.locator('tr').filter({ has: page.locator('td', { hasText: new RegExp(`^\\s*${mmaaaa.replace('/', '\\/')}\\s*$`) }) }).first();
  const ver = fila.locator('button[title="Ver"]');
  const nueva = context.waitForEvent('page', { timeout: 20000 }).catch(() => null);
  await ver.click();
  let vista = await nueva;
  if (!vista) vista = page;
  await vista.waitForLoadState('domcontentloaded', { timeout: 60000 });
  await vista.waitForURL(/verVistaPrevia\.do/, { timeout: 60000 });
  await vista.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await vista.waitForTimeout(1500);

  // Totales de la vista previa: "Operaciones de Compras  <neto>  <exento>  <credito>".
  const texto = await vista.innerText('body');
  const linea = (re) => { const m = texto.match(re); return m ? { neto: aNumero(m[1]), exento: aNumero(m[2]), iva: aNumero(m[3]) } : null; };
  const n = '([\\d.]+,\\d{2})';
  const totales = {};
  for (const libro of ['Compras', 'Ventas']) {
    totales[libro.toLowerCase()] = {
      operaciones: linea(new RegExp(`Operaciones de ${libro}\\s+${n}\\s+${n}\\s+${n}`)),
      notasCredito: linea(new RegExp(`Notas de Cr.dito de ${libro}\\s+${n}\\s+${n}\\s+${n}`)),
    };
  }
  return { vista, elegida, secuencias: filas, totales };
}

// Desde la vista previa (o un libro abierto) al menu de la presentacion.
async function irAlMenuPresentacion(page) {
  await page.goto(page.url().replace(/(verVistaPrevia|verCompras|verVentas)\.do.*$/i, 'menuPresentacion.do'), { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.locator('a.panel, a, .panel').filter({ hasText: /Libro\s+(Compras|Ventas)/i }).first()
    .waitFor({ state: 'visible', timeout: 30000 });
}

module.exports = {
  IVA_URL, abrirPortalIva, representadoActual, asegurarRepresentacion, clickBotonDeTarjeta,
  abrirBorrador, datosIniciales, abrirLibro, importarDesdeArca, descargarCsv,
  clickSeguro, cerrarModales, SELECTORES_PROHIBIDOS,
  listarPresentadas, abrirPresentada, irAlMenuPresentacion,
};
