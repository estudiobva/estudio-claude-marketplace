// arba-dfe.js — bandeja del Domicilio Fiscal Electrónico (DFE) de ARBA.
//
// Solo LEE el listado. NUNCA abre "Visualizar": abrir una comunicación le pone
// "Fecha leído" y, si todavía no estaba notificada, puede darla por notificada
// antes de tiempo. El listado ya trae todo lo que se necesita para controlar:
// fechas, impuesto, operativo y título, destinatario, expediente y carácter.
//
// Dónde está:
//   app.arba.gov.ar/SIIntimaciones/dfe/inicioDomicilioFiscalElectronico.do
//   (es el destino de "Notificaciones DFE" en Gestionar/services/alertas).
//   La página arma una tabla DataTables por tipo de comunicación
//   (#listaNotificacionesTC1 = Intimaciones, #listaNotificacionesTC2 =
//   Notificaciones, …) y una #listaNotificacionesTCTodas. Las columnas ocultas
//   traen el estado ("TC2NotificadaSILeidaNO") y las fechas en AAAAMMDD.
//   Tipos: Intimaciones y Notificaciones se notifican formalmente (tienen
//   "Fecha notificado"); Comunicaciones son avisos informativos (p. ej. "Baja
//   Retención A-122-R") que ARBA no notifica nunca.
//   Se leen TODAS las filas con la API de DataTables (no solo la página visible).

const { ingresar } = require('./arba-deudas');

const DFE = 'https://app.arba.gov.ar/SIIntimaciones/dfe/inicioDomicilioFiscalElectronico.do?Frame=';

const limpio = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
const aIso = (s) => (/^\d{8}$/.test(String(s)) ? `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}` : null);

async function leerBandeja(context, cred, { timeout = 45000 } = {}) {
  const page = await context.newPage();
  try {
    try {
      await ingresar(page, DFE, cred, { timeout });
    } catch (e) {
      // Las sociedades que además son agentes de recaudación pasan por
      // "Seleccione un rol" después de la clave (no es un rechazo). La bandeja
      // del DFE es la misma con cualquier rol: se elige Contribuyente.
      if (!/Seleccione un rol/i.test(e.message)) throw e;
      const sel = page.locator('select[name=rol]');
      const opciones = await sel.locator('option').evaluateAll(os => os.map(o => ({ v: o.value, t: o.textContent.trim() })));
      const op = opciones.find(o => /Contribuyente/i.test(o.t)) || opciones[0];
      if (op) await sel.selectOption(op.v);
      await Promise.all([
        page.waitForURL(u => !/^https:\/\/(sso|idp)\./.test(String(u)), { timeout: 30000 }).catch(() => {}),
        page.locator('button[type=submit]').first().click(),
      ]);
      if (/^https:\/\/(sso|idp)\./.test(page.url())) {
        const err = new Error('ARBA pidió elegir rol y no avanzó al DFE'); err.code = 'arba_rol'; throw err;
      }
    }
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    // DataTables se inicializa en document.ready; hasta entonces el bloque está oculto.
    await page.waitForFunction(() => {
      const ts = [...document.querySelectorAll("table[id^='listaNotificacionesTC']")];
      return ts.length && window.jQuery && ts.every(t => jQuery.fn.dataTable.isDataTable(t));
    }, null, { timeout: 20000 }).catch(() => {});

    const datos = await page.evaluate(() => {
      const txt = (h) => { const d = document.createElement('div'); d.innerHTML = h == null ? '' : String(h); return d.innerText.replace(/\s+/g, ' ').trim(); };
      const href = (h) => { const m = String(h || '').match(/href="([^"]+)"/); return m ? m[1].replace(/&amp;/g, '&') : null; };
      const nombres = {};
      document.querySelectorAll('#tabs ul li a[href^="#tabs-"]').forEach(a => {
        nombres[a.getAttribute('href').slice(6)] = a.innerText.replace(/\(\d+\)/, '').replace(/\s+/g, ' ').trim();
      });
      const hayTablas = !!document.querySelector("table[id^='listaNotificacionesTC']");
      const tablas = [...document.querySelectorAll("table[id^='listaNotificacionesTC']")]
        .filter(t => t.id !== 'listaNotificacionesTCTodas')
        .map(t => {
          const tc = t.id.replace('listaNotificacionesTC', '');
          const dt = window.jQuery && jQuery.fn.dataTable.isDataTable(t) ? jQuery(t).DataTable() : null;
          const filas = dt ? dt.rows().data().toArray() : [];
          return { tc, tipo: nombres[tc] || `Tipo ${tc}`, filas: filas.map(f => ({
            filtro: txt(f[0]),
            disp: String(f[2] || ''), notif: String(f[4] || ''), leido: String(f[6] || ''),
            impuesto: txt(f[8]), operativo: txt(f[9]), destinatario: txt(f[10]),
            expediente: txt(f[11]), caracter: txt(f[12]), link: href(f[13]),
          })) };
        });
      const contribuyente = (document.body.innerText.match(/Contribuyente:\s*([^\n]+)/) || [])[1] || '';
      return { hayTablas, tablas, contribuyente: contribuyente.replace(/\s+/g, ' ').trim(),
               texto: document.body.innerText.replace(/\s+/g, ' ').trim().slice(0, 400) };
    });

    // Bandeja sin ninguna comunicación: ARBA muestra solo la bienvenida.
    if (!datos.hayTablas && /Gracias por acceder a su Domicilio Fiscal/i.test(datos.texto)) {
      return { contribuyente: datos.contribuyente, comunicaciones: [], vacia: true };
    }
    if (!datos.hayTablas) {
      const e = new Error(`La bandeja del DFE no mostró el listado: ${datos.texto.slice(0, 200)}`);
      e.code = /adher/i.test(datos.texto) ? 'dfe_sin_adhesion' : 'dfe_sin_listado';
      throw e;
    }

    const comunicaciones = [];
    for (const t of datos.tablas) {
      for (const f of t.filas) {
        const q = f.link ? new URL(f.link, 'https://app.arba.gov.ar').searchParams : null;
        const m = f.operativo.match(/^(\d+)\s*-\s*(\d+)\s+(.*)$/);
        comunicaciones.push({
          tipo: t.tipo,
          puestaADisposicion: aIso(f.disp),
          notificado: aIso(f.notif),
          leido: aIso(f.leido),
          estaLeida: /LeidaSI/i.test(f.filtro) || Boolean(aIso(f.leido)),
          estaNotificada: /NotificadaSI/i.test(f.filtro) || Boolean(aIso(f.notif)),
          impuesto: f.impuesto,
          operativo: q ? q.get('nroOperativo') : (m ? m[1] : ''),
          nroNotificacion: q ? q.get('nroNotificacion') : (m ? m[2] : ''),
          titulo: limpio(m ? m[3] : f.operativo),
          destinatario: f.destinatario,
          expediente: f.expediente,
          caracter: f.caracter,
        });
      }
    }
    comunicaciones.sort((a, b) => String(b.puestaADisposicion).localeCompare(String(a.puestaADisposicion)));
    return { contribuyente: datos.contribuyente, comunicaciones };
  } finally {
    await page.close().catch(() => {});
  }
}

module.exports = { DFE, leerBandeja };
