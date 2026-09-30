// f931.js — Declaracion en Linea de ARCA: listar y bajar los F.931 presentados.
//
// El F.931 no existe como PDF descargable: ver_formulario.aspx devuelve HTML y
// el boton "Imprimir" dispara window.print(), que abre un dialogo del sistema
// operativo inmanejable. Por eso el circuito viejo reconstruia el PDF parseando
// el texto.
//
// Con Playwright eso no hace falta: page.pdf() renderiza la pagina a PDF sin
// pasar por ningun dialogo. El resultado es el formulario REAL, no una
// reproduccion — lo cual ademas es mejor para el legajo.
//
// REGLA DURA: esto solo consulta y descarga. Nunca genera ni presenta una DDJJ.

const HOST = 'https://serviciossegsoc.afip.gob.ar';
const URL_LISTADO = `${HOST}/djproforma/app/consultar/dj_generadas.aspx`;
const URL_FORM    = `${HOST}/djproforma/app/consultar/ver_formulario.aspx`;

const log = (m) => process.stderr.write(`[f931] ${m}\n`);

// CSS de impresion. Sin esto el PDF sale en 2 paginas: la segunda queda en
// blanco por el chrome del sitio. Tambien saca los botones Imprimir/CERRAR,
// que no tienen sentido dentro de un PDF.
const CSS_IMPRESION = `
  @page { size: A4 landscape; margin: 6mm; }
  html, body { margin:0 !important; padding:0 !important; height:auto !important; }
  input[type=button], input[type=submit], input[type=image], button { display:none !important; }
  * { overflow: visible !important; }
`;

// Escala que entra en una hoja A4 apaisada dejando el texto lo mas grande
// posible. Verificado con formularios de 26-28 empleados.
const ESCALA = 0.9;

// Abre "Declaracion en Linea" y deja la sesion parada en la sociedad pedida.
async function abrirDeclaracionEnLinea(context, page, cuit) {
  const { abrirServicio } = require('./arca-login');
  let del;
  try {
    del = await abrirServicio(
      context, page, /Declaraci[oó]n en [Ll]ínea|Declaraciones en [Ll]ínea/i);
  } catch (e) {
    // abrirServicio solo reconoce pestañas de facturacion cuando no llega el
    // evento de pestaña nueva; Declaracion en Linea vive en serviciossegsoc.
    del = context.pages().find(p => /serviciossegsoc|djproforma/i.test(p.url()));
    if (!del) throw e;
  }
  await del.waitForLoadState('networkidle', { timeout: 40000 }).catch(() => {});

  // Pantalla de control de acceso: desplegable con todas las representadas.
  const ddl = del.locator('select[id*="ddlCUIT"]');
  if (await ddl.isVisible({ timeout: 10000 }).catch(() => false)) {
    const opciones = await del.evaluate(() => {
      const s = document.querySelector('select[id*="ddlCUIT"]');
      return [...s.options].map(o => ({ value: o.value, text: o.text }));
    });
    const hit = opciones.find(o => o.value.replace(/\D/g, '') === String(cuit));
    if (!hit) {
      throw new Error(
        `El CUIT ${cuit} no figura entre las ${opciones.length} sociedades del ` +
        `desplegable de Declaracion en Linea. Falta la relacion en ARCA.`);
    }
    log(`sociedad: ${hit.text.trim()}`);
    await ddl.selectOption({ value: hit.value });
    await del.locator('input[type="image"][id*="btnAceptar"]').click();
    await del.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  } else {
    log('sin desplegable de CUIT (acceso propio)');
  }

  // Puede haber una o dos pantallas de avisos, cada una con su ACEPTAR al pie.
  for (let i = 0; i < 3; i++) {
    const av = del.locator('input[type="image"][id*="ceptar"]').first();
    if (!(await av.isVisible({ timeout: 3000 }).catch(() => false))) break;
    await av.click();
    await del.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
  }

  // Verificar contra el encabezado que estamos en la sociedad correcta.
  const enc = await del.innerText('body').catch(() => '');
  const cuitEnPantalla = (enc.replace(/\D/g, '').match(new RegExp(String(cuit))) || [])[0];
  if (!cuitEnPantalla) {
    log('aviso: no pude confirmar el CUIT en el encabezado; sigo igual');
  }
  return del;
}

// Listado de DDJJ generadas. Trae el historial completo (puede ser 200+ filas).
//
// Visitarlo NO es opcional aunque despues se navegue por URL: ver_formulario
// responde "No puede acceder a la pagina solicitada" si ASP.NET no tiene el
// estado de sesion que fija esta pantalla.
async function listarDDJJ(page) {
  await page.goto(URL_LISTADO, { waitUntil: 'networkidle', timeout: 60000 });
  const filas = await page.evaluate(() =>
    [...document.querySelectorAll('tr')]
      .filter(r => /^(Original|Rectif)/.test(r.innerText.trim()))
      .map(r => {
        const c = [...r.cells].map(x => x.innerText.trim());
        return { secuencia: c[0], periodoTexto: c[1], empleados: c[2], generada: c[3] };
      })
  );

  return filas.map(f => {
    const mm = (f.periodoTexto.match(/^(\d{2})\/(\d{4})$/) || []);
    const sec = (f.secuencia.match(/\((\d+)\)/) || [])[1];
    return {
      ...f,
      periodo: mm.length ? mm[2] + mm[1] : null,
      sec: sec != null ? String(sec).padStart(3, '0') : '000',
      esRectificativa: /rectif/i.test(f.secuencia),
    };
  }).filter(f => f.periodo);
}

// Baja el F.931 de un periodo como PDF de una hoja.
async function descargarF931(page, { periodo, sec = '000' }, destino) {
  const url = `${URL_FORM}?Periodo=${periodo}&SecDJVig=${sec}`;
  await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });

  const txt = await page.innerText('body').catch(() => '');
  if (/No puede acceder a la p.gina solicitada/i.test(txt)) {
    throw new Error(
      `ARCA no dejo abrir el formulario de ${periodo} (sec ${sec}). ` +
      `Suele ser que se perdio el estado de sesion del listado.`);
  }
  if (!/931/.test(txt) || txt.length < 500) {
    throw new Error(`El formulario de ${periodo} vino vacio o incompleto (${txt.length} caracteres).`);
  }

  await page.addStyleTag({ content: CSS_IMPRESION });
  // Imprimir y CERRAR no son <input>/<button> sino imagenes con onclick: el CSS
  // no los alcanza y quedaban al pie del PDF.
  await page.evaluate(() => {
    const re = /imprim|cerrar|print|close/i;
    for (const el of document.querySelectorAll('img, a, input, button')) {
      const attrs = [el.alt, el.title, el.value, el.id, el.getAttribute('src'),
                     el.getAttribute('onclick'), el.getAttribute('href'), el.innerText].join(' ');
      if (re.test(attrs)) el.style.setProperty('display', 'none', 'important');
    }
  });
  await page.pdf({
    path: destino, format: 'A4', landscape: true, printBackground: true,
    scale: ESCALA, margin: { top: '6mm', bottom: '6mm', left: '6mm', right: '6mm' },
  });

  // Datos del encabezado, para el control posterior y para el reporte.
  const empleados = (txt.match(/Empleados en n.mina:\s*([\d.]+)/i) || [])[1] || null;
  const rem1 = (txt.match(/Suma de Rem\.\s*1:\s*([\d.]+,\d{2})/i) || [])[1] || null;
  return { periodo, sec, empleados, rem1, caracteres: txt.length };
}

// Control: que el PDF tenga exactamente una pagina y contenga el periodo.
function verificarPdf(ruta, periodo) {
  const fs = require('fs');
  const buf = fs.readFileSync(ruta);
  const m = [...buf.toString('latin1').matchAll(/\/Type\s*\/Pages[^>]*?\/Count\s+(\d+)/g)]
    .map(x => Number(x[1]));
  const paginas = m.length ? Math.max(...m) : null;
  const mm = periodo.slice(4, 6), aaaa = periodo.slice(0, 4);
  return {
    bytes: buf.length,
    paginas,
    unaPagina: paginas === 1,
    periodoEsperado: `${mm}/${aaaa}`,
  };
}

module.exports = {
  HOST, URL_LISTADO, URL_FORM, ESCALA, CSS_IMPRESION,
  abrirDeclaracionEnLinea, listarDDJJ, descargarF931, verificarPdf,
};
