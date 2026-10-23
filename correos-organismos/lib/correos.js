// correos.js — lee los correos electronicos declarados de una sociedad en ARCA
// y ARBA. Solo LEE: nunca toca "Modificar correo", "Registrar" ni "Borrar".
//
// ARCA (login con clave fiscal, una vez por apoderado):
//   - Sistema Registral, portada del representado
//     (ResponsiveIndexInternetAction.do?idPersona=<CUIT>): el correo principal
//     con su estado (VERIFICADO / ...), el telefono y si esta ADHERIDO al DFE.
//   - Sistema Registral > Consulta > Datos registrales (la "lista sabana"): TODOS
//     los correos con su tipo, los telefonos y la leyenda del DFE (a veces dice
//     "debe asociar su email y telefono al servicio del DFE").
//   - Domicilio Fiscal Electronico > "Asociar temas a un correo electronico":
//     una tabla por apoderado con el correo al que llegan los avisos del DFE de
//     cada representado. Se lee una vez por login.
//
// ARBA (login con CUIT + CIT, uno por sociedad, contexto nuevo cada vez):
//   - Autogestion > Datos de contacto. La pantalla es un formulario; los datos
//     salen de la respuesta JSON de services/consultaDC (correos con caracter y
//     fecha de verificacion, telefonos) y de services/gen/datosUsuario (el
//     correo con el que se registro la CIT, que suele ser otro).
//
// AGIP no se automatiza: su firewall bloquea el navegador automatizado ("Web
// Page Blocked", Attack ID 20000051) y no se intenta esquivar.

const SETI = 'https://seti.afip.gob.ar/padron-puc-consulta-internet';
const ARBA_CONTACTO = 'https://app.arba.gov.ar/DatosContacto/dist/';

const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/;
const limpio = (s) => String(s || '').replace(/\s+/g, ' ').trim();

// ── ARCA ────────────────────────────────────────────────────────────────────

// Abre un servicio del portal por el titulo exacto de su tarjeta. abrirServicio
// de arca-login busca por texto parcial y con "Sistema Registral" puede pegarle
// a otra tarjeta que lo menciona en la descripcion.
async function abrirServicioExacto(context, portal, titulo, { timeout = 30000 } = {}) {
  const tarjeta = () => portal.locator('h3', { hasText: titulo }).first();
  if (!(await tarjeta().isVisible({ timeout: 3000 }).catch(() => false))) {
    await portal.locator('text=Ver todos').first().click({ timeout: 10000 }).catch(() => {});
    await portal.waitForTimeout(2000);
  }
  try {
    const [p] = await Promise.all([
      context.waitForEvent('page', { timeout }),
      tarjeta().click({ timeout: 10000 }),
    ]);
    await p.waitForLoadState('domcontentloaded', { timeout });
    return p;
  } catch (e) {
    const err = new Error(`No pude abrir ${titulo}: puede no estar delegado a este login. ` +
                          `Detalle: ${e.message.split('\n')[0]}`);
    err.code = 'servicio_no_disponible';
    throw err;
  }
}

async function abrirRegistral(context, portal) {
  const p = await abrirServicioExacto(context, portal, /^\s*SISTEMA REGISTRAL\s*$/i);
  await p.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  return p;
}

// Portada del representado: correo principal + estado, telefono, DFE.
async function leerPortadaRegistral(reg, cuit) {
  await reg.goto(`${SETI}/ResponsiveIndexInternetAction.do?idPersona=${cuit}`, { timeout: 45000 });
  await reg.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  let marco = null;
  for (let i = 0; i < 20 && !marco; i++) {
    marco = reg.frames().find(f => /ContenidoInternet/.test(f.url()));
    if (!marco) await reg.waitForTimeout(500);
  }
  if (!marco) {
    const t = limpio(await reg.innerText('body').catch(() => ''));
    const e = new Error(/no corresponde al usuario/i.test(t)
      ? 'El CUIT no esta entre los representados de este login en Sistema Registral'
      : `Sistema Registral no mostro la portada: ${t.slice(0, 160)}`);
    e.code = /no corresponde al usuario/i.test(t) ? 'cuit_no_disponible' : 'registral_sin_portada';
    throw e;
  }
  await marco.waitForSelector('text=Datos de contacto', { timeout: 20000 }).catch(() => {});
  const t = await marco.innerText('body');

  const dfeBoton = await marco.locator('a[href*="/padron-puc-consulta-internet/dfe"]').first()
    .innerText({ timeout: 3000 }).catch(() => '');

  // Bloque "Datos de contacto" … "MODIFICAR CORREO" … "MODIFICAR TELÉFONO"
  const i = t.indexOf('Datos de contacto');
  const jc = t.indexOf('MODIFICAR CORREO', i);
  const jt = t.indexOf('MODIFICAR TEL', jc);
  const bloqueCorreo = i >= 0 && jc > i ? t.slice(i, jc) : '';
  const bloqueTel = jc >= 0 && jt > jc ? t.slice(jc + 16, jt) : '';

  let correo = '', estado = '';
  for (const linea of bloqueCorreo.split('\n')) {
    const m = linea.match(EMAIL_RE);
    if (m) { correo = m[0]; estado = limpio(linea.replace(m[0], '')); break; }
  }
  const telefonos = [...new Set(bloqueTel.split('\n').map(limpio).filter(x => /\d{6,}/.test(x)))];

  const estadoCuit = (t.match(/Estado de la CUIT\s*\n\s*([^\n]+)/) || [])[1] || '';
  return {
    correo, estado, telefonos,
    dfe: limpio(dfeBoton) || null,     // ADHERIDO / ADHERIRME / …
    estadoCuit: limpio(estadoCuit),
    marco,
  };
}

// Lista sabana (Consulta > Datos registrales). Es Vaadin: carga en diferido.
async function leerSabana(reg, marco) {
  await marco.locator('a[href="#2"]').first().click({ timeout: 10000 });
  await marco.locator('a[onclick="DoPost();"]').first().click({ timeout: 10000 });
  let s = null;
  for (let i = 0; i < 40 && !s; i++) {
    await reg.waitForTimeout(500);
    s = reg.frames().find(f => /listasabana-web/.test(f.url()));
  }
  if (!s) { const e = new Error('No abrio Datos registrales (lista sabana)'); e.code = 'sabana_no_abre'; throw e; }

  // Esperar a que desaparezcan los "Cargando e-mails/telefonos…"
  let texto = '';
  for (let i = 0; i < 60; i++) {
    texto = await s.innerText('body').catch(() => '');
    if (/E-Mails/.test(texto) && !/Cargando (e-mails|tel)/i.test(texto)) break;
    await reg.waitForTimeout(500);
  }
  if (/Cargando (e-mails|tel)/i.test(texto)) {
    const e = new Error('La lista sabana no termino de cargar los correos'); e.code = 'sabana_timeout'; throw e;
  }

  // Cada fila es un layout horizontal; sus celdas se toman de los textos hoja
  // (sin elementos hijos), porque innerText de la fila pega mail y tipo.
  const filas = await s.evaluate(() => {
    const out = [];
    document.querySelectorAll('.v-slot-row > .v-horizontallayout').forEach(row => {
      const hojas = [...row.querySelectorAll('*')]
        .filter(e => !e.children.length && e.textContent.trim())
        .map(e => e.textContent.replace(/\s+/g, ' ').trim())
        .filter(t => /[\p{L}\p{N}@]/u.test(t));   // afuera los iconos (glifos)
      out.push(hojas);
    });
    return out;
  });
  const correos = [], telefonos = [];
  const vistos = new Set();
  for (const c of filas) {
    const k = c.join('|');
    if (vistos.has(k)) continue;
    vistos.add(k);
    if (c.length >= 2 && /@/.test(c[0])) correos.push({ correo: c[0], tipo: c[1] || '' });
    else if (c.length === 2 && /^[\d\s()+-]{6,}$/.test(c[0]) && !/\d/.test(c[1])) telefonos.push({ numero: c[0], tipo: c[1] });
  }
  // Respaldo por si Vaadin cambia la estructura: sacar los mails del texto.
  if (!correos.length) {
    const a = texto.indexOf('E-Mails'), z = texto.indexOf('Teléfonos', a);
    const bloque = a >= 0 ? texto.slice(a, z > a ? z : undefined) : '';
    for (const linea of bloque.split('\n')) {
      const m = linea.match(EMAIL_RE);
      if (m) correos.push({ correo: m[0].replace(/(COMERCIAL|PERSONAL|LEGAL|FISCAL|PARTICULAR|LABORAL|OTRO)$/, ''), tipo: '' });
    }
  }
  const dfe = limpio((texto.match(/adherido al domicilio fiscal Electr[oó]nico:\s*([^\n]+)/i) || [])[1]);
  return { correos, telefonos, dfe };
}

// Relevamiento ARCA completo de un CUIT (con Sistema Registral ya abierto).
async function relevarArca(reg, cuit) {
  const portada = await leerPortadaRegistral(reg, cuit);
  const { marco, ...p } = portada;
  let sabana = null, errorSabana = null;
  try { sabana = await leerSabana(reg, marco); }
  catch (e) { errorSabana = e.message; }
  return { ...p, sabana, errorSabana };
}

// DFE > "Asociar temas a un correo electronico": { '<CUIT>': [{correo, tipo, tema}] }
async function leerAsociacionesDfe(context, portal) {
  const dfe = await abrirServicioExacto(context, portal, /^\s*DOMICILIO FISCAL ELECTR[OÓ]NICO\s*$/i);
  try {
    await dfe.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    const menu = dfe.getByText('Asociar temas a un correo electrónico').first();
    await menu.waitFor({ state: 'visible', timeout: 30000 });
    // El cartel #dfeModal ("Notificaciones / Mensajes …") aparece un rato
    // despues de cargar y tapa el menu. Se cierra con "Recordar más tarde", que
    // no deja ninguna preferencia guardada en la cuenta (a diferencia de
    // "Entendido").
    for (let i = 0; i < 3; i++) {
      await dfe.waitForTimeout(1500);
      const cerrar = dfe.locator('#dfeModal button', { hasText: /Recordar más tarde/i });
      if (await cerrar.isVisible().catch(() => false)) {
        await cerrar.click().catch(() => {});
        await dfe.locator('#dfeModal').waitFor({ state: 'hidden', timeout: 5000 }).catch(() => {});
      }
      try { await menu.click({ timeout: 5000 }); break; }
      catch (e) { if (i === 2) throw e; }
    }
    await dfe.waitForSelector('text=Asociaciones vigentes', { timeout: 20000 });
    await dfe.waitForTimeout(1500);
    const sels = dfe.locator('select#per-page-select');
    for (let i = 0; i < await sels.count(); i++) {
      if (await sels.nth(i).isVisible()) await sels.nth(i).selectOption('100').catch(() => {});
    }
    await dfe.waitForTimeout(2000);

    const mapa = {};
    for (let pagina = 1; pagina <= 20; pagina++) {
      const filas = await dfe.evaluate(() => [...document.querySelectorAll('table tbody tr')]
        .map(r => [...r.querySelectorAll('td')].map(c => c.innerText.replace(/\s+/g, ' ').trim()))
        .filter(r => /^\d{2}-\d{8}-\d$/.test(r[1] || '')));
      for (const r of filas) {
        const cuit = r[1].replace(/\D/g, '');
        const m = (r[2] || '').match(/^(.*?)\s*(?:\(([^)]*)\))?$/);
        (mapa[cuit] ||= []).push({ correo: (m ? m[1] : r[2]).trim(), tipoCorreo: m && m[2] || '',
                                    tipo: r[3] || '', tema: r[4] || '' });
      }
      // Paginador (b-pagination): "›" / "Go to next page"
      const sig = dfe.locator('[aria-label="Go to next page"]:not([aria-disabled="true"]), ' +
                              'li.page-item:not(.disabled) > [aria-label*="next" i]').last();
      if (!(await sig.isVisible({ timeout: 1000 }).catch(() => false))) break;
      const antes = JSON.stringify(filas.slice(0, 2));
      await sig.click().catch(() => {});
      await dfe.waitForTimeout(1500);
      const despues = await dfe.evaluate(() => [...document.querySelectorAll('table tbody tr')].slice(0, 6)
        .map(r => r.innerText).join('|'));
      if (!despues || despues.includes(antes)) break;
    }
    return mapa;
  } finally {
    await dfe.close().catch(() => {});
  }
}

// ── ARBA ────────────────────────────────────────────────────────────────────

const CARACTER_ARBA = { 1: 'Contribuyente', 2: 'Contable/Impositivo/Legal', 3: 'Administración', 4: 'Otros' };

// Loguea con la CIT en un contexto NUEVO (el SSO de ARBA recuerda la sesion) y
// devuelve lo que responde services/consultaDC y services/gen/datosUsuario.
async function relevarArba(browser, { cuit, usuario, clave }, { timeout = 45000 } = {}) {
  const context = await browser.newContext({ locale: 'es-AR', viewport: { width: 1366, height: 900 } });
  const page = await context.newPage();
  const respuestas = {};
  page.on('response', async (r) => {
    const u = r.url();
    const clave_ = /DatosContacto\/services\/consultaDC/.test(u) ? 'consultaDC'
      : /DatosContacto\/services\/gen\/datosUsuario/.test(u) ? 'datosUsuario' : null;
    if (!clave_ || r.status() !== 200) return;
    try { respuestas[clave_] = await r.json(); } catch { /* 302 o html */ }
  });
  try {
    await page.goto(ARBA_CONTACTO, { waitUntil: 'domcontentloaded', timeout });
    await page.waitForSelector('#CUIT', { timeout: 20000 });
    await page.fill('#CUIT', usuario || cuit);
    await page.fill('#clave_Cuit', clave);
    await Promise.all([
      page.waitForURL(u => !/sso\.arba\.gov\.ar/.test(String(u)), { timeout: 25000 }).catch(() => {}),
      page.locator('#frmDatos button[type=submit], button[type=submit]').first().click(),
    ]);
    if (/sso\.arba\.gov\.ar/.test(page.url())) {
      const t = limpio(await page.innerText('body').catch(() => ''));
      const msg = (t.match(/Ingresá con tu Clave de Identificación Tributaria \(CIT\)\s*(.*?)\s*Completá el formulario/) || [])[1] || t.slice(0, 160);
      const e = new Error(`ARBA no acepto el ingreso: ${msg || 'sin mensaje'}`);
      e.code = 'arba_login_rechazado';
      throw e;
    }
    for (let i = 0; i < 40 && !respuestas.consultaDC; i++) await page.waitForTimeout(500);
    if (!respuestas.consultaDC) {
      const t = limpio(await page.innerText('body').catch(() => ''));
      const e = new Error(`ARBA no devolvio los datos de contacto (${page.url().split('?')[0]}): ${t.slice(0, 160)}`);
      e.code = 'arba_sin_datos';
      throw e;
    }
    // El texto de los caracteres sale del propio formulario si esta.
    const caracteres = await page.evaluate(() => {
      const s = document.querySelector('select');
      return s ? Object.fromEntries([...s.options].map(o => [o.value, o.text])) : {};
    }).catch(() => ({}));

    const dc = respuestas.consultaDC;
    if (dc.exito && dc.exito !== 'S') {
      const e = new Error(`ARBA respondio error: ${dc.mensaje || JSON.stringify(dc.errors)}`);
      e.code = 'arba_error'; throw e;
    }
    const s = dc.salida || {};
    const tel = (l) => (l || []).filter(x => x && x.telefono && x.telefono !== '0')
      .map(x => `${x.area && x.area !== '0' ? x.area + '-' : ''}${x.telefono}`);
    const nombreCaracter = (c) => {
      const t = caracteres[String(c)];
      return t && !/Ã/.test(t) ? t : (CARACTER_ARBA[c] || `Carácter ${c}`);
    };
    const props = Object.fromEntries(((respuestas.datosUsuario || {}).properties || []).map(p => [p.key, p.value]));
    return {
      razonSocial: limpio(s.razonSocial),
      correos: (s.emails || []).map(x => ({
        correo: limpio(x.email),
        caracter: nombreCaracter(x.caracter),
        verificado: x.verificado ? new Date(x.verificado).toISOString().slice(0, 10) : null,
      })),
      telefonosFijos: tel(s.telefonosFijos),
      telefonosMoviles: tel(s.telefonosMoviles),
      correoCit: props.mail || null,
      telefonoCit: props.telephonenumber || null,
    };
  } finally {
    await context.close().catch(() => {});
  }
}

module.exports = {
  abrirServicioExacto, abrirRegistral, relevarArca, leerAsociacionesDfe, relevarArba,
  EMAIL_RE, CARACTER_ARBA,
};
