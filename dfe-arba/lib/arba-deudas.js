// arba-deudas.js — situación de una sociedad en ARBA: deuda por objeto, DDJJ
// sin presentar, notificaciones DFE y, para los agentes de recaudación, el
// cruce período por período de lo declarado contra lo pagado.
//
// Solo LEE. No genera boletas, planes, liquidaciones ni presentaciones.
//
// De dónde sale cada dato:
//   - Gestionar mis planes (app.arba.gov.ar/Gestionar): al entrar, la propia
//     página pide services/padron/objetos (IIBB, agente R/P, sellos…) y
//     services/objeto/deuda por cada objeto de IIBB y de agente. Esa respuesta
//     trae deuda prejudicial/judicial, saldo a favor (total_excedente) y
//     cant_sin_ddjj (DDJJ que ARBA da por no presentadas). Se leen las
//     respuestas que pide la página; no hay captcha.
//     (La "Consulta de deuda" de AvisoDeudas pide captcha aun logueado: no se usa.)
//   - ARPresentaciones: DDJJ cerradas y pagos registrados de agente, por año
//     (más de 40 resultados por consulta los corta: por eso se consulta año a año).
//   - IBPresentaciones: DDJJ de anticipo cerradas (para listar los meses sin DDJJ).
//   - ARSuspension-BFF: estado del agente (Activo / Inactivo / Suspendido). Es
//     público, sin clave.
//
// El cartel "El Agente se encuentra suspendido" de A-122R NO sirve como dato:
// su componente lo muestra siempre que infoAgente responde bien (revisado en el
// código de A-122R el 26/10/2026). El estado real es el de ARSuspension.

const GESTIONAR = 'https://app.arba.gov.ar/Gestionar/dist/planespago';
const ARP = 'https://app.arba.gov.ar/ARPresentaciones/';
const IBP = 'https://app.arba.gov.ar/IBPresentaciones/';
const ESTADO_AGENTE = 'https://app.arba.gov.ar/ARSuspension-BFF/bff/estadoDeAgente';

// Códigos de impuesto de services/padron/objetos (enum de la app Gestionar).
const IMPUESTOS = {
  0: 'Inmobiliario', 1: 'Automotor', 2: 'Ingresos Brutos', 3: 'Embarcaciones', 4: 'Sellos',
  7: 'Agente de recaudación', 9: 'TGB', 10: 'Inmobiliario complementario', 22: 'Monotributo',
};

const limpio = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
const num = (s) => {
  const t = String(s == null ? '' : s).replace(/[^\d,.-]/g, '');
  return t ? Number(t.replace(/\./g, '').replace(',', '.')) : null;
};
const hostSso = (u) => { try { return /^(sso|idp)\./.test(new URL(String(u)).hostname); } catch { return false; } };

// ── Login ───────────────────────────────────────────────────────────────────
// ARBA tiene dos pantallas: el CAS viejo (sso.arba.gov.ar, #CUIT/#clave_Cuit)
// y el Keycloak nuevo (idp.arba.gov.ar, #username/#password). Si el contexto
// ya tiene sesión, el destino abre directo y no se toca nada.
async function ingresar(page, destino, { usuario, clave }, { timeout = 45000 } = {}) {
  await page.goto(destino, { waitUntil: 'domcontentloaded', timeout });
  if (!hostSso(page.url())) return;
  const campo = await page.waitForSelector('#CUIT:visible, #username:visible', { timeout: 25000 })
    .catch(() => null);
  if (!campo) {
    if (!hostSso(page.url())) return;
    const e = new Error('ARBA no mostró el formulario de ingreso'); e.code = 'arba_sin_login'; throw e;
  }
  const keycloak = await page.$('#username:visible');
  await page.fill(keycloak ? '#username' : '#CUIT', usuario);
  await page.fill(keycloak ? '#password' : '#clave_Cuit', clave);
  await Promise.all([
    page.waitForURL(u => !hostSso(u), { timeout: 30000 }).catch(() => {}),
    keycloak ? page.click('#kc-login') : page.locator('button[type=submit]').first().click(),
  ]);
  if (hostSso(page.url())) {
    const t = limpio(await page.innerText('body').catch(() => ''));
    const msg = (t.match(/(El usuario ingresado[^.]*\.|Usuario o contraseña incorrectos\.|[^.]*bloquead[^.]*\.)/i) || [])[1];
    const e = new Error(`ARBA no aceptó el ingreso: ${msg || t.slice(0, 160)}`);
    e.code = 'arba_login_rechazado';
    throw e;
  }
}

// ── Gestionar mis planes: deuda por objeto ──────────────────────────────────
async function leerSituacion(context, cred, { timeout = 30000 } = {}) {
  const page = await context.newPage();
  let padron = null;
  const deudas = [];
  page.on('response', async (r) => {
    const u = r.url();
    if (/\/Gestionar\/services\/padron\/objetos\//.test(u)) {
      try { padron = await r.json(); } catch { /* html de error */ }
    } else if (/\/Gestionar\/services\/objeto\/deuda\?/.test(u)) {
      const q = new URL(u).searchParams;
      let j;
      try { j = await r.json(); } catch { j = { message: `respuesta no JSON (HTTP ${r.status()})` }; }
      deudas.push({ impuestoConsulta: q.get('impuesto'), objetoConsulta: limpio(q.get('objeto')), ...j });
    }
  });
  try {
    await ingresar(page, GESTIONAR, cred);
    // La página pide la deuda de cada objeto de IIBB y de agente; se espera a
    // tenerlas todas (o al timeout: lo que falte queda como "sin respuesta").
    const t0 = Date.now();
    const esperadas = () => (padron?.objetos || []).filter(o => [2, 7].includes(o.cod_impuesto)).length;
    while (Date.now() - t0 < timeout) {
      if (padron && deudas.length >= esperadas()) break;
      await page.waitForTimeout(500);
    }
    await page.waitForTimeout(800);
    if (!padron) {
      const t = limpio(await page.innerText('body').catch(() => ''));
      const e = new Error(`Gestionar no devolvió los objetos del CUIT: ${t.slice(0, 160)}`);
      e.code = 'arba_sin_objetos';
      throw e;
    }
    const alertas = await page.evaluate(() => fetch('/Gestionar/services/alertas', { credentials: 'include' })
      .then(r => r.json()).catch(() => null)).catch(() => null);

    const objetos = (padron.objetos || []).map((o) => {
      const objeto = limpio(o.objeto);
      const d = deudas.find(x => limpio(x.objeto || x.objetoConsulta) === objeto);
      const agente = o.cod_impuesto === 7 ? (/R\d{4}$/.test(objeto) ? 'Retención' : /P\d{4}$/.test(objeto) ? 'Percepción' : '') : '';
      const fila = {
        codImpuesto: o.cod_impuesto,
        impuesto: IMPUESTOS[o.cod_impuesto] || `Impuesto ${o.cod_impuesto}`,
        agente,
        actividad: o.cod_impuesto === 7 ? String(Number(objeto.slice(-2))) : null,
        objeto,
        vinculo: limpio(o.vinculo_desc),
        vigenteDesde: o.vigente_desde ? String(o.vigente_desde) : null,
        cesado: o.fechaCesado || null,
        consultada: Boolean(d),
      };
      if (d) {
        if (d.status || d.message && d.deuda_total == null) {
          fila.error = limpio(d.message).replace(/actualizaci\S? n\b/, 'actualización') || limpio(d.status);
        } else {
          Object.assign(fila, {
            deudaPrejudicial: d.deuda_prejudicial, deudaJudicial: d.deuda_judicial,
            deudaTotal: d.deuda_total, saldoAFavor: d.total_excedente,
            ddjjSinPresentar: d.cant_sin_ddjj || 0, planDePago: Boolean(d.existe_plan_pago),
            planJudicial: Boolean(d.existe_plan_judicial), fechaCese: d.fecha_cese || null,
          });
        }
      }
      return fila;
    });
    return {
      razonSocial: limpio(padron.apellido_rsocial),
      dfe: padron.adhesion_dfe || null,
      objetos,
      alertas: Array.isArray(alertas) ? alertas.map(a => ({ descripcion: limpio(a.descripcion), cantidad: Number(a.cantidad) || 0 })) : [],
    };
  } finally {
    await page.close().catch(() => {});
  }
}

// ── Estado del agente (público) ─────────────────────────────────────────────
async function estadoAgente(request, cuit, hoy = new Date()) {
  const periodo = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}`;
  try {
    const r = await request.get(`${ESTADO_AGENTE}?cuit=${cuit}&periodo=${periodo}`, { timeout: 20000 });
    const j = await r.json();
    if (j.estado == null) return { error: limpio(j.message || (j.errors || []).join(' ')) || `HTTP ${r.status()}` };
    return { estado: limpio(j.estado), suspendido: Boolean(j.suspendido), periodo };
  } catch (e) {
    return { error: e.message.split('\n')[0] };
  }
}

// ── Períodos y vencimientos aproximados ─────────────────────────────────────
// "Q1 10/2026" | "Q2 09/2026" | "M 09/2026". Los vencimientos son aproximados
// (el calendario real depende del año): solo sirven para no marcar como deuda
// vencida la quincena/mes que todavía no venció.
function parsePeriodo(s) {
  const m = String(s).match(/(Q1|Q2|M)\s+(\d{2})\/(\d{4})/);
  return m ? { tipo: m[1], mes: Number(m[2]), anio: Number(m[3]) } : null;
}
const clavePeriodo = (p) => `${p.anio}${String(p.mes).padStart(2, '0')}${p.tipo === 'Q2' ? '2' : '1'}`;
function vencimientoAgente(p) {
  if (p.tipo === 'Q1') return new Date(p.anio, p.mes - 1, 26);
  return new Date(p.anio, p.mes, 12);   // Q2 y mensual: ~día 12 del mes siguiente
}

// ── ARPresentaciones: declarado vs pagado ───────────────────────────────────
async function leerTablaArp(page) {
  return page.evaluate(() => {
    const t = [...document.querySelectorAll('table')]
      .find(x => !x.querySelector('table') && x.rows.length > 1 && /Per[ií]odo/.test(x.rows[0].innerText));
    const filas = t ? [...t.rows].slice(1).map(r => [...r.cells].map(c => c.innerText.replace(/\s+/g, ' ').trim())) : [];
    const pags = [...new Set([...document.querySelectorAll('a')]
      .filter(a => /^\d+$/.test(a.innerText.trim()) && /d-\d+-p=/.test(a.getAttribute('href') || ''))
      .map(a => a.href))];
    return { filas, pags };
  });
}

async function consultaArp(page, pagina, boton, valor, anio, hoy) {
  await page.goto(ARP + pagina, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('#actividadIB', { timeout: 20000 });
  const hasta = anio === hoy.getFullYear() ? String(hoy.getMonth() + 1).padStart(2, '0') : '12';
  await page.fill('#anioDesde', String(anio));
  await page.selectOption('#mesDesde', '01');
  await page.fill('#anioHasta', String(anio));
  await page.selectOption('#mesHasta', hasta);
  await page.selectOption('#actividadIB', valor);
  await Promise.all([
    page.waitForLoadState('domcontentloaded').catch(() => {}),
    page.locator(`input[value="${boton}"]`).click(),
  ]);
  await page.waitForTimeout(1200);
  const { filas, pags } = await leerTablaArp(page);
  const todas = [...filas];
  for (const u of pags) {
    await page.goto(u, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(500);
    todas.push(...(await leerTablaArp(page)).filas);
  }
  return todas;
}

// Devuelve, por actividad de agente, los períodos con saldo (declarado − pagado
// > $1), los períodos vencidos sin DDJJ y las DDJJ iniciadas sin enviar.
// `codigos`: actividades de los objetos de agente de Gestionar (6, 7…): datos
// formales no lista las de un agente cesado. `activos`/`cesados`: actividades
// con DDJJ sin presentar según ARBA. Para un activo se listan todos los
// períodos vencidos sin DDJJ; para un cesado no se sabe hasta cuándo debía
// presentar, así que se informa el período siguiente a la última DDJJ.
async function saldosAgente(context, cred, { desdeAnio, hoy = new Date(), codigos = [], activos = [], cesados = [] } = {}) {
  const page = await context.newPage();
  page.on('dialog', d => d.dismiss().catch(() => {}));
  try {
    await ingresar(page, ARP + 'datosFormales.do', cred);
    await page.goto(ARP + 'datosFormales.do', { waitUntil: 'domcontentloaded' });
    const actividades = await page.evaluate(() => [...document.querySelectorAll('tr')]
      .map(r => [...r.cells].map(c => c.innerText.trim()))
      .filter(r => r.length === 4 && /^\d+$/.test(r[0]) && /\d{2}\/\d{2}\/\d{4}/.test(r[1])));
    await page.goto(ARP + 'pagosRegistrados.do', { waitUntil: 'domcontentloaded' });
    const opciones = await page.evaluate(() => [...document.querySelector('#actividadIB').options]
      .map(o => [o.value, o.text.trim()]));

    const periodos = [];
    const faltantes = [];
    const errores = [];
    const lista = actividades.filter(a => a[2] === '212');   // 212 = IIBB (los de Sellos van aparte)
    for (const c of codigos) if (!lista.some(a => a[0] === c)) lista.push([c, `01/01/${desdeAnio}`, '212', '']);
    for (const [codigo, inicio] of lista) {
      const op = opciones.find(o => o[1].startsWith(`${codigo} - `));
      if (!op) { errores.push(`actividad ${codigo} sin opción en ARPresentaciones`); continue; }
      const anioInicio = Number(inicio.slice(-4));
      const ddjj = [];
      const pagos = [];
      for (let anio = Math.max(desdeAnio, anioInicio); anio <= hoy.getFullYear(); anio++) {
        for (const f of await consultaArp(page, 'djCerradas.do', 'Novedades', op[0], anio, hoy)) {
          ddjj.push({ periodo: f[2], comprobante: f[3], declarado: num(f[4]), rectificativa: Number(f[5]) || 0 });
        }
        for (const f of await consultaArp(page, 'pagosRegistrados.do', 'Enviar', op[0], anio, hoy)) {
          pagos.push({ periodo: f[0], concepto: f[2], fecha: f[3], importe: num(f[4]), compensado: f[5] });
        }
      }
      // La DDJJ vigente de cada período es la de rectificativa más alta.
      const vigente = new Map();
      for (const d of ddjj) {
        const v = vigente.get(d.periodo);
        if (!v || d.rectificativa >= v.rectificativa) vigente.set(d.periodo, d);
      }
      const pagado = new Map();
      const ultimoPago = new Map();
      for (const p of pagos) {
        if (!/obligaci/i.test(p.concepto)) continue;
        pagado.set(p.periodo, (pagado.get(p.periodo) || 0) + (p.importe || 0));
        ultimoPago.set(p.periodo, p.fecha);
      }
      const nombre = op[1].replace(/^\d+ - /, '');
      for (const [per, d] of vigente) {
        const pg = pagado.get(per) || 0;
        const saldo = Math.round(((d.declarado || 0) - pg) * 100) / 100;
        if (saldo <= 1) continue;             // ARBA redondea el pago a 10 centavos
        const p = parsePeriodo(per);
        const venc = p ? vencimientoAgente(p) : null;
        const hoy0 = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
        periodos.push({
          actividad: codigo, regimen: nombre, periodo: per, orden: p ? clavePeriodo(p) : per,
          comprobante: d.comprobante, rectificativa: d.rectificativa, declarado: d.declarado,
          pagado: Math.round(pg * 100) / 100, saldo, ultimoPago: ultimoPago.get(per) || null,
          vencimientoAprox: venc ? venc.toISOString().slice(0, 10) : null,
          estado: venc && venc >= hoy0 ? 'A vencer' : 'Vencido',
        });
      }
      if (cesados.includes(codigo)) {
        const ultima = [...vigente.keys()].map(parsePeriodo).filter(Boolean)
          .sort((a, b) => clavePeriodo(a).localeCompare(clavePeriodo(b))).pop();
        if (ultima) {
          const quincenal = ultima.tipo !== 'M';
          const sig = ultima.tipo === 'Q1' ? { tipo: 'Q2', mes: ultima.mes, anio: ultima.anio }
            : { tipo: quincenal ? 'Q1' : 'M', mes: ultima.mes === 12 ? 1 : ultima.mes + 1, anio: ultima.mes === 12 ? ultima.anio + 1 : ultima.anio };
          const per = `${sig.tipo} ${String(sig.mes).padStart(2, '0')}/${sig.anio}`;
          faltantes.push({ impuesto: 'Agente de recaudación', actividad: codigo, regimen: nombre, periodo: per,
                           orden: clavePeriodo(sig), pago: pagado.get(per) || null,
                           nota: `Agente cesado: última DDJJ ${[...vigente.keys()].find(k => parsePeriodo(k) && clavePeriodo(parsePeriodo(k)) === clavePeriodo(ultima))}; probablemente falta la del período del cese` });
        } else {
          faltantes.push({ impuesto: 'Agente de recaudación', actividad: codigo, regimen: nombre, periodo: `antes de ${desdeAnio}`,
                           orden: '0', nota: `Agente cesado sin DDJJ desde ${desdeAnio}: la que falta es anterior; buscarla en ARPresentaciones` });
        }
      }
      // Períodos vencidos sin DDJJ cerrada de un agente activo.
      if (activos.includes(codigo)) {
        const quincenal = ddjj.some(d => /^Q[12]/.test(d.periodo));
        const desde = new Date(Math.max(desdeAnio, anioInicio), anioInicio >= desdeAnio ? Number(inicio.slice(3, 5)) - 1 : 0, 1);
        for (let f = new Date(desde); f <= hoy; f.setMonth(f.getMonth() + 1)) {
          const mm = String(f.getMonth() + 1).padStart(2, '0');
          for (const tipo of quincenal ? ['Q1', 'Q2'] : ['M']) {
            const per = `${tipo} ${mm}/${f.getFullYear()}`;
            const p = parsePeriodo(per);
            if (vencimientoAgente(p) >= hoy || vigente.has(per)) continue;
            faltantes.push({ impuesto: 'Agente de recaudación', actividad: codigo, regimen: nombre,
                             periodo: per, orden: clavePeriodo(p), pago: pagado.get(per) || null });
          }
        }
      }
    }
    await page.goto(ARP + 'djPendientes.do', { waitUntil: 'domcontentloaded' });
    const txt = limpio(await page.innerText('body').catch(() => ''));
    const iniciadas = /No se recuperaron declaraciones juradas pendientes/i.test(txt) ? []
      : await leerTablaArp(page).then(r => r.filas.map(f => f.join(' | '))).catch(() => []);
    periodos.sort((a, b) => a.actividad.localeCompare(b.actividad) || a.orden.localeCompare(b.orden));
    faltantes.sort((a, b) => a.actividad.localeCompare(b.actividad) || a.orden.localeCompare(b.orden));
    return { actividades: actividades.map(a => ({ codigo: a[0], inicio: a[1], impuesto: a[2], regimen: a[3] })),
             periodos, faltantes, iniciadas, errores };
  } finally {
    await page.close().catch(() => {});
  }
}

// ── IBPresentaciones: meses sin DDJJ de anticipo ────────────────────────────
// Solo se llama cuando ARBA marca DDJJ de IIBB sin presentar. Busca de a 12
// meses (es el máximo que acepta) y lista los meses vencidos sin DDJJ cerrada.
async function faltantesIibb(context, cred, { desde, hoy = new Date() } = {}) {
  const page = await context.newPage();
  page.on('dialog', d => d.dismiss().catch(() => {}));
  try {
    await ingresar(page, IBP + 'preBuscarDDJJCerradas.do', cred);
    const presentados = new Set();
    // Último anticipo seguro vencido: el de dos meses atrás (vencen ~día 20 del mes siguiente).
    const fin = new Date(hoy.getFullYear(), hoy.getMonth() - (hoy.getDate() >= 25 ? 1 : 2), 1);
    for (let ini = new Date(desde); ini <= fin; ini.setMonth(ini.getMonth() + 12)) {
      const tope = new Date(Math.min(new Date(ini.getFullYear(), ini.getMonth() + 11, 1), fin));
      await page.goto(IBP + 'preBuscarDDJJCerradas.do', { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(600);
      await page.click('#ui-id-2');
      await page.waitForTimeout(300);
      await page.evaluate(([a, b, c, d]) => {
        const set = (id, v) => { const e = document.getElementById(id); e.value = v; e.dispatchEvent(new Event('change', { bubbles: true })); };
        set('anioMensualDesde', a); set('mesDesde', b); set('anioMensualHasta', c); set('mesHasta', d);
      }, [String(ini.getFullYear()), String(ini.getMonth() + 1), String(tope.getFullYear()), String(tope.getMonth() + 1)]);
      await page.click('#buscar');
      await page.waitForTimeout(2000);
      const filas = await page.evaluate(() => {
        try { return window.jQuery('#resultados').dataTable().fnGetData().map(r => r.map(x => String(x).replace(/<[^>]+>/g, '').trim())); }
        catch { return []; }
      });
      for (const r of filas) presentados.add(`${r[3]}${String(r[4]).padStart(2, '0')}`);
    }
    if (!presentados.size) {
      // Ninguna DDJJ en todo el rango: no son N meses omitidos, es que el CUIT
      // no presenta por acá (casi siempre Convenio Multilateral → SIFERE).
      return [{ impuesto: 'Ingresos Brutos', periodo: `desde ${String(desde.getMonth() + 1).padStart(2, '0')}/${desde.getFullYear()}`,
                orden: '0', nota: 'No hay ninguna DDJJ de anticipo en IBPresentaciones: probablemente presenta por Convenio Multilateral (SIFERE); verificar' }];
    }
    const faltan = [];
    for (let f = new Date(desde); f <= fin; f.setMonth(f.getMonth() + 1)) {
      const k = `${f.getFullYear()}${String(f.getMonth() + 1).padStart(2, '0')}`;
      if (!presentados.has(k)) faltan.push({ impuesto: 'Ingresos Brutos', periodo: `${k.slice(4)}/${k.slice(0, 4)}`, orden: k });
    }
    return faltan;
  } finally {
    await page.close().catch(() => {});
  }
}

module.exports = { ingresar, leerSituacion, estadoAgente, saldosAgente, faltantesIibb, IMPUESTOS };
