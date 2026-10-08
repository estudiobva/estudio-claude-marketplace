// zetti.js — T&S Web (Zetti): login, entidades por CUIT y reportes de FTWeb.
//
// T&S Web es una app React sobre una API REST propia:
//   - login con usuario y clave en #/login (OAuth2 por detras; el token queda
//     en localStorage 'zweb-auth-common|AUTH|authData').
//   - /api-rest/user/me trae las entidades del usuario (~140): Sociedades
//     ("SOC <nombre>", nivel 5) y sus Locales (nivel 6),
//     sin CUIT.
//   - /api-rest/v2/<id>/nodes/<id> trae el CUIT de cada entidad.
//   - Los reportes son de FTWeb (la app vieja) y se abren con
//     /ftweb/loginOAuth?...&formName=<reporte>&access_token=...&node=<nombre>,
//     igual que hace el menu de T&S Web. No hace falta elegir la entidad en
//     pantalla: el reporte toma el nodo del parametro node.
//
// Para una sociedad se usa la entidad de mas arriba con ese CUIT (la Sociedad):
// el nodo de la Sociedad abarca todos sus locales (hay sociedades con mas de 10).

const BASE = (process.env.ZETTI_URL || 'http://zwebx.ddns.net:8085').replace(/\/+$/, '');
const AUTH_KEY = 'zweb-auth-common|AUTH|authData';

const normCuit = (s) => String(s || '').replace(/\D/g, '');

// Login en T&S Web. Devuelve { page, auth } con la pagina parada en
// "Select entity" y el token de la sesion.
async function login(ctx, { usuario, clave }, log = () => {}) {
  const page = await ctx.newPage();
  log(`login en ${BASE}/tysweb como ${usuario}`);
  await page.goto(`${BASE}/tysweb/home/#/login`, { waitUntil: 'networkidle', timeout: 90000 });
  await page.locator('input[type=text]').first().fill(usuario);
  await page.locator('input[type=password]').fill(clave);
  await page.getByRole('button', { name: 'Login' }).click();

  const ok = await page.waitForURL(/#\/select-entity/, { timeout: 60000 }).then(() => true).catch(() => false);
  if (!ok) {
    const texto = (await page.locator('body').innerText().catch(() => '')).replace(/\s+/g, ' ').slice(0, 200);
    const e = new Error(`T&S Web no paso del login (sigue en ${page.url()}). Pantalla: ${texto}`);
    e.code = /incorrect|invalid|inv[aá]lid|bad credentials/i.test(texto) ? 'clave_rechazada' : 'login';
    throw e;
  }
  const auth = await page.evaluate((k) => JSON.parse(localStorage.getItem(k) || 'null'), AUTH_KEY);
  if (!auth || !auth.accessToken) throw Object.assign(new Error('T&S Web no dejo el token de la sesion en localStorage.'), { code: 'login' });
  return { page, auth };
}

// Entidades del usuario con su CUIT: [{ id, nombre, corto, nivel, codigo, cuit, razon }].
async function entidades(page) {
  return page.evaluate(async (k) => {
    const a = JSON.parse(localStorage.getItem(k));
    const h = { Authorization: 'Bearer ' + a.accessToken };
    const me = await (await fetch('/api-rest/user/me', { headers: h })).json();
    const out = [];
    const cola = [...me.entidades];
    // De a 8 en paralelo: ~140 pedidos, unos segundos.
    await Promise.all(Array.from({ length: 8 }, async () => {
      while (cola.length) {
        const e = cola.shift();
        let n = {};
        try { n = await (await fetch(`/api-rest/v2/${e.id}/nodes/${e.id}`, { headers: h })).json(); } catch { /* sin datos */ }
        out.push({
          id: e.id, nombre: e.nombre, corto: e.nombreCorto, nivel: e.nivelNombre,
          codigo: e.codigoJerarquico, cuit: String(n.cuit || '').replace(/\D/g, ''), razon: n.companyName || '',
        });
      }
    }));
    return out.sort((x, y) => x.codigo.localeCompare(y.codigo));
  }, AUTH_KEY);
}

// La entidad de una sociedad: de las que tienen ese CUIT, la de mas arriba en
// el arbol (su codigo jerarquico es prefijo del de todas las demas).
function entidadDeCuit(lista, cuit) {
  const c = normCuit(cuit);
  const hits = lista.filter((e) => e.cuit === c);
  if (!hits.length) {
    throw Object.assign(new Error(`Ninguna entidad de T&S Web tiene el CUIT ${c}.`), { code: 'sin_entidad' });
  }
  const raiz = hits.find((e) => hits.every((o) => o === e || o.codigo.startsWith(e.codigo + '.')));
  if (!raiz) {
    const e = new Error(
      `El CUIT ${c} esta en ${hits.length} entidades de T&S Web sin una que las abarque a todas: ` +
      hits.map((h) => `${h.nombre} (${h.codigo})`).join(' | ') + '. Pasa --entidad="<nombre>" para elegir.'
    );
    e.code = 'entidad_ambigua';
    throw e;
  }
  return { entidad: raiz, locales: hits.filter((h) => h !== raiz) };
}

// Abre un formulario de reporte de FTWeb para una entidad (por su nombre).
async function abrirReporte(ctx, auth, formName, entidad) {
  const q = new URLSearchParams({
    servlet: 'eventhandler', event: 'showfrmreport', formName,
    access_token: auth.accessToken, expires_in: '30000', node: entidad.nombre,
    refresh_token: auth.refreshToken || '', token_type: 'bearer',
  });
  const f = await ctx.newPage();
  await f.goto(`${BASE}/ftweb/loginOAuth?${q}`, { waitUntil: 'networkidle', timeout: 90000 });
  const nodo = await f.locator('input[name=nodo]').inputValue({ timeout: 30000 }).catch(() => null);
  if (nodo === null) {
    const texto = (await f.locator('body').innerText().catch(() => '')).replace(/\s+/g, ' ').slice(0, 200);
    await f.close();
    throw Object.assign(new Error(`FTWeb no abrio el reporte ${formName} para ${entidad.nombre}. Pantalla: ${texto}`), { code: 'reporte' });
  }
  if (nodo.trim() !== entidad.codigo) {
    await f.close();
    throw Object.assign(new Error(
      `FTWeb abrio el reporte con el nodo ${nodo} y la entidad ${entidad.nombre} es ${entidad.codigo}: no sigo.`
    ), { code: 'reporte' });
  }
  return f;
}

// 5.6.5 Reporte de IVA ventas, en CSV, de desde a hasta (DD/MM/AAAA). Guarda el
// CSV en `destino` y devuelve la ruta.
async function subdiarioIvaVentas(ctx, auth, entidad, { desde, hasta }, destino) {
  const f = await abrirReporte(ctx, auth, 'rpt_subdiario_iva_venta', entidad);
  try {
    await f.fill('input[name=fecEDesde]', desde);
    await f.fill('input[name=fecEHasta]', hasta);
    await f.fill('input[name=periodo]', '');
    await f.selectOption('select[name=tipo_tipo]', { label: 'Todos' });
    await f.selectOption('select[name=orden]', { label: 'Tipo Comprobante' });
    await f.selectOption('select[name=tipo_codif]', { label: 'Mostrar todos' });
    for (const n of ['agruparFeB', 'zeta', 'zetanc', 'resumen', 'detallar', 'afectarZ', 'anulNoFiscal', 'hideHeader']) {
      await f.locator(`input[name=${n}]`).setChecked(false);
    }
    await f.selectOption('select[name=tipo]', 'csv');

    const descarga = ctx.waitForEvent('download', { timeout: 300000 });
    await f.getByText('Aceptar', { exact: true }).click();
    const dl = await descarga.catch(() => null);
    if (!dl) {
      const texto = (await f.locator('body').innerText().catch(() => '')).replace(/\s+/g, ' ').slice(0, 300);
      throw Object.assign(new Error(`El reporte no bajo ningun archivo en 5 minutos. Pantalla: ${texto}`), { code: 'sin_descarga' });
    }
    await dl.saveAs(destino);
    return destino;
  } finally {
    await f.close().catch(() => {});
  }
}

module.exports = { BASE, login, entidades, entidadDeCuit, abrirReporte, subdiarioIvaVentas };
