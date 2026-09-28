// bbva.js — BBVA Empresas (net cash; en Claves_Bancos.xlsx figura como FRANCES).
// Relevado el 24/09/2026.
//
//   #cod_emp (Codigo de empresa) · #cod_usu (Codigo de usuario) ·
//   #eai_password (Clave de acceso) · #btn_submit "Ingresar"
//
// Codigo de empresa = ACCESO/DNI (5 digitos; la pagina lo completa con ceros a 8) ·
// Codigo de usuario = USUARIO · Clave = CLAVE ID.

const comun = require('./comun');

const SPEC = {
  banco: 'BBVA',
  url: 'https://netcash.bbva.com.ar/local_pibee/SolicitarCredenciales.html',
  // La pagina vacia "Codigo de empresa" al terminar de inicializar (corrida del
  // 28/09/2026): se espera a que cargue del todo antes de completar.
  preparar: async (page) => {
    await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
    await page.waitForTimeout(2000);
  },
  pasos: [{
    campos: [
      { sel: '#cod_emp',      dato: 'acceso', nombre: 'ACCESO/DNI (codigo de empresa)' },
      { sel: '#cod_usu',      dato: 'usuario' },
      { sel: '#eai_password', dato: 'clave' },
    ],
    enviar: '#btn_submit',
  }],
  marcadorLogin: '#eai_password',
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'formulario (mapeo de columnas a confirmar)',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
