// hipotecario.js — BH Office Banking. Relevado el 24/09/2026.
//
//   input[name=Documento] (DNI) · input[name=Usuario] (tipo password) ·
//   input[name=Pass] (Contraseña) · #default "Iniciar sesión"
//
// Abre con un aviso de seguridad ("Entendido", #left-button) que se cierra antes.
// El DNI sale de la columna ACCESO/DNI de Claves_Bancos.xlsx.

const comun = require('./comun');

const SPEC = {
  banco: 'HIPOTECARIO',
  url: 'https://empresas.hipotecario.com.ar/OB/login',
  preparar: async (page) => {
    await page.waitForSelector('input[name="Documento"]', { timeout: 20000 }).catch(() => {});
    if (await comun.visible(page, '#left-button')) await page.click('#left-button').catch(() => {});
  },
  pasos: [{
    campos: [
      { sel: 'input[name="Documento"]', dato: 'acceso', nombre: 'ACCESO/DNI' },
      { sel: 'input[name="Usuario"]',   dato: 'usuario' },
      { sel: 'input[name="Pass"]',      dato: 'clave' },
    ],
    enviar: '#default',
  }],
  marcadorLogin: 'input[name="Pass"]',
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'formulario',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
