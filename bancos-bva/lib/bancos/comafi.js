// comafi.js — eBanking Empresas. Relevado el 24/09/2026.
//
//   Usuario · Clave de Acceso · Clave de Identificacion · "Ingresar"
//   (los campos no tienen id: se ubican por placeholder)
//
// Es el banco que explica las columnas de Claves_Bancos.xlsx:
//   USUARIO = Usuario · ACCESO/DNI = Clave de Acceso · CLAVE ID = Clave de Identificacion.

const comun = require('./comun');

const SPEC = {
  banco: 'COMAFI',
  url: 'https://ebanking.comafiempresas.com.ar/login',
  pasos: [{
    campos: [
      { sel: 'input[placeholder="Usuario"]',                 dato: 'usuario' },
      { sel: 'input[placeholder="Clave de Acceso"]',         dato: 'acceso', nombre: 'ACCESO/DNI (clave de acceso)' },
      { sel: 'input[placeholder="Clave de Identificación"]', dato: 'clave',  nombre: 'CLAVE ID (clave de identificacion)' },
    ],
    enviar: 'button:visible:has-text("Ingresar")',   // sin [type=submit]: el boton no trae el atributo (relevado 24/09/2026)
  }],
  marcadorLogin: 'input[placeholder="Clave de Identificación"]',
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'formulario',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
