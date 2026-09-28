// macro.js — Banca Internet Empresas. Relevado el 24/09/2026.
//
// Login en DOS pasos:
//   1. #textField1 (usuario, tipo password) · #processCustomerLogin "CONTINUAR"
//   2. clave — vista en la primera corrida real (24/09/2026):
//      muestra la imagen de seguridad del usuario, "Ingresá tu clave" y el boton
//      "INGRESÁ" (con acento: por eso comun.boton prueba las variantes con voseo).
//
// Nota conocida: desloguea al cambiar de sociedad, hay que loguear por sociedad.

const comun = require('./comun');

const CLAVE = 'input[type="password"]:visible:not(#textField1)';

const SPEC = {
  banco: 'MACRO',
  url: 'https://www.macro.com.ar/biempresas/',
  pasos: [
    { campos: [{ sel: '#textField1', dato: 'usuario' }], enviar: '#processCustomerLogin' },
    { campos: [{ sel: CLAVE, dato: 'clave' }], enviar: comun.boton('Ingresar', 'Continuar', 'Aceptar') },
  ],
  marcadorLogin: CLAVE,
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'pasos 1 y 2 (corrida real)',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
