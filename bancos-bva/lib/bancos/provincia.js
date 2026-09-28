// provincia.js — BIP Empresas (Banca Internet Provincia). Relevado el 24/09/2026.
//
//   #username (Tu usuario) · #password (Tu clave) · button "Ingresar"
//
// La pantalla de login vive en /spa/ (sin "login" en la URL): el ingreso se
// detecta porque desaparece el campo de clave, no por la URL.

const comun = require('./comun');

const SPEC = {
  banco: 'PROVINCIA',
  url: 'https://www.bancoprovincia.bancainternet.com.ar/spa-empresas/',
  pasos: [{
    campos: [
      { sel: '#username', dato: 'usuario' },
      { sel: '#password', dato: 'clave' },
    ],
    enviar: 'button:visible:has-text("Ingresar")',   // sin [type=submit]: el boton no trae el atributo (relevado 24/09/2026)
  }],
  marcadorLogin: '#password',
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'formulario',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
