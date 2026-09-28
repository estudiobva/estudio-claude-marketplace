// galicia.js — Galicia Office Banking (empresas). Relevado el 24/09/2026.
//
//   #userInput (Usuario) · #userPassword (Clave) · button "Ingresar"
//
// Carga un antifraude (iframe portals3.galicia.ar/.../lwsa.html): correr visible.
// Credenciales: login propio de la sociedad (bloque Galicia) o login unico.

const comun = require('./comun');

const SPEC = {
  banco: 'GALICIA',
  url: 'https://empresas.bancogalicia.com.ar/login',
  pasos: [{
    campos: [
      { sel: '#userInput',    dato: 'usuario' },
      { sel: '#userPassword', dato: 'clave' },
    ],
    enviar: 'button:visible:has-text("Ingresar")',   // sin [type=submit]: el boton no trae el atributo (relevado 24/09/2026)
  }],
  marcadorLogin: '#userPassword',
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'formulario',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
