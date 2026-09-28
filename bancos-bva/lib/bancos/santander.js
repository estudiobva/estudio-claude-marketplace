// santander.js — Online Banking Empresas. Relevado el 24/09/2026.
//
//   #user (Usuario) · #password (Contraseña) · #pe-login-btningresar
//
// La portada ya habla de "Token de seguridad": por eso comun.js mira solo el
// texto nuevo que aparece despues de enviar.

const comun = require('./comun');

const SPEC = {
  banco: 'SANTANDER',
  url: 'https://empresas.santander.com.ar/login',
  pasos: [{
    campos: [
      { sel: '#user',     dato: 'usuario' },
      { sel: '#password', dato: 'clave' },
    ],
    enviar: '#pe-login-btningresar',
  }],
  marcadorLogin: '#password',
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'formulario',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
