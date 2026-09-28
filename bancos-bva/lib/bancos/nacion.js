// nacion.js — BNA Digital (BNA+), la plataforma que figura en Claves_Bancos.xlsx
// (link "BNA Digital" -> https://digital.bna.com.ar/loginStep1). Relevado el 24/09/2026.
//
// Login en DOS pasos:
//   1. #document (DNI, sale de ACCESO/DNI) · #username (Usuario) ·
//      boton "Continuar" (id="global.continue", arranca deshabilitado hasta
//      completar los campos)
//   2. clave — no se pudo ver sin enviar un usuario real: se busca el primer
//      campo de clave visible y un boton Ingresar/Continuar/Aceptar. Si la
//      pantalla es otra, termina en login_sin_confirmar y se ajusta aca.
//
// Antes (hasta el 24/09/2026) este adaptador apuntaba a Nacion Empresa 24 3.0
// (bee3.redlink.com.ar/bna3), que es OTRA plataforma, con reCAPTCHA: no es la
// que usa el estudio.

const comun = require('./comun');

const DNI = (c) => String(c.acceso || '').replace(/\D/g, '');
const CLAVE = 'input[type="password"]:visible';

const SPEC = {
  banco: 'NACION',
  url: 'https://digital.bna.com.ar/loginStep1',
  pasos: [
    {
      campos: [
        { sel: '#document', dato: DNI, nombre: 'ACCESO/DNI' },
        { sel: '#username', dato: 'usuario' },
      ],
      // El id tiene puntos: se usa como atributo para no escaparlo en CSS.
      enviar: 'button[id="global.continue"]',
    },
    { campos: [{ sel: CLAVE, dato: 'clave' }], enviar: comun.boton('Ingresar', 'Continuar', 'Aceptar') },
  ],
  marcadorLogin: CLAVE,
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'paso 1',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
