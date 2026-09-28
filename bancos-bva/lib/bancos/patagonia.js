// patagonia.js — Patagonia eBank Empresas. Relevado el 24/09/2026.
//
// Login en DOS pasos:
//   1. #selectFieldIdentificationType (CDI/CUIL/CUIT/DI Extranjero/PASAPORTE) ·
//      #textFieldDocumento (numero) · #textFieldUsuario (usuario, tipo password) ·
//      #login "Continuar"
//   2. clave — no se pudo ver sin enviar un usuario real (mismo criterio que Macro).
//
// El numero sale de ACCESO/DNI; el tipo se deduce: 11 digitos que empiezan con
// 30/33/34 = CUIT, otros 11 digitos = CUIL. Con otro formato corta y avisa.

const comun = require('./comun');

const tipoDoc = (cred) => {
  const d = String(cred.acceso || '').replace(/\D/g, '');
  if (d.length !== 11) return null;
  return /^3[034]/.test(d) ? 'CUIT' : 'CUIL';
};

const CLAVE = 'input[type="password"]:visible:not(#textFieldUsuario)';

const SPEC = {
  banco: 'PATAGONIA',
  url: 'https://ebankempresas.bancopatagonia.com.ar/desktop-webserver/sso',
  pasos: [
    {
      campos: [
        { sel: '#selectFieldIdentificationType', dato: tipoDoc, tipo: 'select',
          nombre: 'tipo de documento (ACCESO/DNI con 11 digitos: CUIL o CUIT)' },
        { sel: '#textFieldDocumento', dato: (c) => String(c.acceso || '').replace(/\D/g, ''), nombre: 'ACCESO/DNI' },
        { sel: '#textFieldUsuario',   dato: 'usuario' },
      ],
      enviar: '#login',
    },
    { campos: [{ sel: CLAVE, dato: 'clave' }], enviar: comun.boton('Ingresar', 'Continuar', 'Aceptar') },
  ],
  marcadorLogin: CLAVE,
};

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'paso 1',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts) };
