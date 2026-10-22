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

// Nombre de los extractos: <cuenta>_<AAAAMMDD>_extractos.pdf, con la fecha de
// EMISION. El resumen mensual se emite los primeros dias del mes siguiente
// (01/09/2026 = agosto) y el extracto diario de una caja de ahorro lleva la fecha
// del dia (03/08/2026 = agosto): con dia <= 5 no se sabe cual es, se pregunta.
// Relevado el 29/09/2026 (una sociedad con CC mensual y CA diaria).
function periodoDeArchivo(nombre) {
  const m = String(nombre || '').match(/(?:^|[^\d])\d{9,12}_(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])_extractos/i);
  if (!m) return undefined;   // otro formato: lo lee la regla general (lib/extractos.js)
  const [a, mes, dia] = [+m[1], +m[2], +m[3]];
  const actual = `${a}${m[2]}`;
  if (dia > 5) return { periodo: actual, dia };
  const prev = mes === 1 ? `${a - 1}12` : `${a}${String(mes - 1).padStart(2, '0')}`;
  return { ambiguo: `Provincia lo emitio el ${m[3]}/${m[2]}/${a}: si es el resumen MENSUAL corresponde a ${prev}; ` +
    `si es un extracto DIARIO, a ${actual}.` };
}

module.exports = { banco: SPEC.banco, LOGIN_URL: SPEC.url, SPEC, verificado: 'formulario',
  login: (page, cred, opts) => comun.login(page, SPEC, cred, opts), periodoDeArchivo };
