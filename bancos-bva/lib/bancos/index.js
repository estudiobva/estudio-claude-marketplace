// bancos/index.js — un adaptador de login por banco.
//
// Cada adaptador exporta { banco, LOGIN_URL, login(page, cred, opts) } armado
// sobre lib/bancos/comun.js. Opcionalmente puede exportar
// `descargar(page, { sociedad, periodos, guardar })` para bajar los resumenes
// sin intervencion; mientras no lo tenga, scripts/extractos-banco.js corre en
// modo asistido (la persona navega, el script toma, renombra y archiva).
//
// Para sumar un banco: crear lib/bancos/<banco>.js y registrarlo aca.

const ADAPTADORES = {
  GALICIA:     require('./galicia'),
  COMAFI:      require('./comafi'),
  SANTANDER:   require('./santander'),
  HIPOTECARIO: require('./hipotecario'),
  MACRO:       require('./macro'),
  NACION:      require('./nacion'),
  BBVA:        require('./bbva'),
  PATAGONIA:   require('./patagonia'),
  PROVINCIA:   require('./provincia'),
};

// Nombres alternativos que aparecen en el Maestro, en Claves_Bancos.xlsx y en
// las carpetas de la unidad.
const ALIAS = { FRANCES: 'BBVA', 'BBVA FRANCES': 'BBVA', 'NACION ARGENTINA': 'NACION',
                'BH': 'HIPOTECARIO', 'BAPRO': 'PROVINCIA' };

// Como se llama cada banco en Claves_Bancos.xlsx, cuando difiere.
const NOMBRE_EN_CLAVES = { BBVA: 'FRANCES' };

const norm = (s) => String(s || '').toUpperCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/^(BANCO|BCO\.?)\s+/, '').replace(/[^A-Z0-9]+/g, ' ').trim();

function canonico(banco) {
  const b = norm(banco);
  return ALIAS[b] || b;
}

function adaptador(banco) {
  const b = canonico(banco);
  const a = ADAPTADORES[b];
  if (!a) {
    const e = new Error(
      `No hay adaptador de login para ${b || '(sin banco)'}. ` +
      `Disponibles: ${Object.keys(ADAPTADORES).join(', ')}.`);
    e.code = 'banco_sin_adaptador';
    throw e;
  }
  return a;
}

const nombreEnClaves = (banco) => NOMBRE_EN_CLAVES[canonico(banco)] || canonico(banco);

module.exports = { adaptador, canonico, nombreEnClaves, norm, bancos: Object.keys(ADAPTADORES) };
