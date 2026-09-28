#!/usr/bin/env node
/**
 * claves-bancos.js — consulta Claves_Bancos.xlsx (cifrado) SIN mostrar claves.
 *
 *   node scripts/claves-bancos.js                          panorama general
 *   node scripts/claves-bancos.js --banco=COMAFI
 *   node scripts/claves-bancos.js --banco=GALICIA --sociedad="<sociedad>"
 *   node scripts/claves-bancos.js --faltantes               filas incompletas
 *
 * Por defecto las credenciales salen ENMASCARADAS (`***(8)` = 8 caracteres).
 * `--revelar` las muestra en claro: usalo solo si sabes quien esta mirando la
 * pantalla, y nunca pegues el resultado en un chat o un archivo.
 */

const { parseArgs, fail } = require('../lib/args');
const { loadEnv } = require('../lib/env');
const cb = require('../lib/claves-bancos');

loadEnv();
const args = parseArgs();
const REVELAR = 'revelar' in args;
const mask = (v) => (REVELAR ? v : (v ? `***(${String(v).length})` : ''));

(async () => {
  try {
    if (args['banco']) {
      const c = await cb.buscar({ banco: args['banco'], sociedad: args['sociedad'] });
      console.log(JSON.stringify({
        banco: c.banco, sociedad: c.sociedad, farmacias: c.farmacias,
        usuario: mask(c.usuario), acceso: mask(c.acceso), clave: mask(c.clave),
        link: c.link, origen: c.origen,
      }, null, 2));
      if (!REVELAR) process.stderr.write('\n(credenciales enmascaradas; --revelar las muestra)\n');
      return;
    }

    const { bancos, galicia, avisoGalicia, archivo } = await cb.load();
    if ('faltantes' in args) {
      const b = bancos.filter(x => !x.usuario || !x.clave).map(x => x.banco);
      const g = galicia.filter(x => !x.usuario || !x.clave).map(x => x.sociedad);
      console.log(JSON.stringify({ bancosIncompletos: b, sociedadesGaliciaIncompletas: g }, null, 2));
      return;
    }

    console.log(JSON.stringify({
      archivo,
      bancosConLoginUnico: bancos.map(x => ({
        banco: x.banco, farmacias: x.farmacias,
        usuario: mask(x.usuario), clave: mask(x.clave), link: x.link,
      })),
      galiciaPorSociedad: {
        total: galicia.length,
        conCredenciales: galicia.filter(x => x.usuario && x.clave).length,
        sinCredenciales: galicia.filter(x => !x.usuario || !x.clave).map(x => x.sociedad),
      },
      avisoGalicia,
    }, null, 2));
    if (!REVELAR) process.stderr.write('\n(credenciales enmascaradas; --revelar las muestra)\n');
  } catch (e) {
    fail(e.message);
  }
})();
