#!/usr/bin/env node
/**
 * carpeta-sociedad.js: imprime la carpeta de la sociedad en la unidad compartida.
 *
 *   node scripts/carpeta-sociedad.js <CUIT>
 *
 * Usa la misma resolucion de rutas que el resto de los plugins del estudio:
 * BVA_UNIDAD_PATH del .env, o la unidad compartida por defecto.
 * La logica esta en lib/rutas-bva.js del plugin portal-iva-descarga.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

function plugin(nombre) {
  const base = path.join(os.homedir(), '.claude', 'plugins', 'cache', 'estudiobva-skills', nombre);
  if (fs.existsSync(base)) {
    const vers = fs.readdirSync(base).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    if (vers.length) return path.join(base, vers[vers.length - 1]);
  }
  return path.join(os.homedir(), '.claude', 'plugins', 'marketplaces', 'estudiobva-skills', nombre);
}

const P = plugin('portal-iva-descarga');
const { loadEnv } = require(path.join(P, 'lib', 'env'));
const rutas = require(path.join(P, 'lib', 'rutas-bva'));
loadEnv();
const cuit = String(process.argv[2] || '').replace(/\D/g, '');
if (cuit.length !== 11) { console.error('Uso: node carpeta-sociedad.js <CUIT>'); process.exit(2); }
console.log(rutas.carpetaSociedad(cuit));
