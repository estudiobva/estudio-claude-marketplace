#!/usr/bin/env node
// ─────────────────────────────────────────────────────────────────────────────
// Vendorizado desde el plugin fisco-ar (MIT) — https://github.com/javiergradiche/fisco-ar-claude-plugin
// Commit de origen: 2ad06df · Copiado el 21/09/2026 al repo del Estudio BVA.
//
// Se copia en vez de depender del plugin para que una actualizacion del
// marketplace no pueda romper un cierre. Si ARCA cambia el login, el arreglo va
// ACA (y conviene mirar upstream por si ya lo resolvieron).
// ─────────────────────────────────────────────────────────────────────────────

/**
 * claves-dump.js — helper interno de claves.js.
 *
 * Descifra y parsea el archivo de claves y escupe el resultado como JSON por
 * stdout. Existe sólo porque el descifrado es asincrónico y los scripts
 * resuelven sus credenciales de forma sincrónica al arrancar: `claves.loadSync()`
 * corre este archivo en un subproceso y se queda con el JSON.
 *
 * No es una interfaz pública: para consultar el archivo desde la línea de
 * comandos está `node scripts/claves.js` (que enmascara las claves por default).
 */

require('./env').loadEnv();
const claves = require('./claves');

claves.load()
  .then((data) => { process.stdout.write(JSON.stringify(data)); })
  .catch((e) => { process.stdout.write(JSON.stringify({ error: e.message })); process.exit(1); });
