// ─────────────────────────────────────────────────────────────────────────────
// Vendorizado desde el plugin fisco-ar (MIT) — https://github.com/javiergradiche/fisco-ar-claude-plugin
// Commit de origen: 2ad06df · Copiado el 21/09/2026 al repo del Estudio BVA.
//
// Se copia en vez de depender del plugin para que una actualizacion del
// marketplace no pueda romper un cierre. Si ARCA cambia el login, el arreglo va
// ACA (y conviene mirar upstream por si ya lo resolvieron).
// ─────────────────────────────────────────────────────────────────────────────

// args.js — parser de los argumentos --clave=valor que usan todos los scripts.
//
//   node scripts/arca-f931.js --cuit=20123456789 --periodo=05/2026
//   → { cuit: '20123456789', periodo: '05/2026' }
//
// Un flag sin `=` queda como cadena vacía, así que `'debug' in args` sigue
// funcionando para los booleanos.

function parseArgs(argv = process.argv.slice(2)) {
  return Object.fromEntries(
    argv.map(a => {
      const [k, ...v] = a.replace(/^--/, '').split('=');
      return [k, v.join('=')];
    })
  );
}

// Imprime el JSON de error que espera quien invoca el script y corta con exit 1.
function fail(message, extra = {}) {
  console.log(JSON.stringify({ error: message, ...extra }));
  process.exit(1);
}

module.exports = { parseArgs, fail };
