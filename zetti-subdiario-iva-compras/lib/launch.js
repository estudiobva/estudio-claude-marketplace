// ─────────────────────────────────────────────────────────────────────────────
// Vendorizado desde el plugin fisco-ar (MIT) — https://github.com/javiergradiche/fisco-ar-claude-plugin
// Commit de origen: 2ad06df · Copiado el 21/09/2026 al repo del Estudio BVA.
//
// Se copia en vez de depender del plugin para que una actualizacion del
// marketplace no pueda romper un cierre. Si ARCA cambia el login, el arreglo va
// ACA (y conviene mirar upstream por si ya lo resolvieron).
// ─────────────────────────────────────────────────────────────────────────────

// Opciones de browser COMPARTIDAS por todos los skills de ARCA.
//
// Parámetros globales (env vars; node las hereda del bot o de la shell):
//   PLAYWRIGHT_HEADLESS=false  → abre el browser VISIBLE para mirar mientras se prueba
//                          (default: true = headless, como en producción).
//   PLAYWRIGHT_SLOWMO=250      → ms de slowMo entre acciones para poder seguir el flujo
//                          (default: 0 en headless; 250 si es visible).
//
// Uso:
//   - Global: poné `PLAYWRIGHT_HEADLESS=false` en .env y reiniciá el bot.
//   - Puntual: `PLAYWRIGHT_HEADLESS=false node playwright_scripts/<skill>.js --cuit=... ...`
const HEADLESS = process.env.PLAYWRIGHT_HEADLESS !== 'false';
const SLOWMO =
  process.env.PLAYWRIGHT_SLOWMO != null && process.env.PLAYWRIGHT_SLOWMO !== ''
    ? Number(process.env.PLAYWRIGHT_SLOWMO)
    : (HEADLESS ? 0 : 250);

// Avisar por stderr cuando corre en modo visible (queda en los logs del skill).
if (!HEADLESS) {
  process.stderr.write(`[browser] modo VISIBLE (headless=false, slowMo=${SLOWMO}ms)\n`);
}

function launchOptions(extra = {}) {
  return {
    headless: HEADLESS,
    slowMo: SLOWMO,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    ...extra,
  };
}

module.exports = { HEADLESS, SLOWMO, launchOptions };
