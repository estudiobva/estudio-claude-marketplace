---
name: setup-bva
description: "Revisa y arregla el entorno para usar los plugins del marketplace estudiobva-skills (planes-facilidades-arca, conciliacion-ncr, anticipos-sct, deudas-ddjj-sct, bancos-bva, portal-iva-descarga y los que se sumen): corre un diagnostico, instala las dependencias que falten (npm, Chromium de Playwright, openpyxl) despues de una actualizacion, instala los plugins nuevos del marketplace y explica como preparar una computadora nueva con el instalador de Windows. Usar SIEMPRE que pidan \"configurar las skills del estudio\", \"instalar los plugins del estudio\", \"setup\", \"revisa el setup\", \"no me anda el script\", \"Cannot find module playwright\", \"falta officecrypto-tool\", \"No encontre la clave fiscal\", \"no encuentro la unidad compartida\", \"se actualizo el plugin y no anda\", \"hay un plugin nuevo\", \"como instalo esto en otra compu\", o cuando el aviso de inicio [setup-bva] diga que faltan dependencias o plugins."
---

# Setup del Estudio BVA

Deja la computadora lista para correr los plugins del marketplace
`estudiobva-skills` tal cual se desarrollaron. Hay dos partes:

- **El instalador de Windows** (`instalar.ps1`, doble clic en `Instalar Claude BVA.cmd`):
  lo corre la persona una vez por computadora. Instala Git, Node, Python, GitHub CLI y
  Claude Code, hace el login de GitHub, agrega el marketplace, instala todos los
  plugins y pide la contrasena de la planilla de claves en su propia ventana.
  Esta en la unidad compartida: `0 - ESTUDIO - 00000000000/automatizacion_bva/setup/`.
- **Esta skill**: para todo lo demas, desde Claude. Nunca pide contrasenas.

## Diagnostico (siempre primero)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/diagnostico.js
```

No entra a ARCA ni a bancos y no muestra claves. Cada linea es `OK`, `AVISO` o
`ERROR` con el arreglo al lado. Contarle al usuario solo lo que no esta OK.

## Arreglos

| Lo que dice el diagnostico | Que hacer |
|---|---|
| `plugin:<x>` falta `npm:...` / `python:...`, o `chromium` no abre | `cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/dependencias.js` (unos minutos; avisar). |
| `plugins` sin instalar: `<x>` | Ofrecer instalarlos: `claude plugin install <x>@estudiobva-skills` por cada uno, despues `node scripts/dependencias.js`, y pedir que reinicie la app de Claude. |
| `env`, `unidad` o `claves-organismos` en ERROR | `cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/configurar.js` (detecta la unidad y el Python y completa el `.env`). Si sigue en error por la **contrasena**, NO pedirla en el chat: que la persona corra el instalador (doble clic en `Instalar Claude BVA.cmd`), que la pide oculta. |
| `github` sin acceso | La persona tiene que estar invitada a la organizacion `estudiobva` en GitHub (se lo pide a Joaquin) y loguearse: el instalador lo hace. |
| `python` no arranca | Correr el instalador (instala Python 3.12). |
| `auto-update` apagado | `cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/configurar.js`. |
| `claves-cifrado` | Solo informar: la planilla de claves de organismos no tiene contrasena de apertura. |

Despues de arreglar, volver a correr el diagnostico y confirmar que quedo en orden.

## Actualizar todo a mano

Con la actualizacion automatica activada Claude baja las versiones nuevas al abrir.
Para forzarlo:

```bash
claude plugin marketplace update estudiobva-skills
claude plugin update <plugin>@estudiobva-skills      # uno por plugin instalado
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/dependencias.js
```

Los plugins nuevos se ven recien despues de reiniciar la app de Claude.

## Scripts

| Script | Que hace |
|---|---|
| `scripts/diagnostico.js [--json]` | Revisa todo. Exit 1 si hay algun ERROR. |
| `scripts/dependencias.js [--revisar] [--forzar] [--plugin=<x>] [--json]` | npm install, Chromium y `pip install --user -r requirements.txt` de cada plugin instalado, solo lo que falte. |
| `scripts/configurar.js [--revisar] [--unidad=<ruta>] [--python=<exe>] [--json]` | Completa `~/.fisco-ar/.env` (`BVA_UNIDAD_PATH`, `CLAVES_ORGANISMOS_PATH`, `BVA_PYTHON`) con copia `.env.bak-*` antes de tocarlo, y activa la actualizacion automatica del marketplace. Las contrasenas solo entran por `BVA_SETUP_CLAVE` / `BVA_SETUP_CLAVE_BANCOS` desde el instalador. |
| `scripts/chequeo-inicio.js` | Hook de inicio: si faltan dependencias, el `.env` o plugins, deja el aviso `[setup-bva]`. |

## Reglas

- Nunca pedir, leer ni mostrar contrasenas o claves fiscales; nunca imprimir el `.env`.
- No borrar ni reescribir el `.env`: `configurar.js` solo agrega o reemplaza sus claves y deja copia.
- En Windows el Python es `python` (o el de `BVA_PYTHON`), no `python3`: `python3` puede abrir la Microsoft Store.
