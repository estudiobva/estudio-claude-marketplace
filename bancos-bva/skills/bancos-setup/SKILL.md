---
name: bancos-setup
description: Prepara y verifica el entorno del plugin bancos-bva — instala Playwright y Chromium, apunta el .env a Claves_Bancos.xlsx y corre un diagnóstico que confirma que los scripts de home banking pueden ejecutarse (sin entrar a ningún banco). Usar cuando digan "configurar bancos-bva", "instalar el plugin de bancos", "no me anda el script de extractos", "configurar Claves_Bancos", "verificar que funciona la descarga de extractos", o antes de la primera descarga de extractos en una PC.
---

# Setup de bancos-bva

Deja la PC lista para `descarga-extractos-bancarios` y **comprueba que anda**
antes de la primera corrida real. Nada de esto entra a un banco.

## 1. Instalar dependencias

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && npm install
```

El `postinstall` baja Chromium (unos 150 MB la primera vez). Si no lo hizo:
`npx playwright install chromium`.

## 2. El `.env`

Es **el mismo que usa fisco-ar**: `~/.fisco-ar/.env` (orden de búsqueda:
`$FISCO_ENV_FILE` → `<raíz del plugin>/.env` → `~/.fisco-ar/.env`). Si fisco-ar
ya está configurado, normalmente no hay que tocar nada: la contraseña de
`Claves_Bancos.xlsx` hoy es la misma que la de `Claves_Organismos.xlsx`
(`CLAVES_ORGANISMOS_PASSWORD`).

Solo si difieren, agregar `CLAVES_BANCOS_PASSWORD`. La contraseña la escribe la
persona en el archivo; nunca repetirla en la respuesta ni hacer `cat` del `.env`.

La unidad compartida se detecta sola (`I:\` en Windows, Google Drive para
escritorio en Mac). Solo si está en otro lado: `BVA_UNIDAD_PATH`.

## 3. Diagnóstico

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/doctor.js
```

Chequea Node, dependencias, que **Chromium abra**, qué `.env` cargó, la unidad,
el Maestro de Sociedades y que `Claves_Bancos.xlsx` **abra con la contraseña**
(informa qué bancos tienen usuario y clave cargados — nunca los valores).

- `✅` en todo → listo.
- `⚠️ claves-bancos` → algún banco sin datos en el archivo: esas descargas van a fallar hasta completarlo.
- `❌` → resolver lo que dice `→` antes de seguir.

## 4. Datos que conviene confirmar una vez

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/claves-bancos.js --banco=FRANCES
```

BBVA pide código de empresa, código de usuario y clave; el adaptador asume
empresa = ACCESO/DNI y usuario = USUARIO. El comando muestra el **largo** de
cada valor enmascarado (`***(8)`), suficiente para confirmar el mapeo. Nunca
usar `--revelar`.

## Problemas frecuentes

- **`Executable doesn't exist`** → `npx playwright install chromium`.
- **`No pude descifrar Claves_Bancos.xlsx`** → contraseña mal en el `.env` (ojo espacios).
- **`No encuentro la unidad compartida`** → Drive no montado o `BVA_UNIDAD_PATH` mal.
- **Un banco cambió la pantalla de login** → `node scripts/banco-login.js --banco=X --quedate` lo abre visible para ver dónde se traba; el arreglo va en `lib/bancos/<banco>.js`.
