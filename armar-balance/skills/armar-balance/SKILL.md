---
name: armar-balance
description: "Arma el libro de Balance (EECC) prototipo de una o varias sociedades en '0 - ESTUDIO - 00000000000/EERR 2026': Situación Patrimonial, Estado de Resultados, Evolución del Patrimonio Neto, Flujo de Efectivo, Nota 2 y Anexos 1, 2 y 3, con fórmulas y el comparativo del ejercicio anterior, a partir del EERR (EERR 2026/EECC anterior) y de las fuentes acordadas (papel de trabajo mensual, T&S Gastos Varios, DRIVE - FCI, F.931, planes de facilidades, anticipos de Ganancias). También baja de PUB de ARCA el balance presentado del ejercicio anterior (pub-descarga) y guía la transcripción del comparativo al config de la sociedad. Usar SIEMPRE que pidan \"armar el balance de [sociedad]\", \"balances de las que cierran el 30/11\", \"cargar el comparativo\", \"transcribir el balance legalizado/anterior\", \"bajar el balance de PUB\", \"actualizar los balances con [mes]\" o \"regenerar los balances\". No confundir con armado-eerr (arma el EERR mensual) ni con auditor-eerr."
---

# Armar balance

Tres piezas, en este orden:

1. **Config por sociedad** (`EERR 2026/Config Balance/<sociedad>.json` en la unidad): nombre, CUIT, cierre y el **comparativo** del balance anterior. Los configs viven en la unidad, no en el plugin (tienen datos de clientes y el repo es público). Otra carpeta: `BALANCE_CONFIG_DIR` en `~/.fisco-ar/.env` o `--config-dir=`.
2. **Generador** (`scripts/generar_lote.py` → `scripts/armar_balance.py`): lee el EERR del ejercicio y las fuentes y escribe `EERR 2026/BALANCE DD-MM-AAAA - <sociedad> - PROTOTIPO.xlsx`.
3. **Comparativo**: si el config no lo tiene, se baja el balance anterior de PUB (`scripts/pub-descarga.js`) y se transcribe (ver `transcribir-comparativo.md`).

En Mac usar `python3` (o el `BVA_PYTHON` del `.env`). Antes de la primera corrida: `pip install -r requirements.txt` y `npm install` + `npx playwright install chromium` en la carpeta del plugin.

## Generar los Balances

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/generar_lote.py <config> [<config> ...] --destino=<carpeta de prueba>
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/generar_lote.py <config>
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/generar_lote.py
```

| Argumento | Para qué |
|---|---|
| `<config>` | Nombre del archivo de config sin `.json` (ej. `sanar-1`). Sin nombres: todas las de la carpeta de configs. |
| `--destino=` | Carpeta de salida. Default: `EERR 2026` (pisa el Balance de la sociedad). **Primero generar en una carpeta de prueba** y validar; copiar a la unidad solo si el usuario lo pide. |
| `--eerr-dir=` | De dónde lee los EERR. Default: `EERR 2026/EECC anterior`. Usar una carpeta con EERR calculados cuando Excel no recalcula (ver abajo). |
| `--config-dir=` | Carpeta de configs (default: `EERR 2026/Config Balance`). |

Claves del config: `cierre` (AAAA-MM-DD; default 2026-11-30), `ultimo_mes_con_datos` (AAAAMM; default el mes anterior al actual), `meses_cargados` (pisa el cálculo anterior), `razon_social`, `registro`, `ejercicio_nro`, `bancos`, `fci_2025`, `deudas_bancarias`, `ganancias_2025`, `mapa_anexo2`, `comparativo` (ver `transcribir-comparativo.md`).

El EERR se busca como `EERR MM-AAAA a MM-AAAA - <nombre>.xlsx` según el cierre. Para un ejercicio que cerró en el año, el EERR es el cerrado; para uno en curso, el vigente.

### Validar (sin Excel)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/herramientas/validar.py "<carpeta o archivo BALANCE>"
```

Muestra por Balance: ventas, gastos, resultado, `ESP25dif` (el comparativo cuadra: tiene que dar 0), `A2` (Anexo 2 = gastos del Resumen de Resultados: 0), pendientes/provisorios y errores de fórmulas (tienen que ser 0). Un resultado muy negativo es normal mientras falte el stock final: la existencia inicial ya entra al CMV.

### Si Excel no recalcula

Los Balances van con recálculo al abrir. Si los EERR no tienen los valores calculados (por ejemplo, una simulación de `armado-eerr` que no se guardó desde Excel): `python scripts/herramientas/calcular_eerr.py <carpeta con EERR> <carpeta destino>` y generar con `--eerr-dir=<carpeta destino>`.

## Comparativo del ejercicio anterior

1. **Buscar el PDF** en `04-Balances y EECC/<año>/EECC Legalizados` de la sociedad. Ojo: hay balances legalizados del año anterior guardados en la carpeta del año de legalización. Verificar siempre el cierre leyendo la carátula.
2. **Si no está, bajarlo de PUB** (solo lectura; nunca toca "Rectificar" ni "Nueva presentación"):
   ```bash
   cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/pub-descarga.js --cuit=<CUIT>[,<CUIT>...] --periodo=2025 [--constancia]
   ```
   Login con la planilla de claves. Si el apoderado de la planilla no representa a la sociedad en PUB, reintenta con el login de `PUB_LOGIN_ALT` (CUIT de una sociedad que entra con el apoderado que representa a más sociedades). Siempre baja el adjunto y lo guarda con el nombre que da ARCA; si ya hay un archivo idéntico no hace nada, y si hay otro con ese nombre guarda el de PUB como `_PUB.pdf`. Estados: `ok`, `sin_presentacion` (no hay balance de ese período en PUB), `sin_representacion` (ningún login la representa), `error`. **El adjunto de PUB puede ser el balance equivocado** (ya pasó: presentación 2025 con el balance 2024 adjunto): verificar el cierre y avisar al usuario.
3. **Transcribir** al config siguiendo `transcribir-comparativo.md` (en esta carpeta). Son PDF escaneados de ~20 páginas: conviene un subagente por sociedad, con ese instructivo, de a 5–9 en paralelo. El control `scripts/herramientas/chk_config.py <config>` tiene que dar 0 (±0,02) en todo.
4. **Regenerar** esas sociedades, validar y copiar a la unidad (con el OK del usuario si pisa archivos).

## Reglas

- Solo lee la unidad, salvo el Balance que escribe y los PDF que baja `pub-descarga`. El papel de trabajo y los EERR no se tocan.
- No pisar un Balance de la unidad sin OK si el usuario no pidió regenerarlo. Antes de pisar, guardar copia.
- Si en el PDF falta el número de registro o hay datos dudosos (FCI sin banco, préstamo sin acreedor, aportes irrevocables), anotarlos en el config y avisar.
