# armar-balance

Libro de Balance (EECC) prototipo por sociedad: Situación Patrimonial, Estado de Resultados, Evolución del PN, Flujo de Efectivo, Nota 2 y Anexos 1, 2 y 3, con fórmulas y el comparativo del ejercicio anterior. No repite las liquidaciones mensuales: toma los totales del EERR y del papel de trabajo y suma lo que el papel no tiene (FCI, planes, F.931, anticipos, saldos patrimoniales con su fecha de corte).

| Pieza | Qué hace |
|---|---|
| `scripts/generar_lote.py` | Genera los Balances de las sociedades de la carpeta de configs (`EERR 2026/Config Balance` en la unidad). |
| `scripts/armar_balance.py` | Un Balance: `armar_balance.py <EERR.xlsx> <salida.xlsx> <config.json>`. |
| `scripts/fuentes.py` | Lectores de las fuentes de la unidad (papel mensual, T&S, DRIVE - FCI, planes, F.931, anticipos, DDJJ Ganancias). |
| `scripts/pub-descarga.js` | Baja de PUB (ARCA) el PDF del balance presentado de un período. Solo lectura. |
| `scripts/herramientas/` | `validar.py` y `calcular_eerr.py` (pycel, sin Excel), `chk_config.py` (control del comparativo transcripto), `zoom.py` (leer PDF escaneados). |

## Instalación

```
/plugin install armar-balance --marketplace estudiobva/estudio-claude-marketplace
```

En la carpeta del plugin:

```
pip install -r requirements.txt
npm install
npx playwright install chromium
```

Variables en `~/.fisco-ar/.env` (las mismas de los otros plugins): `BVA_UNIDAD_PATH` si la unidad no está en `I:\Unidades compartidas\…` ni en Google Drive de Mac; `BALANCE_CONFIG_DIR` para otra carpeta de configs; `PUB_LOGIN_ALT` (CUIT de una sociedad cuyo login representa a muchas sociedades en PUB) para el segundo intento de `pub-descarga`. Credenciales de ARCA: planilla de claves, como el resto.

Los configs **no** van en este repo: tienen datos de los balances de clientes.

Uso detallado: `skills/armar-balance/SKILL.md`. Transcripción del comparativo: `skills/armar-balance/transcribir-comparativo.md`.
