---
name: portal-iva-descarga
description: "Baja del Portal IVA de ARCA el Libro IVA Compras y el Libro IVA Ventas de una o varias sociedades para un periodo (por defecto el mes anterior), importando primero desde ARCA al borrador, los convierte a Excel y los archiva en la unidad compartida como 'AAAAMM - PORTAL IVA - COMPRAS.xlsx' y 'AAAAMM - PORTAL IVA - VENTAS.xlsx' en la carpeta IVA del mes, sin dejar el CSV. Corre un script de Playwright que se loguea solo con la planilla de claves del estudio. Usar SIEMPRE que pidan \"bajar el Portal IVA\", \"descargar compras y ventas del Portal IVA\", \"bajar el libro IVA de [sociedad] [mes]\", \"traer las compras/ventas de ARCA\" o \"preparar el IVA del mes\". Es el paso 1 del circuito conciliacion-ncr: conciliar-ncr.py lee este Excel de compras igual que el CSV. Con --presentada baja los libros de una DDJJ YA PRESENTADA (ultima secuencia, solo lectura, controlados contra la vista previa): usarlo cuando pidan \"los libros de la DDJJ presentada\", \"el libro IVA presentado de [mes]\" o un periodo ya presentado. No confundir con portal-iva-compras (carga CSV ya bajados en el template input_portal) ni con arca-portal-iva de fisco-ar (solo totales)."
---

# Descarga del Portal IVA — Compras y Ventas

Un script de Playwright entra a ARCA con la clave de la planilla del estudio,
cambia la representada, abre el borrador del periodo, y por cada libro:
**Importar desde ARCA** (espera a que la tarea figure Procesada), baja el CSV, lo
convierte a Excel y lo archiva. El CSV se borra: queda solo el `.xlsx`. No hace
falta loguearse a mano.

Se trabaja a **mes vencido**: se corre los primeros dias del mes siguiente, antes
de presentar. Sin `--periodo` el script toma el mes anterior al de hoy.

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/portal-iva-descarga.js --nombre="<sociedad>"                # mes anterior
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/portal-iva-descarga.js --nombre="<sociedad>" --periodo=08/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/portal-iva-descarga.js --nombres="<A>;<B>;<C>" --periodo=08/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/portal-iva-descarga.js --nombre="<sociedad>" --periodo=08/2026 --presentada   # DDJJ ya presentada
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit` | Una sociedad (titular de la planilla de claves, o CUIT). |
| `--nombres` / `--cuits` | Varias, separadas por `;`. Un login por sociedad, una despues de otra. |
| `--periodo` | `MM/AAAA` o `AAAAMM`. Default: el mes anterior. |
| `--libros` | `compras`, `ventas` o `compras,ventas` (default los dos). |
| `--si-existe` | Si ya esta el `.xlsx`: `abortar` (default), `saltear` (ni entra a ARCA si estan los dos) o `reemplazar`. |
| `--presentada` | Baja los libros de la DDJJ **ya presentada** (ver abajo). No importa ni toca el borrador. |
| `--aunque-presentada` | Entra **al borrador** aunque haya acuse de la DDJJ IVA del periodo (ver Reglas). Para bajar lo presentado, usar `--presentada`. |
| `--no-archivar` | Deja los Excel en `~/Documents/BVA-salidas/portal-iva/` y no toca el Drive. Igual importa al borrador. |
| `--ver` | Navegador visible. Para lotes conviene headless. |

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
Las credenciales van en `~/.fisco-ar/.env` (copiar `.env.example`), que sobrevive a
las actualizaciones. En Mac, `BVA_UNIDAD_PATH` tiene que apuntar a la unidad
compartida montada por Google Drive.

## DDJJ ya presentadas (`--presentada`)

Para bajar los libros tal como quedaron presentados (por ejemplo, para cruzar con
el padron o armar el estado de resultados de meses cerrados). El script entra
solo, igual que en el modo normal, y va por:

**Declaraciones juradas presentadas → Libro IVA → "Ver"** de la ultima secuencia
del periodo (si hay rectificativas, la mas reciente) → **Libro Compras / Ventas**,
que se abren sin el boton IMPORTAR → CSV.

- **No toca el borrador ni importa nada.** Nunca clickea "Rectificar" (esta en la
  lista de botones prohibidos del modulo).
- **Control antes de archivar:** el neto gravado del CSV (operaciones y notas de
  credito por separado) tiene que coincidir al centavo con la vista previa de la
  DDJJ. Si no coincide, no archiva. El IVA se informa (`control_ddjj`) pero no se
  exige: ARCA redondea por alicuota y difiere en unos pocos pesos.
- No mira el acuse de la carpeta: si ARCA no tiene presentacion del periodo, corta
  con `sin_presentacion`.
- Mismo nombre y carpeta que el modo normal (`AAAAMM - PORTAL IVA - COMPRAS.xlsx`).
  Si ya existe el del borrador, pasar `--si-existe=reemplazar` para quedarse con el
  presentado.
- El JSON trae `ddjj`: formulario, secuencia, fecha de presentacion y cuantas
  secuencias tiene el periodo.

## Que entrega

En la carpeta IVA del periodo de cada sociedad:

- `AAAAMM - PORTAL IVA - COMPRAS.xlsx` — hoja "Compras Portal IVA".
- `AAAAMM - PORTAL IVA - VENTAS.xlsx` — hoja "Ventas Portal IVA" (el nombre que
  espera `cruce-ventas-zetti-portal-iva`).

Mismas columnas y orden que el CSV de ARCA, con tipos reales (fechas, enteros,
importes), encabezado fijo, autofiltro, una fila en blanco y la fila **TOTAL** con
formulas `=SUM()`. Las notas de credito vienen en negativo: el total ya es neto.

La carpeta IVA se resuelve igual que en conciliacion-ncr: `AAAAMM/AAAAMM-IVA` o
`AAAAMM-IVA` si ya existen, si no `AAAAMM/IVA` (se crea). Si la sociedad usa carpeta
de año (`01-Impuestos Mensuales/2026/202608`), el mes va adentro del año.

Por stdout sale un JSON con, por libro, los registros importados, las filas
bajadas, las filas con fecha fuera del periodo y la ruta del archivo. Reportarle al
usuario sociedad, periodo, comprobantes de cada libro y donde quedo.

## Reglas

- **Nunca presenta** ni toca Presentar / Confirmar / Generar DJ / Eliminar todos /
  Descartar: el modulo `lib/portal-iva.js` aborta si un click cae en esa lista.
- **DDJJ ya presentada:** si en la carpeta del mes hay un acuse `AAAAMM-IVA-ACUSE-*`,
  el modo normal no entra al borrador (con el periodo presentado, "Nueva declaracion
  jurada" puede abrir una rectificativa). Para bajar lo presentado: `--presentada`.
  `--aunque-presentada` solo si el usuario pide expresamente entrar al borrador.
- **Un solo borrador por sociedad:** si ARCA tiene otro periodo sin presentar, corta
  con `otro_borrador_abierto`. Descartarlo borra lo cargado: nunca sin autorizacion
  expresa del usuario.
- **No pisa** un Excel existente salvo `--si-existe=reemplazar`. Si al lado quedo un
  `.csv` viejo con el mismo nombre, avisa (conciliar-ncr encontraria dos archivos de
  compras y cortaria) pero no lo borra.
- Si ARCA pide captcha o rechaza la clave, corta el lote: reintentar bloquea la clave.
- Circuito NCR: una importacion desde ARCA posterior a eliminar las NCR las vuelve a
  meter en el borrador. Si ya se depuro el mes, no volver a bajarlo.
