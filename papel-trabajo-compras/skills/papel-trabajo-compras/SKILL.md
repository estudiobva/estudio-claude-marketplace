---
name: papel-trabajo-compras
description: "Actualiza el papel de trabajo del mes de una o varias sociedades con las compras: lee el Libro Compras ACTUAL del Portal IVA (borrador sin importar, o el de la DDJJ presentada), baja el Subdiario de IVA Compras de Zetti (5.6.4), controla que esten los servicios (AYSA, EDESUR, METROGAS) y que no quede ninguna NCR, y agrega al MISMO papel las hojas Compras Zetti, Compras Portal IVA (clasificada, con el cuadro de Compras y el Prorrateo IVA Compras de Servicios al final), Cruce compras, Resumen cruce compras y Clasificacion compras. Login automatico y navegador oculto. Usar SIEMPRE que pidan \"actualizar el papel de trabajo de [sociedad] [mes] con las compras\", \"compras de Zetti y Portal IVA en el papel\", \"cruce y clasificacion por cliente/proveedor\", \"el cuadro de compras al final del portal\", \"fijate que esten los servicios\" o \"rechequea que no esten las NCR\". Va despues de cruze-zetti-portaliva-ventas (papel con liquidacion) y de conciliacion-ncr (NCR eliminadas)."
---

# Compras en el papel de trabajo (circuito completo)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/papel-trabajo-compras.js --nombre="<sociedad>" --periodo=09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/papel-trabajo-compras.js --cuit=<CUIT> --periodo=09/2026 --si-existe=reemplazar --no-es-ncr=503-86152
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/papel-trabajo-compras.js --nombres="<A>;<B>" --periodo=09/2026
```

Tarda unos minutos por sociedad (ARCA + Zetti + Excel). Correrlo en segundo plano.

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit`, `--nombres` / `--cuits` | Sociedad(es), como en los otros scripts. |
| `--periodo` | `MM/AAAA` o `AAAAMM`. Default: el mes anterior. |
| `--si-existe` | `abortar` (default) o `reemplazar`: si el papel ya tiene las hojas de compras, las rehace (se pierde lo escrito a mano en esas 5 hojas; el resto del papel no se toca). |
| `--no-es-ncr=PV-NRO;...` | NC de Suizo/Monroe que el estudio confirmo que NO son NCR (ej. `503-86152`). |
| `--login-nombre=<titular>` | Si la clave de la sociedad no tiene la representada en el Portal IVA (paso con Sanar 3, Mega Sol y Tofanelli: entrar con un titular de la planilla que las represente). Tambien `BVA_ARCA_LOGIN_NOMBRE` en el .env. |
| `--papel=<ruta>` | Papel a usar si en la carpeta del mes hay mas de uno. |
| `--zetti-unidad` | Usa el `AAAAMM - ZETTI - COMPRAS.xlsx` de la carpeta IVA en vez de bajarlo de nuevo. |
| `--no-archivar` | Arma y controla, pero no toca la unidad: el resultado queda en la carpeta de trabajo. |

## Reglas del estudio que el script respeta

- **No duplicar el papel**: se pisa el mismo `AAAAMM-papeldetrabajo-*.xlsx`; nada de copias ni
  backups en la unidad. El respaldo queda en `~/Documents/BVA-salidas/papel-compras/<cuit>-<periodo>-<hora>/papel-original.xlsx`.
- **Fuente del Portal = el Libro Compras actual de ARCA, solo lectura y SIN IMPORTAR** (importar
  traeria de vuelta las NCR). Si la DDJJ ya esta presentada usa el libro presentado (ultima secuencia).
- **Servicios**: AYSA, EDESUR y METROGAS de Zetti tienen que estar en el Portal. Si falta alguno se agrega
  a mano en `Compras Portal IVA` (filas en naranja) y se avisa que hay que cargarlo en ARCA. **No se carga en ARCA.**
  Tambien avisa si un servicio del mes anterior no aparece este mes.
- **NCR**: corta (`code=ncr_en_borrador`) si alguna NCR conciliada sigue en el libro, o si alguna NC de
  Suizo/Monroe del libro figura en cualquier listado de NCR de la unidad (todas las farmacias y meses).
  Una NC grande de drogueria solo en el Portal (> $1 M) sale como "REVISAR posible NCR"; no se borra. Si el
  estudio confirma que no es NCR, volver a correr con `--si-existe=reemplazar --no-es-ncr=PV-NRO`.
- "Clasificacion por cliente" = por proveedor (hoja Clasificacion compras, con el Maestro de Proveedores).
- El cuadro de compras + prorrateo va al final de `Compras Portal IVA` (formato Soy 59 / Sanar 3 / Soy 51 Solano):
  Factura A por alicuota y Mercaderia/Servicios, NC, totales, IVA computable de servicios x % de Venta Gravada de
  la hoja `liquidacion`, CF de mercaderia completo, servicios publicos y control contra el libro.

## Que hace, por sociedad

1. Ubica el papel del mes (tiene que existir y tener `liquidacion` con "Venta Gravada": lo arma
   cruze-zetti-portaliva-ventas) y la conciliacion NCR (`ncr/AAAAMM-COMPRAS-SIN-NCR-*.xlsx`; sin eso corta con
   `code=sin_conciliacion`: correr antes conciliacion-ncr y la eliminacion).
2. `leer-borrador-compras.js`: exporta el Libro Compras del borrador actual (o de la DDJJ presentada).
3. `zetti-iva-compras.js --no-archivar`: baja Zetti 5.6.4 y lo compara con el de la unidad (si la unidad no lo
   tiene, lo archiva en la carpeta IVA).
4. `armar_papel_compras.py`: arma las 5 hojas antes de `liquidacion` (las demas no se tocan).
5. Excel recalcula (oculto) y `verificar_papel.py` controla: cruce y clasificacion en 0, cuadro vs libro < $1,
   sin errores de Excel y hojas previas iguales. Si algo no da, NO toca la unidad (`code=verificacion`).
6. Controla que nadie haya cambiado ni tenga abierto el papel (`~$`) y lo pisa.

## Respuesta al usuario

Por sociedad, desde el JSON: papel actualizado; fuente del libro (borrador o DDJJ presentada); comprobantes y
total de Portal y Zetti (y si el Zetti de hoy es igual al de la unidad); NCR eliminadas y ninguna en el libro;
servicios (en el Portal / agregados a mano / faltantes); cruce (coinciden, con diferencia, solo en Portal, solo en
Zetti); controles; los `hallazgos` (van tambien en la hoja Resumen cruce compras) y los proveedores pendientes del
Maestro. Si hay "REVISAR posible NCR", preguntar al usuario si es NCR.

Errores (`code`): `sin_papel`, `papel_ambiguo`, `sin_conciliacion`, `ncr_en_borrador`, `hojas_existen`,
`papel_sin_liquidacion`, `tipo_desconocido`, `sin_representacion`, `borrador_no_iniciado`, `zetti_formato`,
`verificacion`, `papel_cambio`, `papel_abierto`, `excel`.

**Requisitos:** Windows con Excel (recalculo por COM). Plugin instalado: `setup-bva` instala las dependencias
(`npm install`, Chromium de Playwright y `pip install -r requirements.txt`).
