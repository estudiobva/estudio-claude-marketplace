---
name: papel-trabajo-automatico
description: "Arma el papel de trabajo unificado de liquidaciones (AAAA00-liquidaciones-IVA, IIBB y Tasa Municipal-<SOC>.xlsx) de los meses ya presentados de una sociedad o de todas, desde las fuentes oficiales: libros del Portal IVA de la unidad, Maestro de Proveedores, DDJJ de IVA F.2051 y DJ de anticipo de ARBA. Lo guarda aparte como '... - AUTOMATICO.xlsx' al lado del papel del equipo, sin tocarlo, y marca en amarillo con comentario lo que no tiene fuente. Usar SIEMPRE que pidan \"generar el papel de trabajo automatico de [sociedad]\", \"armar el papel de liquidaciones desde las fuentes\", \"rehacer el papel automatico hasta [mes]\", \"comparar el papel del equipo contra las DDJJ\" o \"papel AUTOMATICO\". No confundir con papeles-trabajo-unificados (Cowork, arma el papel desde los Excel mensuales del equipo) ni con cruze-zetti-portaliva-ventas / papel-trabajo-compras (hojas del papel mensual)."
---

# Papel de trabajo automático (liquidaciones)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python3 scripts/generar_papel.py --cuit=<CUIT> --hasta=202609
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python3 scripts/generar_papel.py --todas --hasta=202609 --dir=<carpeta>
```

(En Windows, `python` en vez de `python3`.)

| Argumento | Para que |
|---|---|
| `--cuit` | Una sociedad. |
| `--todas` | Las sociedades que tienen papel unificado propio. Las que no, de a una con `--cuit`. |
| `--hasta` | Ultimo mes a cargar, `AAAAMM` (default: el mes anterior). El año del papel es el de `--hasta`. |
| `--salida` | Ruta del archivo (solo con `--cuit`). Default: al lado del papel del equipo, `... - AUTOMATICO.xlsx`. |
| `--dir` | Con `--todas`: deja cada papel en esa carpeta en vez de al lado del original. |

## Antes de correr: que esten las fuentes

1. **Libros del Portal IVA** de cada mes en la carpeta IVA (`portal-iva-descarga`, con
   `--presentada` para meses presentados). Mes sin libros: queda como estaba en el molde.
2. **DDJJ de IVA**: `ddjj-iva` (deja el JSON en `~/Documents/BVA-salidas/ddjj-iva/`), o el
   PDF `AAAAMM-IVA-DDJJ-*` de la carpeta del mes. Sin DDJJ: retenciones y saldos en amarillo.
3. **DJ de ARBA**: `constancias-anticipo-arba`, o el PDF `AAAAMM-IIBB-DJ-*` de la carpeta del
   mes (si es un escaneo sin texto no sirve). Sin DJ (CM, AGIP o sin CIT): IIBB en amarillo.

## Que hace

- Molde: el papel unificado de la sociedad; si no tiene, el de otra del mismo municipio y
  jurisdiccion (Maestro de Sociedades), vaciado.
- Por mes presentado: ventas y compras por tipo de comprobante desde los libros, compras
  clasificadas con el Maestro de Proveedores (proveedor sin Maestro = servicio, se avisa),
  retenciones de IVA y saldos anteriores desde la DDJJ (manda lo presentado; si difiere del
  encadenado queda un comentario), IIBB desde la DJ de ARBA con el ajuste de redondeo contra
  el saldo de la DJ, base de Seguridad e Higiene = ventas gravadas, concepto fijo de
  Avellaneda = el valor mas repetido del mes en los papeles del equipo.
- Regenera todas las formulas y reapunta la solapa EERR del papel. Quita vinculos externos
  sin uso que traen algunos papeles (Excel en Mac no los abre).
- El archivo se abre con Excel para que calcule: openpyxl no deja valores calculados.

## Respuesta al usuario

Donde quedo el papel y, por mes, las notas del informe (ok, retenciones pendientes, IIBB
pendiente, ajuste de IIBB > $1, proveedores sin Maestro, concepto fijo pendiente). Recordar
que el crédito fiscal sale de la clasificacion del Maestro y puede diferir del papel del
equipo (criterio en revision con el equipo).

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`python3 -m pip install -r requirements.txt`.
