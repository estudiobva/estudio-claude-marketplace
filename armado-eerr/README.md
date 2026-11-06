# armado-eerr

Arma y actualiza los Estados de Resultados (EERR) mensuales de las sociedades de la unidad compartida
(`0 - ESTUDIO - 00000000000/EERR 2026`). Es el armado con el que se hicieron los EERR 2026 (octubre 2026), guardado
como herramienta para que cualquiera del equipo lo pueda repetir mes a mes.

## Uso

```bash
python scripts/armar_eerr.py --nombre="<sociedad>" --simular      # ver qué cargaría, sin tocar la unidad
python scripts/armar_eerr.py --nombre="<sociedad>"                # completar los últimos meses
python scripts/armar_eerr.py --todas                              # todos los EERR
python scripts/armar_eerr.py --nombre="<sociedad>" --crear        # EERR del ejercicio nuevo
python scripts/armar_eerr.py --todas --crear                      # las sociedades cuyo ejercicio cerró
```

Opciones: `--hasta=AAAAMM`, `--desde=AAAAMM`, `--pisar=AAAAMM,...`, `--cuit=...`. Detalle en
`skills/armado-eerr/SKILL.md`.

## Fuentes

| Concepto | De dónde sale | Herramienta que la deja en la unidad |
|---|---|---|
| Ventas, IIBB, IVA prorrateo, Seg. e Higiene, PyP | papel de liquidaciones del equipo (hoja EERR) | el equipo |
| … si el papel no tiene el mes | libro del Portal IVA / DJ de ARBA / cálculo | portal-iva-descarga, constancias-anticipo-arba |
| CMV y gastos | libro de compras del Portal IVA + Maestro de Proveedores | portal-iva-descarga / conciliacion-ncr |
| Sueldos | F.931 | f931-descarga-arca |

El control contra las DDJJ lo hace `auditor-eerr`.

Reglas:

- Solo completa celdas vacías: lo que cargó o corrigió el equipo no se toca (salvo `--pisar`).
- Cada carga queda registrada con su archivo de origen en la hoja **Armado** del EERR.
- Respaldo del EERR anterior en `EERR 2026/_respaldos/`; si el archivo cambió mientras corría, no se pisa.
- Desde 1.0.3: las NCR que eliminó la conciliación (hoja Eliminados del `AAAAMM-COMPRAS-SIN-NCR`) no entran a las compras
  aunque el libro elegido todavía las traiga (antes dejaban el CMV subvaluado: Sanar 1 y Soy Nueva Ventura en 09/2026).
- Necesita Microsoft Excel para recalcular (Windows: PowerShell; Mac: AppleScript).

## Instalación

```bash
python -m pip install -r requirements.txt
```

Variables (en `~/.fisco-ar/.env`, como el resto de los plugins): `BVA_UNIDAD_PATH`, `BVA_SALIDAS_PATH`.

## Plantilla

`assets/Plantilla EERR.xlsx`: hojas EERR y Sueldos vacías (sin datos de ninguna sociedad), con las filas de intereses
por facilidades de pago e indemnización por despido. Las hojas "Compras imputadas", "Ventas Portal IVA",
"IVA prorrateo" y "Armado" las crea el script.
