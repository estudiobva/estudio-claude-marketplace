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
| Ventas | libro de ventas del Portal IVA | portal-iva-descarga (`--presentada`) |
| CMV y gastos | libro de compras del Portal IVA + Maestro de Proveedores | portal-iva-descarga / conciliacion-ncr |
| IIBB | DJ de anticipo de ARBA | constancias-anticipo-arba |
| Sueldos | F.931 | f931-descarga-arca |
| IVA prorrateo | libros del Portal IVA | — |
| Seg. e Higiene, Publicidad | papel de trabajo de liquidaciones | papel del equipo / papel-trabajo-automatico |

Reglas:

- Solo completa celdas vacías: lo que cargó o corrigió el equipo no se toca (salvo `--pisar`).
- Cada carga queda registrada con su archivo de origen en la hoja **Armado** del EERR.
- Respaldo del EERR anterior en `EERR 2026/_respaldos/`; si el archivo cambió mientras corría, no se pisa.
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
