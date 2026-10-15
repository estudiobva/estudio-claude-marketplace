# clasificacion-compras

Clasifica las compras del mes con el Maestro de Proveedores del estudio: a cada
comprobante del Libro Compras del Portal IVA le pone la cuenta contable de su proveedor y
el rubro del EERR, y lista los proveedores que faltan en el Maestro o estan sin confirmar.
Usa el libro ya depurado de NCR si esta la conciliacion de `conciliacion-ncr`.
No entra a ARCA ni modifica el Maestro.

## Instalación

En las máquinas del equipo lo instala `setup-bva`. A mano:

```
/plugin install clasificacion-compras --marketplace estudiobva/estudio-claude-marketplace
```

```bash
python3 -m pip install -r requirements.txt
cp .env.example .env   # opcional: ruta de la unidad compartida y del Maestro
```

## Uso

```bash
python3 scripts/clasificar-compras.py --nombre="<sociedad>" --periodo=MM/AAAA --solo-revisar
python3 scripts/clasificar-compras.py --nombre="<sociedad>" --periodo=MM/AAAA
python3 scripts/clasificar-compras.py --nombres="A;B" --desde=01/2026 --hasta=08/2026
python3 scripts/clasificar-compras.py --todas --periodo=MM/AAAA
```

Deja `AAAAMM - COMPRAS CLASIFICADAS.xlsx` en la carpeta IVA del mes: detalle completo,
resumen por rubro, total por proveedor, pendientes del Maestro y control.

## Reglas

- El Maestro manda: Propuesta usa la cuenta propuesta; Revisar va a GASTOS VARIOS;
  Bloqueado y COMPRA DE EQUIPOS quedan excluidos del resultado; un CUIT que no está en el
  Maestro va a GASTOS VARIOS como SIN MAESTRO y se informa para darlo de alta.
- Importe imputado: neto (no gravado + exento + netos gravados) o, sin IVA discriminado,
  total menos percepciones. Las notas de crédito vienen en negativo y restan solas.
- No pisa archivos: lo anterior va a `_anteriores/` con `--si-existe=reemplazar`.

## Estructura

- `scripts/clasificar-compras.py`: la clasificación (solo openpyxl).
- `skills/clasificacion-compras/SKILL.md`: la habilidad.
