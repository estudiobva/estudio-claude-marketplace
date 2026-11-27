"""Proceso TSH / IIBB: copia la primera hoja de un Excel como solapa nueva del papel de trabajo.

    python copiar_solapa.py <origen.xlsx> <papel.xlsx> "<nombre de la solapa>" [--antes="<otra solapa>"]

Copia valores, formulas, formatos, anchos de columna, filas inmovilizadas y filtro,
tal cual el original. Si la solapa ya existe en el papel, la reemplaza en el mismo
lugar. Con --antes, si la solapa es nueva queda justo antes de esa otra (si existe).
Imprime un JSON con filas y columnas copiadas.
"""
import json
import sys
from copy import copy

from openpyxl import load_workbook


def main():
    opciones = [a for a in sys.argv[1:] if a.startswith('--')]
    posicionales = [a for a in sys.argv[1:] if not a.startswith('--')]
    if len(posicionales) != 3:
        sys.exit(__doc__)
    origen, papel, nombre = posicionales
    antes = next((a.split('=', 1)[1] for a in opciones if a.startswith('--antes=')), None)
    src = load_workbook(origen).worksheets[0]
    wb = load_workbook(papel)

    pos = len(wb.sheetnames)
    if antes and antes in wb.sheetnames:
        pos = wb.sheetnames.index(antes)
    if nombre in wb.sheetnames:
        pos = wb.sheetnames.index(nombre)
        del wb[nombre]
    ws = wb.create_sheet(nombre, pos)

    # Los formatos se resuelven una vez por combinacion: copiar font/border/etc.
    # celda por celda en ~2,5 millones de celdas tarda mas de 10 minutos.
    estilos = {}
    for fila in src.iter_rows():
        for c in fila:
            if c.value is None and not c.has_style:
                continue
            d = ws.cell(c.row, c.column, c.value)
            if c.has_style:
                k = tuple(c._style)
                if k not in estilos:
                    d.font = copy(c.font)
                    d.number_format = c.number_format
                    d.alignment = copy(c.alignment)
                    d.border = copy(c.border)
                    d.fill = copy(c.fill)
                    estilos[k] = copy(d._style)
                else:
                    d._style = copy(estilos[k])
    for k, dim in src.column_dimensions.items():
        ws.column_dimensions[k].width = dim.width
    for rango in src.merged_cells.ranges:
        ws.merge_cells(str(rango))
    ws.freeze_panes = src.freeze_panes
    if src.auto_filter.ref:
        ws.auto_filter.ref = src.auto_filter.ref

    wb.save(papel)
    print(json.dumps({'ok': True, 'solapa': nombre, 'filas': src.max_row, 'columnas': src.max_column,
                      'solapas': wb.sheetnames}, ensure_ascii=False))


if __name__ == '__main__':
    main()
