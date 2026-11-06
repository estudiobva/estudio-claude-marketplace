"""Proceso TSH / IIBB, paso 3b: fila TOTAL al pie de la solapa "Ventas MCR".

    python totales_mcr.py <papel.xlsx>

Igual que el TOTAL del Portal IVA: una fila en blanco, "TOTAL" en A (negrita) con
=SUM de cada columna de importes (de "Imp. Neto Gravado IVA 0%" a "Imp. Total"),
y debajo una nota con la cantidad de comprobantes. Tipo Cambio y Moneda no se suman.
Mis Comprobantes trae las NC en positivo: el total suma facturas y NC (no es neto),
y la nota lo aclara. Si ya hay una fila TOTAL, la reemplaza.
"""
import json
import sys

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import fuente

HOJA = 'Ventas MCR'
FMT = '#,##0.00;-#,##0.00;-'


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    papel = sys.argv[1]
    wb = load_workbook(papel)
    ws = wb[HOJA]

    tit = {str(c.value).strip(): c.column for c in ws[1] if c.value}
    desde, hasta = tit['Imp. Neto Gravado IVA 0%'], tit['Imp. Total']

    # ws.max_row recorre todas las celdas cada vez que se consulta: se toma una sola vez
    # (dentro del while hacia el calculo cuadratico; Salud y Belleza, 77.000 filas, tardaba horas).
    maxr = ws.max_row
    ult = 1
    while ult < maxr and ws.cell(ult + 1, 1).value is not None and not isinstance(ws.cell(ult + 1, 1).value, str):
        ult += 1
    if maxr > ult:
        ws.delete_rows(ult + 1, maxr - ult)

    ft, fn = ult + 2, ult + 3
    ws.cell(ft, 1, 'TOTAL').font = Font(bold=True)
    for c in range(desde, hasta + 1):
        L = get_column_letter(c)
        cel = ws.cell(ft, c, f'=SUM({L}2:{L}{ult})')
        cel.number_format, cel.font = FMT, Font(bold=True)
    n = ult - 1
    ws.cell(fn, 1, f'{n} comprobantes · Fuente: Mis Comprobantes de ARCA, emitidos · Las notas de credito '
                   f'vienen con importe positivo, o sea que el total suma facturas y NC (no es neto).')
    fuente.aplicar(ws, fuente.ARCA, [c for fila in ws.iter_rows(min_row=ft, max_row=fn) for c in fila if c.value is not None])
    wb.save(papel)
    print(json.dumps({'ok': True, 'comprobantes': n, 'fila_total': ft,
                      'columnas': f'{get_column_letter(desde)}:{get_column_letter(hasta)}'}))


if __name__ == '__main__':
    main()
