"""Etapa 2 (DDJJ de IVA), paso 6: titulo PROVEEDORES en la columna AG de "Compras portal iva".

    python titulo_proveedores.py --papel=<papel de trabajo.xlsx>

Escribe PROVEEDORES en AG1 con el estilo de AF1 (Arial negrita blanca sobre azul), ancho 22,
y extiende el filtro de la fila 1 hasta AG. La columna queda vacia debajo del titulo (la
completa clasificar_proveedores.py). Si ya esta, no cambia nada.
"""
import json
import re
import sys
from copy import copy

from openpyxl import load_workbook

HOJA = 'Compras portal iva'


def main():
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel = opts.get('papel')
    if not papel:
        sys.exit(__doc__)
    wb = load_workbook(papel)
    if HOJA not in wb.sheetnames:
        sys.exit(json.dumps({'error': f'Falta la solapa "{HOJA}"', 'solapas': wb.sheetnames}, ensure_ascii=False))
    ws = wb[HOJA]
    if str(ws['F1'].value or '').strip().lower() != 'nro. doc. vendedor':
        sys.exit(json.dumps({'error': 'La solapa no tiene la distribucion esperada (F1 = Nro. Doc. Vendedor)',
                             'F1': ws['F1'].value}, ensure_ascii=False))
    actual = str(ws['AG1'].value or '').strip()
    if actual and actual.upper() != 'PROVEEDORES':
        sys.exit(json.dumps({'error': f'AG1 ya tiene otro titulo ("{actual}"): no lo piso'}, ensure_ascii=False))
    ws['AG1'] = 'PROVEEDORES'
    ws['AG1']._style = copy(ws['AF1']._style)
    ws.column_dimensions['AG'].width = 22
    ref = ws.auto_filter.ref
    if ref:
        m = re.match(r'^([A-Z]+)(\d+):[A-Z]+(\d+)$', ref)
        if m:
            ws.auto_filter.ref = f'{m.group(1)}{m.group(2)}:AG{m.group(3)}'
    wb.save(papel)
    print(json.dumps({'papel': papel, 'AG1': 'PROVEEDORES', 'filtro': ws.auto_filter.ref,
                      'ya_estaba': actual.upper() == 'PROVEEDORES'}, ensure_ascii=False))


if __name__ == '__main__':
    main()
