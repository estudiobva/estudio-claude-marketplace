"""Etapa 2 (DDJJ de IVA), paso 7: cuenta contable para proveedores NO CLASIFICADOS.

    python asignar_cuentas.py --papel=<papel de trabajo.xlsx> --cuentas="<CUIT>=<CUENTA>;<CUIT>=<CUENTA>"

En "Compras portal iva" escribe la cuenta indicada en la columna AG (PROVEEDORES) de cada
comprobante de ese CUIT (columna F) que este NO CLASIFICADO o vacio. No toca las filas que ya
tienen cuenta ni el Maestro de Proveedores (para que quede para la proxima, cargarlo en el
Maestro). Imprime un JSON con lo asignado y los NO CLASIFICADOS que quedan.
"""
import json
import re
import sys
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font

HOJA = 'Compras portal iva'
NO_CLASIFICADO = 'NO CLASIFICADO'
COL_CUIT, COL_NOMBRE, COL_PROV = 6, 7, 33  # F, G, AG


def cuit_norm(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        v = str(int(v))
    return re.sub(r'\D', '', str(v or '')).lstrip('0')


def main():
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel, txt = opts.get('papel'), opts.get('cuentas', '')
    if not papel or not txt:
        sys.exit(__doc__)
    cuentas = {}
    for par in re.split(r'[;\n]+', txt):
        if not par.strip():
            continue
        c, sep, cuenta = par.partition('=')
        if not sep or not cuit_norm(c) or not cuenta.strip():
            sys.exit(json.dumps({'error': f'No entiendo "{par}": va CUIT=CUENTA'}, ensure_ascii=False))
        cuentas[cuit_norm(c)] = cuenta.strip().upper()

    wb = load_workbook(papel)
    if HOJA not in wb.sheetnames:
        sys.exit(json.dumps({'error': f'Falta la solapa "{HOJA}"'}, ensure_ascii=False))
    ws = wb[HOJA]
    if str(ws.cell(1, COL_PROV).value or '').strip().upper() != 'PROVEEDORES':
        sys.exit(json.dumps({'error': 'AG1 no es PROVEEDORES: correr antes los pasos 6 y 7'}, ensure_ascii=False))

    asignados, quedan = {}, {}
    for r in range(2, ws.max_row + 1):
        if not isinstance(ws.cell(r, 1).value, datetime):
            continue
        actual = str(ws.cell(r, COL_PROV).value or '').strip()
        if actual and actual != NO_CLASIFICADO:
            continue
        c = cuit_norm(ws.cell(r, COL_CUIT).value)
        if c in cuentas:
            celda = ws.cell(r, COL_PROV, cuentas[c])
            celda.font = Font(name='Arial', size=10)
            asignados[c] = asignados.get(c, 0) + 1
        else:
            q = quedan.setdefault(c, {'cuit': c, 'nombre': ws.cell(r, COL_NOMBRE).value, 'comprobantes': 0})
            q['comprobantes'] += 1
    no_usados = sorted(set(cuentas) - set(asignados))
    wb.save(papel)
    print(json.dumps({'papel': papel,
                      'asignados': [{'cuit': c, 'cuenta': cuentas[c], 'comprobantes': n} for c, n in asignados.items()],
                      'cuits_sin_no_clasificados': no_usados,
                      'no_clasificados': list(quedan.values())}, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
