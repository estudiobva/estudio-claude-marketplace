"""Etapa 2 (DDJJ de IVA), paso 4: filtra en la solapa "Compras Zetti" las facturas de servicios
AySA (agua), EDESUR (luz) y METROGAS (gas).

    python filtrar_servicios.py --papel=<papel de trabajo.xlsx>

Deja el filtro de la fila 2 aplicado en la columna Proveedor (F) con los nombres que figuran en
Zetti y oculta las demas filas de comprobantes (la fila TOTAL queda visible). Los busca por CUIT
y, por las dudas, por nombre. No cambia ningun dato. Imprime un JSON con los comprobantes.
"""
import json
import re
import sys

from openpyxl import load_workbook
from openpyxl.worksheet.filters import FilterColumn, Filters

HOJA = 'Compras Zetti'
SERVICIOS = {
    '30709565075': 'AySA (agua)',
    '30655116512': 'EDESUR (luz)',
    '30657863676': 'METROGAS (gas)',
}
NOMBRES = {'AYSA': 'AySA (agua)', 'AGUA Y SANEAMIENTOS': 'AySA (agua)', 'EDESUR': 'EDESUR (luz)',
           'METROGAS': 'METROGAS (gas)'}


def servicio(cuit, proveedor):
    c = re.sub(r'\D', '', str(cuit or ''))
    if c in SERVICIOS:
        return SERVICIOS[c]
    p = str(proveedor or '').upper()
    return next((s for k, s in NOMBRES.items() if k in p), None)


def main():
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel = opts.get('papel')
    if not papel:
        sys.exit(__doc__)
    wb = load_workbook(papel)
    if HOJA not in wb.sheetnames:
        sys.exit(json.dumps({'error': f'Falta la solapa "{HOJA}"', 'solapas': wb.sheetnames}, ensure_ascii=False))
    ws = wb[HOJA]
    tit = [str(c.value or '').strip().lower() for c in ws[2]]
    c_prov, c_cuit = tit.index('proveedor'), tit.index('cuit')
    c_tot = tit.index('total')

    encontrados, nombres = [], []
    ult = 2
    for r in range(3, ws.max_row + 1):
        fecha = ws.cell(r, 1).value
        if fecha is None:
            continue  # fila TOTAL o vacias
        ult = r
        prov = ws.cell(r, c_prov + 1).value
        s = servicio(ws.cell(r, c_cuit + 1).value, prov)
        ws.row_dimensions[r].hidden = s is None
        if s:
            encontrados.append({'fila': r, 'servicio': s, 'fecha': fecha.strftime('%d/%m/%Y'),
                                'comprobante': f"{ws.cell(r, 2).value} {ws.cell(r, 3).value} "
                                               f"{ws.cell(r, 4).value}-{ws.cell(r, 5).value}",
                                'proveedor': prov, 'cuit': ws.cell(r, c_cuit + 1).value,
                                'total': ws.cell(r, c_tot + 1).value})
            if prov not in nombres:
                nombres.append(prov)

    ws.auto_filter.ref = f'A2:{ws.cell(2, ws.max_column).column_letter}{ult}'
    ws.auto_filter.filterColumn = []
    if nombres:
        ws.auto_filter.filterColumn.append(FilterColumn(colId=c_prov, filters=Filters(filter=nombres)))
    else:
        for r in range(3, ult + 1):
            ws.row_dimensions[r].hidden = False
    wb.save(papel)

    faltan = sorted(set(SERVICIOS.values()) - {e['servicio'] for e in encontrados})
    print(json.dumps({'papel': papel, 'filtro': nombres, 'comprobantes': encontrados,
                      'sin_comprobantes_en_zetti': faltan}, ensure_ascii=False, indent=1, default=str))


if __name__ == '__main__':
    main()
