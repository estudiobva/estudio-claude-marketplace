"""Resume la solapa "cruce ventas" de un papel de trabajo (lo usa papel-ventas.js).

    python resumen_cruce.py <papel.xlsx> [<salida de cruce_3_fuentes.py.json>]

Imprime un JSON con: Zetti - Portal IVA del mes, comprobantes de otros periodos,
cantidad por estado y las observaciones que NO son corrimientos ni redondeos
(las que hay que mirar), hasta 40.
"""
import json
import sys

import openpyxl


def main():
    if len(sys.argv) not in (2, 3):
        sys.exit(__doc__)
    cruce = {}
    if len(sys.argv) == 3:
        try:
            cruce = json.load(open(sys.argv[2], encoding='utf-8'))
        except (OSError, ValueError):
            cruce = {}
    wb = openpyxl.load_workbook(sys.argv[1], read_only=True)
    if 'cruce ventas' not in wb.sheetnames:
        sys.exit(json.dumps({'error': 'El papel no tiene la solapa "cruce ventas"'}, ensure_ascii=False))
    ws = wb['cruce ventas']
    en_detalle, obs, estados, total = False, [], {}, None
    for r in ws.iter_rows(max_col=20, values_only=True):
        if r[0] == 'Zetti - Portal IVA':
            total = r[6]
        if r[0] == 'Fecha':
            en_detalle = True
            continue
        if en_detalle and r[0] == 'TOTAL':
            break
        if en_detalle and r[0]:
            estados[str(r[3])] = estados.get(str(r[3]), 0) + 1
            o = str(r[4])
            if 'No es una diferencia' not in o and 'redondeo' not in o and 'Corrimiento' not in str(r[3]):
                obs.append(f'{str(r[0])[:10]} {str(r[1])[:40]} | {o[:300]}')
    wb.close()
    print(json.dumps({
        'zetti_menos_portal': total,
        'otros_periodos': cruce.get('otros_periodos'),
        'estados': estados,
        'para_revisar': obs[:40],
        'para_revisar_total': len(obs),
    }, ensure_ascii=False, default=str))


if __name__ == '__main__':
    main()
