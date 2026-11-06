"""Proceso TSH / IIBB: depura la solapa "Ventas Zetti (depurado)" contra el Portal IVA.

    python depurar_pv.py <papel.xlsx> [--excepcion="2:NC B"] [--simular]

Trabaja SOLO sobre la solapa "Ventas Zetti (depurado)"; "Ventas Zetti" y
"Ventas portal iva" no se tocan.

1. Compara los puntos de venta de la columna D (Zetti depurado) con la columna C
   ("Punto de Venta") de "Ventas portal iva".
2. Elimina de la depurada los comprobantes cuyo punto de venta no figura en el
   Portal IVA: son comprobantes que Zetti registra (fallidos, PV 1, 2, 999999...)
   y que no se informan a ARCA.
3. Excepciones (las aclara la persona al pedir la tarea; sin aclaracion no hay):
   --excepcion="PV:TC M[,TC M...]" elimina de ese punto de venta solo esos tipos de
   comprobante, figure o no el PV en el Portal. Ej. SANAR 66: --excepcion="2:NC B"
   borra las NC B del PV 2 y deja las Z. Varias excepciones separadas por ";".
4. Debajo de los datos deja tres filas, de I (Exen) a M (Total):
     TOTAL ZETTI       =SUM de la depurada (fondo beige)
     TOTAL PORTAL IVA  totales del Portal agrupados como Zetti (fondo celeste):
                       Exen  = No Gravado + Exento
                       Grav  = Neto Gravado de todas las alicuotas (0 a 27%)
                       IVA   = IVA de todas las alicuotas (no la columna Total IVA de ARCA)
                       P.IB  = Percepciones de IIBB + Per. o Pagos a Cta. de Otros Imp. Nac.
                       Total = Importe Total
     diferencias       TOTAL ZETTI - TOTAL PORTAL IVA
   Todo con formulas, con el rotulo en la columna N, en cursiva.

--simular informa que eliminaria sin guardar. Imprime un JSON con lo eliminado.
"""
import json
import sys
from collections import defaultdict
from copy import copy

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

import fuente
from openpyxl.utils import get_column_letter

HOJA_Z = 'Ventas Zetti (depurado)'
HOJA_P = 'Ventas portal iva'
BEIGE = PatternFill('solid', fgColor='DDD9C4')
CELESTE = PatternFill('solid', fgColor='B8CCE4')
CURSIVA = fuente.letra(fuente.ZETTI, italic=True)
FMT = '#,##0.00'  # formato numero, como el Zetti original (no contabilidad)


def norm_pv(v):
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v or '').strip()
    return int(s) if s.isdigit() else s


def excepciones(arg):
    out = {}
    for parte in filter(None, (p.strip() for p in (arg or '').split(';'))):
        pv, _, tipos = parte.partition(':')
        out[norm_pv(pv)] = {' '.join(t.split()).upper() for t in tipos.split(',') if t.strip()}
    return out


def main():
    pos = [a for a in sys.argv[1:] if not a.startswith('--')]
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    if len(pos) != 1:
        sys.exit(__doc__)
    papel = pos[0]
    exc = excepciones(opts.get('excepcion'))
    simular = 'simular' in opts

    wb = load_workbook(papel)
    wz, wp = wb[HOJA_Z], wb[HOJA_P]

    # --- Portal: encabezados, puntos de venta y fila TOTAL ---
    tit = {str(c.value).strip(): c.column for c in wp[1] if c.value}
    col_pv = tit['Punto de Venta']
    pv_portal, fila_tot_p = set(), None
    for r in range(2, wp.max_row + 1):
        a = wp.cell(r, 1).value
        if isinstance(a, str) and a.strip().upper() == 'TOTAL':
            fila_tot_p = r
            break
        if a is not None:
            pv_portal.add(norm_pv(wp.cell(r, col_pv).value))
    if fila_tot_p is None:
        sys.exit(json.dumps({'error': f'No encontre la fila TOTAL en "{HOJA_P}".'}))

    def ref(*nombres):
        cols = [tit[n] for n in nombres if n in tit]
        faltan = [n for n in nombres if n not in tit]
        if faltan:
            sys.exit(json.dumps({'error': f'Faltan columnas en "{HOJA_P}": {faltan}'}, ensure_ascii=False))
        return '=' + '+'.join(f"'{HOJA_P}'!{get_column_letter(c)}{fila_tot_p}" for c in cols)

    netos = [n for n in tit if n.startswith('Neto Gravado IVA')]
    ivas = [n for n in tit if n.startswith('Importe IVA ')]
    portal = {
        9: ref('Importe No Gravado', 'Importe Exento'),
        10: ref(*netos),
        11: ref(*ivas),
        12: ref('Importe de Percepciones de Ingresos Brutos', 'Importe de Per. o Pagos a Cta. de Otros Imp. Nac.'),
        13: ref('Importe Total'),
    }

    # --- Zetti depurado: filas de datos (desde la 3 hasta la primera sin fecha) ---
    ult = 2
    while wz.cell(ult + 1, 1).value is not None:
        ult += 1
    ncol = 14
    filas = [[wz.cell(r, c) for c in range(1, ncol + 1)] for r in range(3, ult + 1)]

    eliminadas = defaultdict(lambda: {'comprobantes': 0, 'total': 0.0, 'tipos': defaultdict(int)})
    quedan = []
    for f in filas:
        pv = norm_pv(f[3].value)
        tipo = f'{f[1].value} {f[2].value}'.upper()
        borrar = (tipo in exc[pv]) if pv in exc else (pv not in pv_portal)
        if borrar:
            e = eliminadas[str(pv)]
            e['comprobantes'] += 1
            e['total'] = round(e['total'] + (f[12].value or 0), 2)
            e['tipos'][tipo] += 1
        else:
            quedan.append([(c.value, c._style) for c in f])

    res = {
        'pv_solo_en_zetti': sorted({str(norm_pv(f[3].value)) for f in filas} - {str(p) for p in pv_portal}),
        'excepciones': {str(k): sorted(v) for k, v in exc.items()},
        'eliminadas': {k: {**v, 'tipos': dict(v['tipos'])} for k, v in eliminadas.items()},
        'comprobantes_antes': len(filas), 'comprobantes_despues': len(quedan),
    }
    if simular:
        print(json.dumps(res, ensure_ascii=False))
        return

    # --- Reescribe la depurada: datos que quedan + las 3 filas de totales ---
    estilo_vacio = copy(wz.cell(1, 20)._style)
    for r in range(3, wz.max_row + 1):
        for c in range(1, ncol + 1):
            cel = wz.cell(r, c)
            cel.value = None
            cel._style = copy(estilo_vacio)
    for i, fila in enumerate(quedan, start=3):
        for j, (v, st) in enumerate(fila, start=1):
            cel = wz.cell(i, j, v)
            cel._style = copy(st)
    nult = 2 + len(quedan)
    fz, fp, fd = nult + 1, nult + 2, nult + 3
    for j in range(9, 14):
        col = get_column_letter(j)
        wz.cell(fz, j, f'=SUM({col}3:{col}{nult})')
        wz.cell(fp, j, portal[j])
        wz.cell(fd, j, f'={col}{fz}-{col}{fp}')
    for fila, rotulo, fondo in ((fz, 'TOTAL ZETTI', BEIGE), (fp, 'TOTAL PORTAL IVA', CELESTE), (fd, 'diferencias', None)):
        wz.cell(fila, 14, rotulo)
        for j in range(9, 15):
            cel = wz.cell(fila, j)
            cel.font = CURSIVA
            if j < 14:
                cel.number_format = FMT
            if fondo:
                cel.fill = fondo
    if wz.max_row > fd:
        wz.delete_rows(fd + 1, wz.max_row - fd)
    wz.freeze_panes = 'A3'
    wz.auto_filter.ref = f'A2:N{nult}'
    fuente.aplicar(wz, fuente.ZETTI)

    wb.save(papel)
    res.update({'fila_total_zetti': fz, 'fila_total_portal': fp, 'fila_diferencias': fd})
    print(json.dumps(res, ensure_ascii=False))


if __name__ == '__main__':
    main()
