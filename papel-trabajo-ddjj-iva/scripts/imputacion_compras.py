"""Etapa 2 (DDJJ de IVA), paso 8: solapa "imputacion compras" del papel de trabajo.

    python imputacion_compras.py --papel=<papel de trabajo.xlsx>

Resumen de "Compras portal iva" segun la clasificacion de proveedores de la columna AG (paso 7).
Si la solapa ya existe (plantilla del equipo) se respetan sus rotulos y formato; si no, se crea
con el mismo esquema:

    A1  Resumen por clasificacion
    2   Clasificacion | Comprobantes | Neto Gravado | No Gravado | Exento | IVA |
        Percepciones y otros impuestos | Importe Total     (C + D + E + F + G = H en cada fila)
    3   CMV           (AG = "CMV")
    4   SERVICIOS     (todo lo demas)
    5   Total
    A7  Detalle Servicios por cuenta contable
    8   Cuenta contable | Comprobantes | gravado + exento
    9.. una fila por cuenta (las de la plantilla mas las que aparezcan en AG y falten, en orden
        alfabetico), y al pie la fila Total.

Fuentes (SUMIFS vivas contra 'Compras portal iva', filas 2 a la ultima de comprobantes):
    Neto Gravado = T + U + W + Y + AA + AC (netos de cada alicuota; no AE, que en las filas agregadas a mano
        suele quedar vacia); Exento = L; Importe Total = H;
    IVA = V + X + Z + AB + AD (IVA de cada alicuota, no "Total IVA");
    No Gravado = K, y en los comprobantes sin IVA discriminado (tipos 6, 8, 11, 13, 15, 111) ademas
        H - (N a S), igual que el cuadro resumen del paso 7 (el Portal los trae solo en Importe Total);
    Percepciones y otros impuestos = N + O + P + Q + R + S (solo en el resumen global, para que cierre);
    gravado + exento = K + L + T + U + Y + AA + AC, mas el mismo agregado de los tipos sin IVA (H - N a S:
        las facturas C entran por su total). El detalle NUNCA incluye percepciones ni otros impuestos.

No hay Excel para recalcular: el script calcula lo mismo en Python e imprime los importes y los
controles (detalle = fila SERVICIOS; neto + no grav. + exento + IVA + N a S = Importe Total).
Sale con codigo 2 si un control no cierra ($1 o mas). Si hay filas sin clasificar en AG corta sin grabar.
"""
import json
import re
import sys
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string as ci

HOJA_P = 'Compras portal iva'
HOJA = 'imputacion compras'
SIN_IVA = (6, 8, 11, 13, 15, 111)
NS = ['N', 'O', 'P', 'Q', 'R', 'S']
IVAS = ['V', 'X', 'Z', 'AB', 'AD']
GRAV_EXENTO = ['K', 'L', 'T', 'U', 'Y', 'AA', 'AC']
NETOS = ['T', 'U', 'W', 'Y', 'AA', 'AC']
FMT = '#,##0.00;[Red]-#,##0.00'
PERC_TIT = 'Percepciones y otros impuestos'
CELESTE = PatternFill('solid', fgColor='DDEBF7')
F = dict(name='Aptos Narrow', size=11)
FINA = Side(style='thin')


def fallar(msg, **extra):
    sys.exit(json.dumps({'error': msg, **extra}, ensure_ascii=False))


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def main():
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel = opts.get('papel')
    if not papel:
        sys.exit(__doc__)
    wb = load_workbook(papel)
    hp = next((h for h in wb.sheetnames if h.strip().lower() == HOJA_P.lower()), None)
    if hp is None:
        fallar(f'Falta la solapa "{HOJA_P}"', solapas=wb.sheetnames)
    wp = wb[hp]
    if str(wp['AG1'].value or '').strip().upper() != 'PROVEEDORES':
        fallar('AG1 de Compras portal iva no dice PROVEEDORES: correr antes los pasos 6 y 7')
    filas = [r for r in range(2, wp.max_row + 1) if isinstance(wp.cell(r, 1).value, datetime)]
    if not filas or filas != list(range(2, filas[-1] + 1)):
        fallar('Los comprobantes de Compras portal iva no estan seguidos desde la fila 2')
    u = filas[-1]
    sin = [r for r in filas if str(wp.cell(r, ci('AG')).value or '').strip() in ('', 'NO CLASIFICADO')]
    if sin:
        fallar('Hay proveedores sin clasificar en AG: primero hay que clasificarlos (paso 7)', filas=sin[:30])

    datos = []
    for r in filas:
        g = lambda c: wp.cell(r, ci(c)).value  # noqa: E731
        d = {c: num(g(c)) for c in ['H', 'K', 'L'] + NS + IVAS + GRAV_EXENTO + NETOS}
        d['tipo'], d['cuenta'] = g('B'), str(g('AG')).strip()
        datos.append(d)

    # --- solapa: la plantilla si existe, si no se crea con el mismo esquema
    if HOJA in wb.sheetnames:
        ws = wb[HOJA]
    else:
        ws = wb.create_sheet(HOJA, wb.sheetnames.index(hp) + 1)
        ws['A1'] = 'Resumen por clasificación'
        for j, t in enumerate(['Clasificación', 'Comprobantes', 'Neto Gravado', 'No Gravado', 'Exento', 'IVA',
                               PERC_TIT, 'Importe Total'], start=1):
            ws.cell(2, j, t)
        ws['A3'], ws['A4'], ws['A5'] = 'CMV', 'SERVICIOS', 'Total'
        ws['A7'] = 'Detalle Servicios por cuenta contable'
        for j, t in enumerate(['Cuenta contable', 'Comprobantes', 'gravado + exento'], start=1):
            ws.cell(8, j, t)
        for c in ('A1', 'A7'):
            ws[c].font = Font(bold=True, **F)
        for r, n in ((2, 8), (8, 3)):
            for j in range(1, n + 1):
                ws.cell(r, j).font, ws.cell(r, j).fill = Font(bold=True, **F), CELESTE
        ws.column_dimensions['A'].width = 37
        for L in 'BCDEFGH':
            ws.column_dimensions[L].width = 16
    if str(ws['G2'].value or '').strip().lower() == 'importe total':
        # plantilla vieja (Importe Total en G): se agrega la columna de percepciones y el total pasa a H
        ws['H2']._style = copy(ws['G2']._style)
        ws['H2'] = 'Importe Total'
        ws['G2'] = PERC_TIT
        ws.column_dimensions['H'].width = ws.column_dimensions['G'].width or 16
        ws.column_dimensions['G'].width = max(ws.column_dimensions['G'].width or 0, 18)
        for r in (3, 4, 5):
            ws.cell(r, 8)._style = copy(ws.cell(r, 7)._style)

    q = f"'{hp}'!"

    def rng(c):
        return f'{q}${c}$2:${c}${u}'

    def sumifs(col, crit):
        return f'SUMIFS({rng(col)},{rng("AG")},"{crit}")'

    def sin_iva(crit):
        """H - (N a S) de los comprobantes sin IVA discriminado."""
        tipos = '{' + ','.join(str(t) for t in SIN_IVA) + '}'
        f = f'SUM(SUMIFS({rng("H")},{rng("AG")},"{crit}",{rng("B")},{tipos}))'
        for c in NS:
            f += f'-SUM(SUMIFS({rng(c)},{rng("AG")},"{crit}",{rng("B")},{tipos}))'
        return f

    def py(pred, cols, con_sin_iva=False):
        tot = 0.0
        for d in datos:
            if pred(d):
                tot += sum(d[c] for c in cols)
                if con_sin_iva and d['tipo'] in SIN_IVA:
                    tot += d['H'] - sum(d[c] for c in NS)
        return round(tot, 2)

    def escribir(c, v, bold=False, fmt=FMT):
        cel = ws.cell(*c) if isinstance(c, tuple) else ws[c]
        cel.value = v
        f = copy(cel.font) if cel.has_style else Font(**F)
        f.b = bold or f.b
        cel.font = f
        if fmt:
            cel.number_format = fmt

    # --- resumen global (filas 3 a 5)
    resumen = {}
    for r, (rot, crit, pred) in ((3, ('CMV', 'CMV', lambda d: d['cuenta'] == 'CMV')),
                                 (4, ('SERVICIOS', '<>CMV', lambda d: d['cuenta'] != 'CMV'))):
        escribir((r, 1), rot, fmt=None)
        escribir((r, 2), f'=COUNTIFS({rng("AG")},"{crit}")', fmt='0')
        escribir((r, 3), '=' + '+'.join(sumifs(c, crit) for c in NETOS))
        escribir((r, 4), '=' + sumifs('K', crit) + '+' + sin_iva(crit))
        escribir((r, 5), '=' + sumifs('L', crit))
        escribir((r, 6), '=' + '+'.join(sumifs(c, crit) for c in IVAS))
        escribir((r, 7), '=' + '+'.join(sumifs(c, crit) for c in NS))
        escribir((r, 8), '=' + sumifs('H', crit))
        resumen[rot] = {'comprobantes': sum(1 for d in datos if pred(d)), 'neto_gravado': py(pred, NETOS),
                        'no_gravado': py(pred, ['K'], True), 'exento': py(pred, ['L']), 'iva': py(pred, IVAS),
                        'importe_total': py(pred, ['H']),
                        'otros_N_a_S': py(pred, NS)}
    escribir((5, 1), 'Total', bold=True, fmt=None)
    for j in range(2, 9):
        L = chr(64 + j)
        escribir((5, j), f'=SUM({L}3:{L}4)', bold=True, fmt='0' if j == 2 else FMT)

    # --- detalle por cuenta (fila 9 en adelante)
    existentes, r = [], 9
    while ws.cell(r, 1).value not in (None, '') and str(ws.cell(r, 1).value).strip().lower() != 'total':
        existentes.append(str(ws.cell(r, 1).value).strip())
        r += 1
    fin_viejo = r if str(ws.cell(r, 1).value or '').strip().lower() == 'total' else r - 1
    estilo_fila = [copy(ws.cell(9, j)._style) for j in (1, 2, 3)] if existentes else None
    en_ag = sorted({d['cuenta'] for d in datos if d['cuenta'] != 'CMV'})
    agregadas = [c for c in en_ag if c.upper() not in {e.upper() for e in existentes}]
    cuentas = sorted(existentes + agregadas, key=lambda s: s.upper())
    for rr in range(9, max(fin_viejo, 9 + len(cuentas)) + 2):  # limpiar el bloque anterior
        for j in (1, 2, 3):
            ws.cell(rr, j).value = None
    detalle = []
    for i, cta in enumerate(cuentas):
        rr = 9 + i
        if estilo_fila:
            for j, st in zip((1, 2, 3), estilo_fila):
                ws.cell(rr, j)._style = copy(st)
        escribir((rr, 1), cta, fmt=None)
        escribir((rr, 2), f'=COUNTIFS({rng("AG")},"{cta}")', fmt='0')
        escribir((rr, 3), '=' + '+'.join(sumifs(c, cta) for c in GRAV_EXENTO) + '+' + sin_iva(cta))
        pred = lambda d, cta=cta: d['cuenta'].upper() == cta.upper()  # noqa: E731
        detalle.append({'cuenta': cta, 'comprobantes': sum(1 for d in datos if pred(d)),
                        'gravado_mas_exento': py(pred, GRAV_EXENTO, True),
                        **({'agregada': True} if cta in agregadas else {})})
    rt = 9 + len(cuentas)
    escribir((rt, 1), 'Total', bold=True, fmt=None)
    escribir((rt, 2), f'=SUM(B9:B{rt - 1})', bold=True, fmt='0')
    escribir((rt, 3), f'=SUM(C9:C{rt - 1})', bold=True)
    for j in (1, 2, 3):
        ws.cell(rt, j).border = Border(top=FINA)
    wb.save(papel)

    serv = resumen['SERVICIOS']
    tot = {k: round(resumen['CMV'][k] + serv[k], 2) for k in serv}
    control = {
        'comprobantes_detalle_vs_servicios': sum(d['comprobantes'] for d in detalle) - serv['comprobantes'],
        'componentes_vs_importe_total': round(tot['neto_gravado'] + tot['no_gravado'] + tot['exento'] + tot['iva']
                                              + tot['otros_N_a_S'] - tot['importe_total'], 2),
    }
    print(json.dumps({
        'papel': papel, 'datos': f"'{hp}' filas 2 a {u}",
        'resumen': {**resumen, 'Total': tot},
        'detalle': detalle,
        'cuentas_agregadas': agregadas,
        'cuentas_en_cero': [d['cuenta'] for d in detalle if d['comprobantes'] == 0],
        'control': control,
    }, ensure_ascii=False, indent=1))
    if control['comprobantes_detalle_vs_servicios'] or abs(control['componentes_vs_importe_total']) >= 1:
        sys.exit(2)


if __name__ == '__main__':
    main()
