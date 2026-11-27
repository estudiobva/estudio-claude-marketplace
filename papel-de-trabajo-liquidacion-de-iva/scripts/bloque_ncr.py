"""Detalle de las NCR eliminadas, debajo del cuadro resumen de "Compras portal iva".

    python bloque_ncr.py --papel=<papel.xlsx> --ncr=<AAAAMM-COMPRAS-SIN-NCR-<SOC>.xlsx>
    python bloque_ncr.py --papel=<papel.xlsx> --sin-ncr      (periodo sin NCR para eliminar)

Toma de la solapa "compras con NCR" las notas de credito que la conciliacion NCR (hoja "Eliminados") restó y las lista
debajo del cuadro de "Compras portal iva" (desde la columna U, dos filas despues de lo ultimo escrito): fecha,
comprobante, drogueria, CUIT, Importe Total y su apertura, fila TOTAL (=SUM) y un control:
compras con NCR - NCR = Compras portal iva (sin los servicios agregados de Zetti).
Titulo "NCR ELIMINADAS". Correr DESPUES de cuadro_compras.py (ese script borra todo lo que hay debajo de la nota).
Si se vuelve a correr, rehace el bloque. Imprime un JSON.
"""
import argparse
import json
import re
import sys

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

TIPOS_NC = {3, 8, 13, 53, 112, 113, 114, 203}
FMT = '#,##0.00;[Red]-#,##0.00'


def fallar(msg, **extra):
    print(json.dumps({'ok': False, 'error': msg, **extra}, ensure_ascii=False, default=str))
    sys.exit(1)


def cuit(v):
    return re.sub(r'\D', '', str(v or '')).lstrip('0')


def num(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def f(v):
    return float(v or 0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--papel', required=True)
    ap.add_argument('--ncr')
    ap.add_argument('--sin-ncr', action='store_true')
    a = ap.parse_args()
    if not a.ncr and not a.sin_ncr:
        fallar('Falta --ncr (conciliacion NCR) o --sin-ncr')

    ncr = {}
    if not a.sin_ncr:
        wn = load_workbook(a.ncr, read_only=True, data_only=True)
        if 'Eliminados' not in wn.sheetnames:
            fallar('La conciliacion NCR no tiene la hoja "Eliminados"', archivo=a.ncr)
        filas = list(wn['Eliminados'].iter_rows(values_only=True))
        h = [str(x or '').strip().lower() for x in filas[0]] if filas else []
        ix = {k: h.index(k) for k in ('_drog', '_cuit', '_pv', '_nro')} if filas else {}
        ncr = {(cuit(r[ix['_cuit']]), num(r[ix['_pv']]), num(r[ix['_nro']])): str(r[ix['_drog']] or '').strip().upper()
               for r in filas[1:] if r and r[ix['_cuit']]}

    wb = load_workbook(a.papel)
    for n in ('compras con NCR', 'Compras portal iva'):
        if n not in wb.sheetnames:
            fallar(f'Falta la solapa "{n}"', solapas=wb.sheetnames)
    wc, wp = wb['compras con NCR'], wb['Compras portal iva']

    # Columnas del libro (A..AF, como lo baja ARCA): B tipo, C PV, D numero, F CUIT, G denominacion, H total,
    # K no gravado, L exento, N..S percepciones/impuestos, T,U,W,Y,AA,AC netos, V,X,Z,AB,AD IVA.
    def tot_con(ws):
        ft = next(r for r in range(2, ws.max_row + 1) if str(ws.cell(r, 1).value or '').strip().upper() == 'TOTAL')
        return ft, [r for r in range(2, ft) if isinstance(ws.cell(r, 2).value, (int, float))]
    _, datos_con = tot_con(wc)
    lista = []
    for r in datos_con:
        if num(wc.cell(r, 2).value) not in TIPOS_NC:
            continue
        k = (cuit(wc.cell(r, 6).value), num(wc.cell(r, 3).value), num(wc.cell(r, 4).value))
        if k in ncr:
            g = lambda c: f(wc.cell(r, c).value)  # noqa: E731
            lista.append({
                'fecha': wc.cell(r, 1).value, 'comp': f'NC A {k[1]:05d}-{k[2]:08d}' if num(wc.cell(r, 2).value) == 3 else f'NC {k[1]:05d}-{k[2]:08d}',
                'drog': ncr[k], 'prov': wc.cell(r, 7).value, 'cuit': wc.cell(r, 6).value,
                'total': g(8), 'nogex': g(11) + g(12),
                'neto': sum(g(c) for c in (20, 21, 23, 25, 27, 29)), 'iva': sum(g(c) for c in (22, 24, 26, 28, 30)),
                'otros': sum(g(c) for c in range(14, 20))})
    if len(lista) != len(ncr):
        fallar('No encontre en "compras con NCR" todas las NCR de la conciliacion', conciliacion=len(ncr), encontradas=len(lista))
    lista.sort(key=lambda x: (x['drog'], x['fecha'], x['comp']))

    # Lugar: dos filas debajo de lo ultimo escrito desde la columna U (cuadro, otros conceptos y control).
    ult = max(r for r in range(1, wp.max_row + 1) if any(wp.cell(r, c).value not in (None, '') for c in range(21, 32)))
    # re-ejecucion: si el bloque ya estaba, se borra y se rehace
    for r in range(1, wp.max_row + 1):
        if str(wp.cell(r, 21).value or '').startswith(('NCR RESTADAS', 'NCR ELIMINADAS')):
            for rr in range(r, wp.max_row + 1):
                for c in range(21, 32):
                    wp.cell(rr, c).value = None
                    wp.cell(rr, c).style = 'Normal'
            ult = max(x for x in range(1, r) if any(wp.cell(x, c).value not in (None, '') for c in range(21, 32)))
            break
    r0 = ult + 3

    fuente = Font(name='Aptos Narrow', size=11)
    negrita = Font(name='Aptos Narrow', size=11, bold=True)
    rojo = Font(name='Aptos Narrow', size=11, color='FFFF0000')
    gris = PatternFill('solid', fgColor='D9D9D9')
    fino = Side(style='thin')
    grueso = Side(style='medium')

    wp.cell(r0, 21, 'NCR ELIMINADAS (notas de credito de recupero de Monroe y Suizo, conciliacion NCR)').font = negrita
    tit = ['Fecha', 'Comprobante', 'Drogueria', 'CUIT', 'Importe Total', 'No grav. + Exento', 'Neto Gravado', 'IVA',
           'Percepciones y otros', 'Proveedor (Portal IVA)']
    hr = r0 + 1
    for j, t in enumerate(tit):
        c = wp.cell(hr, 21 + j, t)
        c.font, c.fill, c.alignment = negrita, gris, Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = Border(top=fino, bottom=fino, left=fino, right=fino)
    r = hr
    if not lista:
        r += 1
        wp.cell(r, 21, 'No hay NCR eliminadas en este periodo' + (' (sin conciliacion NCR)' if a.sin_ncr else '')).font = fuente
    for x in lista:
        r += 1
        vals = [x['fecha'], x['comp'], x['drog'], x['cuit'], x['total'], x['nogex'], x['neto'], x['iva'], x['otros'], x['prov']]
        for j, v in enumerate(vals):
            c = wp.cell(r, 21 + j, v)
            c.font = rojo if j in (1, 2) else fuente
            if j == 0:
                c.number_format = 'dd/mm/yyyy'
            elif 4 <= j <= 8:
                c.number_format = FMT
    rt = r + 1
    wp.cell(rt, 21, 'TOTAL NCR ELIMINADAS').font = negrita
    wp.cell(rt, 23, f'{len(lista)} NCR').font = negrita
    for c in range(25, 30):
        x = wp.cell(rt, c, f'=SUM({L(c)}{hr + 1}:{L(c)}{r})' if lista else 0)
        x.font, x.number_format = negrita, FMT
    for c in range(21, 31):
        wp.cell(rt, c).border = Border(top=fino, bottom=grueso)

    # Control: compras con NCR - NCR = Compras portal iva sin los servicios agregados de Zetti
    ft_con, _ = tot_con(wc)
    ft_p, datos_p = tot_con(wp)
    rc = rt + 2
    wp.cell(rc, 21, 'CONTROL: compras con NCR - NCR eliminadas = Compras portal iva (antes de agregar los servicios de Zetti)').font = negrita
    filas_ctrl = [
        ('Importe Total solapa "compras con NCR"', f"='compras con NCR'!H{ft_con}"),
        ('NCR eliminadas', f'=+Y{rt}'),
        ('Servicios agregados de Compras Zetti', None),
        ('Importe Total "Compras portal iva" (columna H)', f'=+H{ft_p}'),
        ('Diferencia', None),
    ]
    # servicios agregados = filas del Portal que no estan en "compras con NCR" (CUIT + tipo + PV + numero)
    claves_con = {(cuit(wc.cell(rr, 6).value), num(wc.cell(rr, 2).value), num(wc.cell(rr, 3).value), num(wc.cell(rr, 4).value)) for rr in datos_con}
    agregadas = [rr for rr in datos_p if (cuit(wp.cell(rr, 6).value), num(wp.cell(rr, 2).value), num(wp.cell(rr, 3).value), num(wp.cell(rr, 4).value)) not in claves_con]
    serv = '=' + ('+'.join(f'H{rr}' for rr in agregadas) if agregadas else '0')
    filas_ctrl[2] = ('Servicios agregados de Compras Zetti', serv)
    for i, (t, fo) in enumerate(filas_ctrl):
        rr = rc + 1 + i
        wp.cell(rr, 21, t).font = negrita if t == 'Diferencia' else fuente
        if t == 'Diferencia':
            fo = f'=+Y{rc + 1}-Y{rc + 2}+Y{rc + 3}-Y{rc + 4}'
        c = wp.cell(rr, 25, fo)
        c.font, c.number_format = (negrita if t == 'Diferencia' else fuente), FMT
        for col in range(21, 26):
            wp.cell(rr, col).border = Border(top=fino, bottom=fino)

    # verificacion en Python (no hay Excel para recalcular)
    total_con = sum(f(wc.cell(rr, 8).value) for rr in datos_con)
    total_ncr = sum(x['total'] for x in lista)
    total_serv = sum(f(wp.cell(rr, 8).value) for rr in agregadas)
    total_p = sum(f(wp.cell(rr, 8).value) for rr in datos_p)
    dif = round(total_con - total_ncr + total_serv - total_p, 2)
    apertura = round(sum(x['nogex'] + x['neto'] + x['iva'] + x['otros'] - x['total'] for x in lista), 2)
    for col in range(21, 31):
        wp.column_dimensions[L(col)].width = max(wp.column_dimensions[L(col)].width or 0, 14)

    wb.save(a.papel)
    por = {}
    for x in lista:
        p = por.setdefault(x['drog'], {'ncr': 0, 'importe': 0.0})
        p['ncr'] += 1
        p['importe'] = round(p['importe'] + x['total'], 2)
    print(json.dumps({'ok': dif == 0 or abs(dif) < 0.05, 'bloque_desde_fila': r0, 'ncr': len(lista), 'por_drogueria': por,
                      'importe_ncr': round(total_ncr, 2), 'control': {'compras_con_ncr': round(total_con, 2),
                      'ncr': round(total_ncr, 2), 'servicios_agregados': round(total_serv, 2), 'compras_portal_iva': round(total_p, 2),
                      'diferencia': dif}, 'apertura_vs_total': apertura}, ensure_ascii=False, default=str))


if __name__ == '__main__':
    main()
