"""Proceso TSH / IIBB: cuadro resumen de ventas en la solapa "Ventas portal iva".

    python cuadro_resumen.py <papel.xlsx> [--factura-b-por-alicuota] [--nc-z-combinada]

Solo toca la solapa "Ventas portal iva" (las demas no). Arma el cuadro debajo de la
fila TOTAL y de la nota (sin modificarlas), dejando dos filas libres:
  - etiqueta en AC ("codigo - descripcion"), valores en AD Neto Gravado,
    AE Exento + No Grav. y AF IVA; una fila por cada codigo de comprobante que
    aparece en la columna B, y al final "TOTAL VENTAS" (con el total de las
    tres columnas en AG).
  - Formulas SUMIFS tomando el codigo del numero al inicio de la etiqueta:
      Neto Gravado      = "Total Neto Gravado" (AF)
      Exento + No Grav. = "Importe Exento" (N) + "Importe No Gravado" (M)
      IVA               = "Importe IVA 10,5%" (AA) + "Importe IVA 21%" (AC), nunca "Total IVA" (AG)
Casos especiales (se activan solos por el nombre del archivo, o con el flag):
  - Salud y Belleza (--factura-b-por-alicuota): Factura B en dos filas,
    "6 - Factura B al 21%" (AB y AC, con todo el Exento + No Grav.) y
    "6 - Factura B al 10,5%" (Z y AA, exento en 0).
  - Sanar 12 (--nc-z-combinada): "110 y 113 - Nota de Credito Z" suma los dos codigos.
Codigos que no reconoce: los agrega como "codigo - Codigo NN (revisar rotulo)" y los
informa en el JSON.

Control extra en la hoja: debajo del Importe Total de la fila TOTAL (J) va la suma
de M a AE de esa fila, y debajo del total general del cuadro (AG) la diferencia
AG - J. Da 0, salvo en agentes de retencion de IIBB , donde la
diferencia son las percepciones de las columnas O + P.

Formato: encabezado gris en negrita, recuadro, notas de credito y negativos en rojo.
Control: el total del cuadro tiene que dar igual que los totales de la planilla
(Neto Gravado, Exento + No Grav. e IVA = AA + AC). Imprime un JSON con el resultado.
"""
import json
import os
import sys
import unicodedata

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import fuente

HOJA = 'Ventas portal iva'
DESC = {
    1: 'Factura A', 2: 'Nota de Débito A', 3: 'Nota de Crédito A', 4: 'Recibo A',
    6: 'Factura B', 7: 'Nota de Débito B', 8: 'Nota de Crédito B', 9: 'Recibo B',
    11: 'Factura C', 12: 'Nota de Débito C', 13: 'Nota de Crédito C', 15: 'Recibo C',
    19: 'Factura E', 20: 'Nota de Débito E', 21: 'Nota de Crédito E',
    51: 'Factura M', 52: 'Nota de Débito M', 53: 'Nota de Crédito M',
    81: 'Ticket Z', 82: 'Ticket Z', 83: 'Ticket Z',
    110: 'Nota de Crédito Z', 112: 'Nota de Crédito Z', 113: 'Nota de Crédito Z', 114: 'Nota de Crédito Z',
    115: 'Nota de Débito Z', 116: 'Nota de Débito Z', 117: 'Nota de Débito Z',
    201: 'Factura de Crédito Electrónica A', 202: 'Nota de Débito FCE A', 203: 'Nota de Crédito FCE A',
    206: 'Factura de Crédito Electrónica B', 207: 'Nota de Débito FCE B', 208: 'Nota de Crédito FCE B',
}
NC = {3, 8, 13, 21, 53, 110, 112, 113, 114, 203, 208}
# Orden del cuadro: facturas, ND y tickets como en el papel; las NC al final.
ORDEN = [1, 6, 2, 81, 82, 83, 7, 11, 12, 19, 20, 51, 52, 201, 202, 206, 207, 115, 116, 117, 4, 9, 15,
         3, 8, 13, 21, 53, 203, 208, 110, 112, 113, 114]
FMT = '_-* #,##0.00_-;[Red]-* #,##0.00_-;_-* "-"??_-;_-@_-'
ROJO = Font(color='FF0000')
GRIS = PatternFill('solid', fgColor='D9D9D9')
GRUESA = Side(style='medium')


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '').strip().lower()).encode('ascii', 'ignore').decode()
    return ' '.join(s.split())


def main():
    pos = [a for a in sys.argv[1:] if not a.startswith('--')]
    flags = {a[2:] for a in sys.argv[1:] if a.startswith('--')}
    if len(pos) != 1:
        sys.exit(__doc__)
    papel = pos[0]
    nombre = norm(os.path.basename(papel))
    por_alicuota = 'factura-b-por-alicuota' in flags or 'salud y belleza' in nombre
    nc_z_combinada = 'nc-z-combinada' in flags or 'sanar 12' in nombre

    wb = load_workbook(papel)
    ws = wb[HOJA]
    tit = {norm(c.value): c.column_letter for c in ws[1] if c.value}
    col = {k: tit[norm(v)] for k, v in {
        'tipo': 'Tipo de Comprobante', 'ng': 'Importe No Gravado', 'ex': 'Importe Exento',
        'n105': 'Neto Gravado IVA 10,5%', 'i105': 'Importe IVA 10,5%', 'n21': 'Neto Gravado IVA 21%',
        'i21': 'Importe IVA 21%', 'neto': 'Total Neto Gravado'}.items()}

    # Fin de los datos (ultima fila con fecha) y fin de la hoja (TOTAL y nota).
    ult = 1
    while ws.cell(ult + 1, 1).value is not None and not isinstance(ws.cell(ult + 1, 1).value, str):
        ult += 1
    fin = ws.max_row
    while fin > ult and all(c.value is None for c in ws[fin]):
        fin -= 1
    if any(c.value is not None for fila in ws.iter_rows(min_row=fin + 1) for c in fila):
        sys.exit(json.dumps({'error': 'Hay celdas debajo de la nota: revisar antes de armar el cuadro.'}))

    codigos = sorted({ws.cell(r, 2).value for r in range(2, ult + 1)} - {None})
    rng = lambda k: f"${col[k]}$2:${col[k]}${ult}"
    crit = f"{rng('tipo')},"

    def sumifs(k, cod):
        return f"SUMIFS({rng(k)},{crit}{cod})"

    filas = []  # (etiqueta, neto, exento, iva, es_nc)
    vistos = set()
    for cod in sorted(codigos, key=lambda c: (ORDEN.index(c) if c in ORDEN else 50, c)):
        if cod in vistos:
            continue
        vistos.add(cod)
        desc = DESC.get(cod)
        if nc_z_combinada and cod in (110, 113):
            vistos |= {110, 113}
            par = (110, 113)
            et = '110 y 113 - Nota de Crédito Z'
            f = lambda k: '=' + '+'.join(sumifs(k, c) for c in par)
            filas.append((et, f('neto'), '=' + '+'.join(f"{sumifs('ex', c)}+{sumifs('ng', c)}" for c in par),
                          '=' + '+'.join(f"{sumifs('i105', c)}+{sumifs('i21', c)}" for c in par), True))
            continue
        if por_alicuota and cod == 6:
            filas.append(('6 - Factura B al 21%', ('n21',), ('ex', 'ng'), ('i21',), False))
            filas.append(('6 - Factura B al 10,5%', ('n105',), None, ('i105',), False))
            continue
        et = f"{cod} - {desc or f'Codigo {cod} (revisar rotulo)'}"
        filas.append((et, ('neto',), ('ex', 'ng'), ('i105', 'i21'), cod in NC))

    inicio = fin + 3
    cab = inicio
    for j, t in zip(('AD', 'AE', 'AF'), ('Neto Gravado', 'Exento + No Grav.', 'IVA')):
        c = ws[f'{j}{cab}']
        c.value, c.font, c.fill = t, Font(bold=True), GRIS
        c.alignment = Alignment(horizontal='center')
    ws[f'AC{cab}'].fill = GRIS

    r = cab + 1
    for et, neto, exen, iva, es_nc in filas:
        ws[f'AC{r}'] = et
        etq = f'$AC{r}'
        cod = f'VALUE(LEFT({etq},FIND(" ",{etq})-1))'
        for letra, partes in (('AD', neto), ('AE', exen), ('AF', iva)):
            if isinstance(partes, str):
                v = partes
            elif partes is None:
                v = 0
            else:
                v = '=' + '+'.join(f"SUMIFS({rng(k)},{crit}{cod})" for k in partes)
            c = ws[f'{letra}{r}']
            c.value, c.number_format = v, FMT
            if es_nc:
                c.font = ROJO
        if es_nc:
            ws[f'AC{r}'].font = ROJO
        r += 1
    ult_cuadro = r - 1

    # Recuadro: borde grueso alrededor de AC:AF, del encabezado a la ultima fila.
    for rr in range(cab, ult_cuadro + 1):
        for j in ('AC', 'AD', 'AE', 'AF'):
            c = ws[f'{j}{rr}']
            c.border = Border(top=GRUESA if rr == cab else None, bottom=GRUESA if rr == ult_cuadro else None,
                              left=GRUESA if j == 'AC' else None, right=GRUESA if j == 'AF' else None)

    tot = ult_cuadro + 2
    ws[f'AC{tot}'] = 'TOTAL VENTAS'
    for j in ('AD', 'AE', 'AF'):
        ws[f'{j}{tot}'] = f'=SUM({j}{cab + 1}:{j}{ult_cuadro})'
    ws[f'AG{tot}'] = f'=SUM(AD{tot}:AF{tot})'
    for j in ('AC', 'AD', 'AE', 'AF', 'AG'):
        c = ws[f'{j}{tot}']
        c.font = Font(bold=True)
        if j != 'AC':
            c.number_format = FMT
    if (ws.column_dimensions['AC'].width or 0) < 30:
        ws.column_dimensions['AC'].width = 30

    # Control contra la fila TOTAL de la planilla: debajo del Importe Total (J),
    # la suma de M a AE de esa fila; y debajo del total del cuadro (AG), la
    # diferencia. Da 0, salvo en los agentes de retencion de IIBB (ej. Salud y
    # Belleza), donde la diferencia son las percepciones (columnas O + P).
    fila_tot = next(r for r in range(ult + 1, fin + 1) if norm(ws.cell(r, 1).value) == 'total')
    ws[f'J{fila_tot + 1}'] = f'=SUM(M{fila_tot}:AE{fila_tot})'
    ws[f'AG{tot + 1}'] = f'=AG{tot}-J{fila_tot + 1}'
    for ref in (f'J{fila_tot + 1}', f'AG{tot + 1}'):
        ws[ref].number_format = FMT
    # Letra del cuadro y de los controles: Aptos Narrow 11 (el resto de la solapa queda en Arial 10)
    fuente.aplicar(ws, fuente.ZETTI, [ws[f'{j}{rr}'] for rr in range(cab, tot + 2) for j in ('AC', 'AD', 'AE', 'AF', 'AG')]
                   + [ws[f'J{fila_tot + 1}']])

    # Control en Python: el cuadro tiene que dar los totales de la planilla.
    idx = {k: ws[f'{v}1'].column - 1 for k, v in col.items()}
    datos = [[c.value for c in fila] for fila in ws.iter_rows(min_row=2, max_row=ult)]
    n = lambda f, k: f[idx[k]] or 0
    hoja = {'neto': sum(n(f, 'neto') for f in datos), 'exen': sum(n(f, 'ex') + n(f, 'ng') for f in datos),
            'iva': sum(n(f, 'i105') + n(f, 'i21') for f in datos)}
    cuadro = {'neto': 0.0, 'exen': 0.0, 'iva': 0.0}
    detalle = []
    for et, neto, exen, iva, _ in filas:
        cods = [110, 113] if et.startswith('110 y 113') else [int(et.split(' ')[0])]
        sel = [f for f in datos if f[idx['tipo']] in cods]
        def s(partes, defecto):
            if isinstance(partes, str):
                return defecto
            return sum(n(f, k) for f in sel for k in (partes or ()))
        v = {'neto': s(neto, sum(n(f, 'neto') for f in sel)),
             'exen': s(exen, sum(n(f, 'ex') + n(f, 'ng') for f in sel)),
             'iva': s(iva, sum(n(f, 'i105') + n(f, 'i21') for f in sel))}
        for k in cuadro:
            cuadro[k] += v[k]
        detalle.append({'fila': et, **{k: round(x, 2) for k, x in v.items()}})
    dif = {k: round(cuadro[k] - hoja[k], 2) for k in hoja}
    # Mismo control que J / AG de la hoja, calculado aca para informarlo.
    suma_m_ae = sum(f[c] or 0 for f in datos for c in range(12, 31) if isinstance(f[c], (int, float)))
    percepciones = sum((f[14] or 0) + (f[15] or 0) for f in datos)
    control = {'suma_M_a_AE': round(suma_m_ae, 2), 'total_cuadro_AG': round(sum(cuadro.values()), 2),
               'diferencia': round(sum(cuadro.values()) - suma_m_ae, 2), 'percepciones_O_P': round(percepciones, 2)}

    wb.save(papel)
    print(json.dumps({
        'ok': all(abs(v) < 0.01 for v in dif.values()), 'codigos': codigos,
        'no_reconocidos': [c for c in codigos if c not in DESC],
        'casos_especiales': {'factura_b_por_alicuota': por_alicuota, 'nc_z_combinada': nc_z_combinada},
        'filas_cuadro': f'AC{cab}:AF{ult_cuadro}', 'fila_total': tot, 'detalle': detalle,
        'total_cuadro': {k: round(v, 2) for k, v in cuadro.items()},
        'total_planilla': {k: round(v, 2) for k, v in hoja.items()}, 'diferencia': dif,
        'control_J_AG': {**control, 'celdas': f'J{fila_tot + 1} y AG{tot + 1}'},
    }, ensure_ascii=False))


if __name__ == '__main__':
    main()
