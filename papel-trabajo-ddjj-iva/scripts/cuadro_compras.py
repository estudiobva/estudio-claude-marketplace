"""Etapa 2 (DDJJ de IVA), paso 7 parte 2: cuadro resumen de "Compras portal iva".

    python cuadro_compras.py --papel=<papel de trabajo.xlsx>

Requiere la columna AG (PROVEEDORES) completa (parte 1, clasificar_proveedores.py). Si hay
filas sin clasificar o NO CLASIFICADO, corta sin grabar.

Abajo de los datos (dos filas libres despues de la nota) arma, con SUMIFS vivas sobre las filas
de comprobantes:
  * Cuadro por tipo de comprobante y MERCADERIA (AG = "CMV") / SERVICIOS (todo lo demas): 15 filas
    fijas en el orden de la usuaria, aunque queden en cero, mas las combinaciones que aparezcan y
    no esten en la lista (se avisan). Columnas: Neto Gravado 0% (T), 10,5 (Y), 21 (AA), 27 (AC),
    No Gravado (K), Exento (L), IVA 10,5 (Z), IVA 21 (AB), IVA 27 (AD) y Total. En los tipos sin IVA
    discriminado (6, 8, 11, 13, 15, 111) No Gravado suma ademas Importe Total (H) - N a S. Los tipos
    cargados como texto ("FC") van en "1 - Factura A / SERVICIOS" (criterio {1,"FC"}), sin tocar B.
    Fila TOTAL PORTAL IVA al pie.
  * Bloque aparte: N a S (Otros Imp. Nac., Perc. IIBB, Municipales, Perc. IVA, Internos, Otros
    Tributos) por MERCADERIA / SERVICIOS / TOTAL.
  * Control: cuadro + otros conceptos contra el Importe Total (H), con la diferencia.

Columnas del cuadro: G (tipo), H (grupo), I a Q (conceptos), R (total). Letra Aptos Narrow 11,
encabezado gris en negrita, recuadro grueso, formato contable con negativos en rojo, tipo y grupo
de las NC (3, 8 y 13) en rojo. Las formulas se verifican calculando lo mismo en Python (no hay
Excel/LibreOffice para recalcular): imprime los totales y la diferencia del control.
"""
import json
import sys
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string as ci

from fuente import ZETTI

HOJA = 'Compras portal iva'
FMT = '_-* #,##0.00_-;[Red]-* #,##0.00_-;_-* "-"??_-;_-@_-'
GRIS = PatternFill('solid', fgColor='D9D9D9')
GRUESA = Side(style='medium')
FINA = Side(style='thin')
F = dict(name=ZETTI[0], size=ZETTI[1])

DESC = {1: 'Factura A', 2: 'Nota de Débito A', 3: 'Nota de Crédito A', 4: 'Recibo A', 6: 'Factura B',
        7: 'Nota de Débito B', 8: 'Nota de Crédito B', 9: 'Recibo B', 11: 'Factura C', 12: 'Nota de Débito C',
        13: 'Nota de Crédito C', 15: 'Recibo C', 51: 'Factura M', 52: 'Nota de Débito M', 53: 'Nota de Crédito M',
        60: 'Cta. de Vta. y Liq. A', 63: 'Liquidaciones A', 64: 'Liquidaciones B', 81: 'Tique Factura A',
        82: 'Tique Factura B', 83: 'Tique', 111: 'Tique Factura C', 112: 'Tique NC A', 113: 'Tique NC B',
        114: 'Tique NC C', 201: 'FCE MiPyME A', 203: 'NC MiPyME A', 206: 'FCE MiPyME B', 211: 'FCE MiPyME C'}
FIJAS = [(1, 'SERVICIOS'), (1, 'MERCADERIA'), (63, 'SERVICIOS'), (4, 'SERVICIOS'), (81, 'SERVICIOS'),
         (81, 'MERCADERIA'), (2, 'SERVICIOS'), (6, 'SERVICIOS'), (11, 'SERVICIOS'), (15, 'SERVICIOS'),
         (111, 'SERVICIOS'), (3, 'MERCADERIA'), (3, 'SERVICIOS'), (8, 'MERCADERIA'), (13, 'SERVICIOS')]
NC = {3, 8, 13, 53, 112, 113, 114, 203}
SIN_IVA = {6, 8, 11, 13, 15, 111}
CONCEPTOS = [('Neto Gravado 0%', 'T'), ('Neto Gravado 10,5', 'Y'), ('Neto Gravado 21', 'AA'),
             ('Neto Gravado 27', 'AC'), ('No Gravado', 'K'), ('Exento', 'L'), ('IVA 10,5', 'Z'),
             ('IVA 21', 'AB'), ('IVA 27', 'AD')]
OTROS = [('Per. o Pagos a Cta. de Otros Imp. Nac.', 'N'), ('Percepciones de Ingresos Brutos', 'O'),
         ('Impuestos Municipales', 'P'), ('Percepciones o Pagos a Cuenta de IVA', 'Q'),
         ('Impuestos Internos', 'R'), ('Otros Tributos', 'S')]
NS = [c for _, c in OTROS]
C_TIPO, C_GRUPO, C0 = 7, 8, 9  # G, H, I..Q; total en R
C_TOT = C0 + len(CONCEPTOS)


def fallar(msg, **extra):
    sys.exit(json.dumps({'error': msg, **extra}, ensure_ascii=False))


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def letra_col(i):
    from openpyxl.utils import get_column_letter
    return get_column_letter(i)


def main():
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel = opts.get('papel')
    if not papel:
        sys.exit(__doc__)
    wb = load_workbook(papel)
    if HOJA not in wb.sheetnames:
        fallar(f'Falta la solapa "{HOJA}"', solapas=wb.sheetnames)
    ws = wb[HOJA]
    if str(ws['AG1'].value or '').strip().upper() != 'PROVEEDORES':
        fallar('AG1 no dice PROVEEDORES: correr antes el paso 6 y la parte 1 (clasificar_proveedores.py)')

    filas = [r for r in range(2, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, datetime)]
    if not filas or filas != list(range(2, filas[-1] + 1)):
        fallar('Los comprobantes no estan seguidos desde la fila 2', primera=filas[:1], ultima=filas[-1:])
    u = filas[-1]
    sin = [r for r in filas if str(ws.cell(r, ci('AG')).value or '').strip() in ('', 'NO CLASIFICADO')]
    if sin:
        fallar('Hay proveedores sin clasificar en AG: primero hay que clasificarlos', filas=sin[:30])

    # --- datos en Python para verificar las formulas
    datos = []
    for r in filas:
        g = lambda c: ws.cell(r, ci(c)).value  # noqa: E731
        datos.append({'tipo': g('B'), 'grupo': 'MERCADERIA' if str(g('AG')).strip() == 'CMV' else 'SERVICIOS',
                      **{c: num(g(c)) for c in ['H'] + [c for _, c in CONCEPTOS] + NS}})

    def clave(d):
        t = d['tipo']
        return (1 if isinstance(t, str) else t, d['grupo'])

    combos = sorted({clave(d) for d in datos}, key=lambda x: (x[0] in NC, x[0], x[1]))
    extras = [c for c in combos if c not in FIJAS]
    # extras antes de las NC, para que las NC queden al final como en la lista fija
    filas_cuadro = [c for c in FIJAS if c[0] not in NC] + [c for c in extras if c[0] not in NC] \
        + [c for c in FIJAS if c[0] in NC] + [c for c in extras if c[0] in NC]
    texto = sorted({d['tipo'] for d in datos if isinstance(d['tipo'], str)})

    # --- limpiar lo que haya debajo de la nota (re-ejecucion) y ubicar el cuadro
    nota = next((r for r in range(u + 1, ws.max_row + 1)
                 if isinstance(ws.cell(r, 1).value, str) and 'comprobantes' in ws.cell(r, 1).value), None)
    base = (nota or next((r for r in range(u + 1, ws.max_row + 1) if ws.cell(r, 1).value == 'TOTAL'), u)) + 1
    for r in range(base, ws.max_row + 1):
        for c in range(1, ci('AH') + 1):
            cel = ws.cell(r, c)
            cel.value = None
            cel.style = 'Normal'
    ini = base + 2

    def rng(c):
        return f'${c}$2:${c}${u}'

    def crit(tipo, grupo):
        g = '"CMV"' if grupo == 'MERCADERIA' else '"<>CMV"'
        t = '{1,"FC"}' if (tipo, grupo) == (1, 'SERVICIOS') else str(tipo)
        return f'{rng("B")},{t},{rng("AG")},{g}'

    def sumifs(col, tipo, grupo):
        f = f'SUMIFS({rng(col)},{crit(tipo, grupo)})'
        return f'SUM({f})' if (tipo, grupo) == (1, 'SERVICIOS') else f

    def valor(col, tipo, grupo):
        tipos = {tipo, 'FC'} if (tipo, grupo) == (1, 'SERVICIOS') else {tipo}
        return sum(d[col] for d in datos if d['tipo'] in tipos and d['grupo'] == grupo)

    def estilo(c, bold=False, red=False, fmt=None, fill=None):
        c.font = Font(bold=bold, color='FF0000' if red else None, **F)
        if fmt:
            c.number_format = fmt
        if fill:
            c.fill = fill

    def recuadro(r1, r2, c1, c2):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                b = ws.cell(r, c).border
                ws.cell(r, c).border = Border(left=GRUESA if c == c1 else b.left, right=GRUESA if c == c2 else b.right,
                                              top=GRUESA if r == r1 else b.top, bottom=GRUESA if r == r2 else b.bottom)

    # --- cuadro
    r = ini
    estilo(ws.cell(r, C_TIPO, 'CUADRO RESUMEN COMPRAS PORTAL IVA (por tipo de comprobante y MERCADERIA / SERVICIOS)'), bold=True)
    r += 1
    cab = r
    for j, t in enumerate(['Tipo de comprobante', 'Grupo'] + [t for t, _ in CONCEPTOS] + ['TOTAL'], start=C_TIPO):
        c = ws.cell(r, j, t)
        estilo(c, bold=True, fill=GRIS)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = Border(bottom=FINA)
    calc = {}
    for tipo, grupo in filas_cuadro:
        r += 1
        es_nc = tipo in NC
        estilo(ws.cell(r, C_TIPO, f'{tipo} - {DESC.get(tipo, f"Tipo {tipo}")}'), red=es_nc)
        estilo(ws.cell(r, C_GRUPO, grupo), red=es_nc)
        fila = {}
        for k, (t, col) in enumerate(CONCEPTOS):
            f = sumifs(col, tipo, grupo)
            v = valor(col, tipo, grupo)
            if col == 'K' and tipo in SIN_IVA:
                f += f'+{sumifs("H", tipo, grupo)}' + ''.join(f'-{sumifs(c, tipo, grupo)}' for c in NS)
                v += valor('H', tipo, grupo) - sum(valor(c, tipo, grupo) for c in NS)
            estilo(ws.cell(r, C0 + k, '=' + f), fmt=FMT)
            fila[t] = v
        L0, L1 = letra_col(C0), letra_col(C_TOT - 1)
        estilo(ws.cell(r, C_TOT, f'=SUM({L0}{r}:{L1}{r})'), bold=True, fmt=FMT)
        fila['TOTAL'] = sum(fila.values())
        calc[(tipo, grupo)] = fila
    d1, d2 = cab + 1, r
    r += 1
    fila_total_cuadro = r
    estilo(ws.cell(r, C_TIPO, 'TOTAL PORTAL IVA'), bold=True)
    for j in range(C0, C_TOT + 1):
        L = letra_col(j)
        c = ws.cell(r, j, f'=SUM({L}{d1}:{L}{d2})')
        estilo(c, bold=True, fmt=FMT)
        c.border = Border(top=FINA)
    recuadro(cab, r, C_TIPO, C_TOT)
    tot_cuadro = {t: sum(f[t] for f in calc.values()) for t in [t for t, _ in CONCEPTOS] + ['TOTAL']}

    # --- bloque aparte: N a S
    r += 3
    estilo(ws.cell(r, C_TIPO, 'OTROS CONCEPTOS (columnas N a S, fuera del cuadro)'), bold=True)
    r += 1
    cab2 = r
    for j, t in enumerate(['Concepto', 'MERCADERIA', 'SERVICIOS', 'TOTAL'], start=C_TIPO):
        c = ws.cell(r, j, t)
        estilo(c, bold=True, fill=GRIS)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    otros = {}
    for t, col in OTROS:
        r += 1
        estilo(ws.cell(r, C_TIPO, f'{t} ({col})'))
        estilo(ws.cell(r, C_TIPO + 1, f'=SUMIFS({rng(col)},{rng("AG")},"CMV")'), fmt=FMT)
        estilo(ws.cell(r, C_TIPO + 2, f'=SUMIFS({rng(col)},{rng("AG")},"<>CMV")'), fmt=FMT)
        estilo(ws.cell(r, C_TIPO + 3, f'=SUM({letra_col(C_TIPO + 1)}{r}:{letra_col(C_TIPO + 2)}{r})'), bold=True, fmt=FMT)
        otros[t] = {g: sum(d[col] for d in datos if d['grupo'] == g) for g in ('MERCADERIA', 'SERVICIOS')}
    o1, o2 = cab2 + 1, r
    r += 1
    fila_total_otros = r
    estilo(ws.cell(r, C_TIPO, 'TOTAL OTROS CONCEPTOS'), bold=True)
    for j in range(C_TIPO + 1, C_TIPO + 4):
        L = letra_col(j)
        c = ws.cell(r, j, f'=SUM({L}{o1}:{L}{o2})')
        estilo(c, bold=True, fmt=FMT)
        c.border = Border(top=FINA)
    recuadro(cab2, r, C_TIPO, C_TIPO + 3)
    tot_otros = sum(v['MERCADERIA'] + v['SERVICIOS'] for v in otros.values())

    # --- control
    r += 3
    estilo(ws.cell(r, C_TIPO, 'CONTROL'), bold=True)
    c_val = C_TIPO + 3  # misma columna que el TOTAL del bloque aparte
    total_h = sum(d['H'] for d in datos)
    renglones = [
        ('Total cuadro (TOTAL PORTAL IVA)', f'={letra_col(C_TOT)}{fila_total_cuadro}'),
        ('Otros conceptos (N a S)', f'={letra_col(C_TIPO + 3)}{fila_total_otros}'),
        ('Cuadro + otros conceptos', None),
        ('Importe Total Portal IVA (columna H)', f'=SUM({rng("H")})'),
        ('Diferencia', None),
    ]
    r0 = r + 1
    for i, (t, f) in enumerate(renglones):
        r += 1
        L = letra_col(c_val)
        if t == 'Cuadro + otros conceptos':
            f = f'={L}{r0}+{L}{r0 + 1}'
        if t == 'Diferencia':
            f = f'={L}{r0 + 2}-{L}{r0 + 3}'
        estilo(ws.cell(r, C_TIPO, t), bold=t in ('Cuadro + otros conceptos', 'Diferencia'))
        estilo(ws.cell(r, c_val, f), bold=t == 'Diferencia', fmt=FMT)
    recuadro(r0, r, C_TIPO, c_val)
    dif = round(tot_cuadro['TOTAL'] + tot_otros - total_h, 2)

    # anchos minimos para que se lean los importes (no cambia datos)
    for j in range(C0, C_TOT + 1):
        L = letra_col(j)
        if (ws.column_dimensions[L].width or 0) < 16:
            ws.column_dimensions[L].width = 16
    wb.save(papel)

    r2 = lambda x: round(x + 0.0, 2)  # noqa: E731
    print(json.dumps({
        'papel': papel, 'cuadro_desde_fila': ini, 'filas_de_datos': f'2 a {u}',
        'filas_cuadro': [{'fila': f'{t} - {DESC.get(t, t)} / {g}', **{k: r2(v) for k, v in calc[(t, g)].items() if abs(v) > 0.004}}
                         for t, g in filas_cuadro],
        'combinaciones_fuera_de_lista': [f'{t} - {DESC.get(t, t)} / {g}' for t, g in extras],
        'tipos_cargados_como_texto': texto,
        'total_cuadro': {k: r2(v) for k, v in tot_cuadro.items()},
        'otros_conceptos': {t: {g: r2(v) for g, v in x.items()} for t, x in otros.items()},
        'control': {'cuadro': r2(tot_cuadro['TOTAL']), 'otros': r2(tot_otros), 'importe_total_H': r2(total_h),
                    'diferencia': dif},
    }, ensure_ascii=False, indent=1))
    if abs(dif) >= 1:
        sys.exit(2)


if __name__ == '__main__':
    main()
