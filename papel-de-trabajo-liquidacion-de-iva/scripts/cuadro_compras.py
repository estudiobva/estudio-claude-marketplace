"""Etapa 2 (DDJJ de IVA), paso 7 parte 2: cuadro resumen de "Compras portal iva".

    python cuadro_compras.py --papel=<papel de trabajo.xlsx>

Requiere la columna AG (PROVEEDORES) completa (parte 1, clasificar_proveedores.py). Si hay
filas sin clasificar o NO CLASIFICADO, corta sin grabar.

Abajo de los datos (encabezado 5 filas debajo de la nota) arma, con SUMIFS vivas sobre las filas
de comprobantes, con el formato de Norte 202609:
  * Cuadro por tipo de comprobante y MERCADERIA (AG = "CMV") / SERVICIOS (todo lo demas): 16 filas
    fijas en el orden de la usuaria, aunque queden en cero, mas las combinaciones que aparezcan y
    no esten en la lista (se avisan). Columnas: Neto Gravado 0% (T), 10,5 (Y), 21 (AA), 27 (AC),
    No Gravado (K), Exento (L), IVA 10,5 (Z), IVA 21 (AB), IVA 27 (AD) y Total. En los tipos sin IVA
    discriminado (6, 8, 11, 13, 15, 111) No Gravado suma ademas, con SUMPRODUCT, Importe Total (H) - N a S.
    Los tipos cargados como texto ("FC") se suman en "1 - Factura A / SERVICIOS", sin tocar B.
    Una fila libre y TOTAL PORTAL IVA (en AE, el total general).
  * Bloque aparte: N a S (Otros Imp. Nac., Perc. IIBB, Municipales, Perc. IVA, Internos, Otros
    Tributos) por MERCADERIA / SERVICIOS / TOTAL.
  * Control: cuadro + otros conceptos contra el Importe Total (H), con la diferencia.

Columnas del cuadro: U (tipo), V a AD (conceptos), AE (grupo MERCADERIA / SERVICIOS; en la fila
TOTAL, el total general). Otros conceptos en U a X y control con los importes en X. Letra Aptos
Narrow 11, encabezado en negrita con linea fina arriba y abajo, formato contable con negativos en
rojo, tipo y grupo de las NC (3, 8 y 13) en rojo. Las formulas se verifican calculando lo mismo en Python (no hay
Excel/LibreOffice para recalcular): imprime los totales y la diferencia del control.
"""
import json
import sys
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import column_index_from_string as ci

from fuente import ZETTI

HOJA = 'Compras portal iva'
FMT = '_-* #,##0.00_-;[Red]\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
FINA = Side(style='thin')
GRUESA = Side(style='medium')
# Retoques de la usuaria sobre Norte 202610 (2026-11-16): encabezado U a AD con relleno gris (blanco del tema
# oscurecido 15 %), recuadro grueso del encabezado al TOTAL (U a AE) y TOTAL con linea gruesa abajo.
GRIS = PatternFill('solid', fgColor=Color(theme=0, tint=-0.1499984740745262))
F = dict(name=ZETTI[0], size=ZETTI[1])

DESC = {1: 'Factura A', 2: 'Nota de Débito A', 3: 'Nota de Crédito A', 4: 'Recibo A', 6: 'Factura B',
        7: 'Nota de Débito B', 8: 'Nota de Crédito B', 9: 'Recibo B', 11: 'Factura C', 12: 'Nota de Débito C',
        13: 'Nota de Crédito C', 15: 'Recibo C', 19: 'Factura por operaciones con el exterior', 20: 'Nota de Débito por operaciones con el exterior',
        21: 'Nota de Crédito por operaciones con el exterior',
        51: 'Factura M', 52: 'Nota de Débito M', 53: 'Nota de Crédito M',
        60: 'Cta. de Vta. y Liq. A', 63: 'Liquidaciones A', 64: 'Liquidaciones B', 81: 'Tique Factura A',
        82: 'Tique Factura B', 83: 'Tique', 111: 'Tique Factura C', 112: 'Tique NC A', 113: 'Tique NC B',
        114: 'Tique NC C', 201: 'FCE MiPyME A', 203: 'NC MiPyME A', 206: 'FCE MiPyME B', 211: 'FCE MiPyME C'}
FIJAS = [(1, 'SERVICIOS'), (1, 'MERCADERIA'), (63, 'SERVICIOS'), (4, 'SERVICIOS'), (81, 'SERVICIOS'),
         (81, 'MERCADERIA'), (2, 'SERVICIOS'), (6, 'SERVICIOS'), (11, 'SERVICIOS'),
         (11, 'MERCADERIA'), (15, 'SERVICIOS'),
         (111, 'SERVICIOS'), (3, 'MERCADERIA'), (3, 'SERVICIOS'), (8, 'MERCADERIA'), (13, 'SERVICIOS')]
NC = {3, 8, 13, 21, 53, 112, 113, 114, 203}
SIN_IVA = {6, 8, 11, 13, 15, 111}
CONCEPTOS = [('Neto Gravado 0%', 'T'), ('Neto Gravado 10,5', 'Y'), ('Neto Gravado 21', 'AA'),
             ('Neto Gravado 27', 'AC'), ('No Gravado', 'K'), ('Exento', 'L'), ('IVA 10,5', 'Z'),
             ('IVA 21', 'AB'), ('IVA 27', 'AD')]
OTROS = [('Per. o Pagos a Cta. de Otros Imp. Nac.', 'N'), ('Percepciones de Ingresos Brutos', 'O'),
         ('Impuestos Municipales', 'P'), ('Percepciones o Pagos a Cuenta de IVA', 'Q'),
         ('Impuestos Internos', 'R'), ('Otros Tributos', 'S')]
NS = [c for _, c in OTROS]


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

    # --- limpiar lo que haya debajo de la nota (re-ejecucion) y ubicar el cuadro.
    # Formato de Norte 202609 (pedido 2026-11-16): cuadro al final de las columnas, tipo en U, los nueve
    # conceptos en V a AD y el grupo (MERCADERIA / SERVICIOS) en AE; encabezado 5 filas debajo de la nota.
    nota = next((r for r in range(u + 1, ws.max_row + 1)
                 if isinstance(ws.cell(r, 1).value, str) and 'comprobantes' in ws.cell(r, 1).value), None)
    fila_tot_datos = next((r for r in range(u + 1, ws.max_row + 1) if ws.cell(r, 1).value == 'TOTAL'), None)
    base = (nota or fila_tot_datos or u) + 1
    for r in range(base, ws.max_row + 1):
        for c in range(1, ci('AH') + 1):
            cel = ws.cell(r, c)
            cel.value = None
            cel.style = 'Normal'
    cab = base + 4

    def rng(c):
        return f'${c}$2:${c}${u}'

    def grupo_crit(grupo):
        return '"CMV"' if grupo == 'MERCADERIA' else '"<>CMV"'

    def sumifs(col, tipo, grupo):
        f = f'SUMIFS({rng(col)},{rng("B")},{tipo},{rng("AG")},{grupo_crit(grupo)})'
        if (tipo, grupo) == (1, 'SERVICIOS') and texto:
            # tipos cargados como texto ("FC", servicios publicos) van en 1 - Factura A / SERVICIOS
            f += ''.join(f'+SUMIFS({rng(col)},{rng("B")},"{t}",{rng("AG")},"<>CMV")' for t in texto)
        return f

    def sumproduct_sin_iva(tipo, grupo):
        cond = '="CMV"' if grupo == 'MERCADERIA' else '<>"CMV"'
        resta = ''.join(f'-{rng(c)}' for c in NS)
        return f'SUMPRODUCT(({rng("B")}={tipo})*({rng("AG")}{cond})*({rng("H")}{resta}))'

    def valor(col, tipo, grupo):
        tipos = {tipo, *texto} if (tipo, grupo) == (1, 'SERVICIOS') else {tipo}
        return sum(d[col] for d in datos if d['tipo'] in tipos and d['grupo'] == grupo)

    def estilo(c, bold=False, red=False, fmt=None):
        c.font = Font(bold=bold, color='FFFF0000' if red else None, **F)
        if fmt:
            c.number_format = fmt

    U, V, AD, AE = ci('U'), ci('V'), ci('AD'), ci('AE')

    # --- cuadro: encabezado (U y AE vacios), filas por tipo/grupo, una fila libre y TOTAL PORTAL IVA
    for j in range(U, AE + 1):
        c = ws.cell(cab, j, CONCEPTOS[j - V][0] if V <= j <= AD else None)
        estilo(c, bold=V <= j <= AD)
        if V <= j <= AD:
            c.alignment = Alignment(horizontal='center', wrap_text=True)
        if U <= j <= AD:
            c.fill = GRIS
        c.border = Border(left=GRUESA if j == U else None, right=GRUESA if j == AE else None,
                          top=GRUESA, bottom=FINA)
    ws.row_dimensions[cab].height = 28.8
    calc = {}
    r = cab
    for tipo, grupo in filas_cuadro:
        r += 1
        es_nc = tipo in NC
        estilo(ws.cell(r, U, f'{tipo} - {DESC.get(tipo, f"Tipo {tipo}")}'), red=es_nc)
        fila = {}
        for k, (t, col) in enumerate(CONCEPTOS):
            f = sumifs(col, tipo, grupo)
            v = valor(col, tipo, grupo)
            if col == 'K' and tipo in SIN_IVA:
                f += '+' + sumproduct_sin_iva(tipo, grupo)
                v += valor('H', tipo, grupo) - sum(valor(c, tipo, grupo) for c in NS)
            estilo(ws.cell(r, V + k, '=' + f), fmt=FMT)
            fila[t] = v
        estilo(ws.cell(r, AE, grupo), red=es_nc)
        fila['TOTAL'] = sum(fila.values())
        calc[(tipo, grupo)] = fila
    d1, d2 = cab + 1, r
    r += 2
    fila_total_cuadro = r
    estilo(ws.cell(r, U, 'TOTAL PORTAL IVA'), bold=True)
    for j in range(V, AD + 1):
        L = letra_col(j)
        estilo(ws.cell(r, j, f'=SUM({L}{d1}:{L}{d2})'), bold=True, fmt=FMT)
    estilo(ws.cell(r, AE, f'=SUM(V{r}:AD{r})'), bold=True, fmt=FMT)
    for j in range(U, AE + 1):
        ws.cell(r, j).border = Border(left=GRUESA if j == U else None, right=GRUESA if j == AE else None,
                                      top=FINA, bottom=GRUESA)
    for rr in range(cab + 1, r):  # costados del recuadro en las filas del cuadro y la fila libre
        ws.cell(rr, U).border = Border(left=GRUESA)
        ws.cell(rr, AE).border = Border(right=GRUESA)
    tot_cuadro = {t: sum(f[t] for f in calc.values()) for t in [t for t, _ in CONCEPTOS] + ['TOTAL']}

    # --- bloque aparte: N a S (U = concepto, V MERCADERIA, W SERVICIOS, X TOTAL)
    r += 3
    estilo(ws.cell(r, U, 'OTROS CONCEPTOS (columnas N a S) - no integran el cuadro'), bold=True)
    r += 1
    for j, t in enumerate([None, 'MERCADERIA', 'SERVICIOS', 'TOTAL'], start=U):
        c = ws.cell(r, j, t)
        estilo(c, bold=t is not None)
        c.border = Border(top=FINA, bottom=FINA)
    o1 = r + 1
    otros = {}
    for col in NS:
        r += 1
        titulo = str(ws[f'{col}1'].value)
        estilo(ws.cell(r, U, titulo))
        estilo(ws.cell(r, U + 1, f'=SUMIF({rng("AG")},"CMV",{rng(col)})'), fmt=FMT)
        estilo(ws.cell(r, U + 2, f'=SUMIF({rng("AG")},"<>CMV",{rng(col)})'), fmt=FMT)
        estilo(ws.cell(r, U + 3, f'=SUM(V{r}:W{r})'), fmt=FMT)
        otros[titulo] = {g: sum(d[col] for d in datos if d['grupo'] == g) for g in ('MERCADERIA', 'SERVICIOS')}
    o2 = r
    r += 1
    fila_total_otros = r
    estilo(ws.cell(r, U, 'TOTAL OTROS CONCEPTOS'), bold=True)
    for j in range(U + 1, U + 4):
        L = letra_col(j)
        estilo(ws.cell(r, j, f'=SUM({L}{o1}:{L}{o2})'), bold=True, fmt=FMT)
    for j in range(U, U + 4):
        ws.cell(r, j).border = Border(top=FINA)
    tot_otros = sum(v['MERCADERIA'] + v['SERVICIOS'] for v in otros.values())

    # --- control (valores en X)
    r += 2
    estilo(ws.cell(r, U, 'CONTROL: cuadro + otros conceptos vs. Importe Total'), bold=True)
    total_h = sum(d['H'] for d in datos)
    r += 1
    estilo(ws.cell(r, U, 'Cuadro (TOTAL PORTAL IVA)'))
    estilo(ws.cell(r, U + 3, f'=+AE{fila_total_cuadro}'), fmt=FMT)
    estilo(ws.cell(r + 1, U, 'Otros conceptos (N a S)'))
    estilo(ws.cell(r + 1, U + 3, f'=+X{fila_total_otros}'), fmt=FMT)
    estilo(ws.cell(r + 2, U, 'Importe Total (columna H)'))
    estilo(ws.cell(r + 2, U + 3, f'=+H{fila_tot_datos}' if fila_tot_datos else f'=SUM({rng("H")})'), fmt=FMT)
    estilo(ws.cell(r + 3, U, 'Diferencia'), bold=True)
    c = ws.cell(r + 3, U + 3, f'=+X{r}+X{r + 1}-X{r + 2}')
    estilo(c, bold=True, fmt=FMT)
    c.border = Border(top=FINA)
    dif = round(tot_cuadro['TOTAL'] + tot_otros - total_h, 2)
    ini = cab

    for L, w in (('U', 42), ('V', 14.33), ('W', 13.66), ('X', 17), ('Y', 20.11), ('Z', 20), ('AA', 16.44),
                 ('AB', 13.22), ('AC', 14.33), ('AD', 14), ('AE', 17.66)):
        ws.column_dimensions[L].width = w
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
