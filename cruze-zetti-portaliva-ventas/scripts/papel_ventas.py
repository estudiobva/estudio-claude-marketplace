"""Vuelca las ventas del Portal IVA y de Zetti en el papel de trabajo del mes y
le agrega el cruce.

    python3 papel_ventas.py --zetti=<ZETTI - VENTAS.xlsx> --portal=<PORTAL IVA - VENTAS.xlsx>
                            --papel=<AAAAMM-papeldetrabajo-X.xlsx> --cuit=<CUIT> --periodo=AAAAMM
                            [--si-existe=abortar|reemplazar] [--respaldos=<carpeta>]

Hojas que escribe:
  "Ventas Portal IVA"  el Libro IVA Ventas del Portal tal como lo da ARCA en el CSV: las
                       mismas columnas (todas) en el mismo orden, fila 1 encabezados y
                       una fila por comprobante, solo que pasado a Excel con tipos reales
                       (fechas, numeros). Sale de AAAAMM - PORTAL IVA - VENTAS.xlsx, que es
                       la conversion 1 a 1 del CSV (csv_a_excel.py); no se copian la fila
                       TOTAL ni la nota al pie que ese archivo agrega.
                       A la derecha (desde la columna AI) el cuadro LIQUIDACION DE IVA -
                       Ventas: por tipo de comprobante Neto Gravado, No Gravado, Exento,
                       IVA (suma de "Importe IVA x%") y Total, con formulas SUMIFS sobre
                       las columnas del CSV; subtotal debito, subtotal NC (en negativo,
                       como en el CSV), ventas netas = debito + NC, control contra el
                       total del libro y prorrateo gravado/exento.
  "Ventas Zetti"       la hoja de AAAAMM - ZETTI - VENTAS.xlsx tal cual (ya sale con el
                       formato del papel).
  "cruze"              el cruce de cruzar_ventas.py (modo dos archivos: separa las NC B
                       del controlador fiscal) y, si las hay, "NC controlador (fuera)".

Si el papel existe, se respalda y se reescriben solo esas hojas (el resto queda
igual). Si alguna ya tiene datos, corta salvo --si-existe=reemplazar: puede
tener trabajo del equipo. Si el papel no existe, se crea con esas hojas.

Imprime un JSON con lo que hizo y el resumen del cruce.
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))

def fallar(msg, **extra):
    print(json.dumps({'ok': False, 'error': msg, **extra}, ensure_ascii=False, default=str))
    sys.exit(1)


def leer(ruta, **kw):
    # Drive a veces corta la primera lectura de un archivo solo-online.
    ultimo = None
    for _ in range(4):
        try:
            return openpyxl.load_workbook(ruta, **kw)
        except (OSError, TimeoutError) as e:
            ultimo = e
    fallar(f'No pude abrir {ruta}: {ultimo}')


def hoja(wb, *palabras):
    """Hoja cuyo nombre (sin mayusculas) contiene todas las palabras."""
    for n in wb.sheetnames:
        if all(p in n.lower() for p in palabras):
            return n
    return None


def filas_portal(ruta):
    """Hoja del Libro IVA Ventas (xlsx de portal-iva-descarga): encabezado y filas
    de comprobantes, sin la fila TOTAL ni la nota al pie. Devuelve la hoja de
    origen y la cantidad de filas a copiar (encabezado incluido)."""
    wb = leer(ruta)
    ws = wb[hoja(wb, 'portal') or wb.sheetnames[0]]
    if str(ws.cell(1, 1).value or '').strip() != 'Fecha de Emisión':
        fallar('El Libro IVA Ventas del Portal no empieza con el encabezado del CSV de ARCA.', archivo=ruta)
    n = 1
    while isinstance(ws.cell(n + 1, 1).value, datetime):
        n += 1
    return ws, n


# Filas del cuadro de liquidacion (modelo del papel de trabajo). Los tipos que
# aparezcan en el libro y no esten aca se agregan al final de su bloque.
DEBITO = [(1, '1 - Factura A'), (2, '2 - Nota de Débito A'), (7, '7 - Nota de Débito B'),
          (6, '6 - Factura B'), (81, '81 - Ticket Z'), (82, '82 - Tique Factura B'), (83, '83 - Tique')]
CREDITO = [(3, '3 - Nota de Crédito A'), (8, '8 - Nota de Crédito B'), (110, '110 - Nota de Crédito Z')]
NC_CODIGOS = {3, 8, 13, 21, 53, 110, 112, 113, 114, 119, 203, 208, 213}
OTROS_TIPOS = {4: 'Recibo A', 9: 'Recibo B', 11: 'Factura C', 12: 'Nota de Débito C', 13: 'Nota de Crédito C',
               19: 'Factura de Exportación E', 21: 'Nota de Crédito E', 51: 'Factura M', 52: 'Nota de Débito M',
               53: 'Nota de Crédito M', 111: 'Tique Factura C', 112: 'Tique Nota de Crédito A',
               113: 'Tique Nota de Crédito B', 114: 'Tique Nota de Crédito C', 115: 'Tique Nota de Débito A',
               116: 'Tique Nota de Débito B', 201: 'Factura de Crédito MiPyMEs A', 203: 'NC MiPyMEs A',
               206: 'Factura de Crédito MiPyMEs B', 208: 'NC MiPyMEs B'}
AZUL, CELESTE = 'FF1F4E79', 'FFDDEBF7'
NUM = '#,##0.00'
PCT = '0.00%'


def cuadro_liquidacion(ws, ultima, col0=35):
    """Cuadro LIQUIDACION DE IVA - Ventas a la derecha del CSV (col0=35 -> AI)."""
    tit = {str(c.value or '').strip(): c.column for c in ws[1]}
    L = lambda nombre: get_column_letter(tit[nombre])
    rango = lambda letra: f'${letra}$2:${letra}${ultima}'
    tipo = rango(L('Tipo de Comprobante'))
    ivas = [get_column_letter(c) for n, c in tit.items() if n.startswith('Importe IVA')]
    presentes = set()
    for r in range(2, ultima + 1):
        try:
            presentes.add(int(ws.cell(r, tit['Tipo de Comprobante']).value))
        except (TypeError, ValueError):
            pass
    nombre = lambda k: f'{k} - {OTROS_TIPOS.get(k, "Comprobante")}'
    debito = DEBITO + [(k, nombre(k)) for k in sorted(presentes) if k not in dict(DEBITO) and k not in NC_CODIGOS]
    credito = CREDITO + [(k, nombre(k)) for k in sorted(presentes) if k not in dict(CREDITO) and k in NC_CODIGOS]

    C = [get_column_letter(col0 + i) for i in range(6)]          # AI..AN
    fino, grueso = Side(style='thin'), Side(style='medium')
    negrita = Font(bold=True)

    def fila(r, valores, bold=False, arriba=None, abajo=None, fmt=NUM):
        for i, v in enumerate(valores):
            c = ws[f'{C[i]}{r}']
            c.value = v
            if i:
                c.number_format = fmt
            if bold:
                c.font = negrita
            if arriba or abajo:
                c.border = Border(top=arriba, bottom=abajo)

    def suma(col, k):
        return f'SUMIFS({rango(col)},{tipo},{k})'

    def por_tipo(k):
        iva = '+'.join(suma(x, k) for x in ivas) or '0'
        return [f'={suma(L("Total Neto Gravado"), k)}', f'={suma(L("Importe No Gravado"), k)}',
                f'={suma(L("Importe Exento"), k)}', f'={iva}', f'={suma(L("Importe Total"), k)}']

    r = 1
    ws[f'{C[0]}{r}'] = 'LIQUIDACIÓN DE IVA'
    for x in C:
        ws[f'{x}{r}'].fill = PatternFill('solid', fgColor=AZUL)
        ws[f'{x}{r}'].font = Font(bold=True, color='FFFFFFFF', size=12)
    r = 2
    ws[f'{C[0]}{r}'] = 'Ventas'
    for x in C:
        ws[f'{x}{r}'].fill = PatternFill('solid', fgColor=CELESTE)
        ws[f'{x}{r}'].font = Font(bold=True, color=AZUL)
    r = 3
    fila(r, ['Tipo de Comprobante', 'Neto Gravado', 'No Gravado', 'Exento', 'IVA', 'Total'], bold=True, abajo=grueso)
    for i in range(1, 6):
        ws[f'{C[i]}{r}'].alignment = Alignment(horizontal='center')

    r = 4
    ini = r
    for k, etiqueta in debito:
        fila(r, [etiqueta] + por_tipo(k))
        r += 1
    sub_deb = r
    fila(r, ['Subtotal Débito Fiscal'] + [f'=SUM({x}{ini}:{x}{r - 1})' for x in C[1:]], bold=True, arriba=fino)
    r += 1
    ini = r
    for k, etiqueta in credito:
        fila(r, [etiqueta] + por_tipo(k))
        r += 1
    sub_nc = r
    fila(r, ['Subtotal Notas de Crédito s/Ventas'] + [f'=SUM({x}{ini}:{x}{r - 1})' for x in C[1:]], bold=True, arriba=fino)
    r += 1
    netas = r
    # Las NC vienen en negativo (como en el CSV): ventas netas = debito + NC.
    fila(r, ['Ventas Netas (s/ Notas de Crédito)'] + [f'={x}{sub_deb}+{x}{sub_nc}' for x in C[1:]],
         bold=True, arriba=fino, abajo=fino)
    r += 1
    fila(r, ['Control: total del libro', None, None, None, None, f'=SUM({rango(L("Importe Total"))})'])
    ws[f'{C[0]}{r}'].font = Font(italic=True)
    r += 1
    fila(r, ['Diferencia (debe ser 0)', None, None, None, None, f'={C[5]}{netas}-{C[5]}{r - 1}'])
    ws[f'{C[0]}{r}'].font = Font(italic=True)

    r += 2
    ws[f'{C[0]}{r}'] = 'Prorrateo para Crédito Fiscal (Compras de Servicios)'
    for x in C[:2]:
        ws[f'{x}{r}'].fill = PatternFill('solid', fgColor=CELESTE)
        ws[f'{x}{r}'].font = Font(bold=True, color=AZUL)
    r += 1
    fila(r, ['Concepto', '%'], bold=True, abajo=fino)
    n, ng, ex = (f'{C[i]}{netas}' for i in (1, 2, 3))
    base = f'({n}+{ng}+{ex})'
    fila(r + 1, ['Venta Gravada', f'=IF({base}=0,0,{n}/{base})'], fmt=PCT)
    fila(r + 2, ['Venta Exenta', f'=IF({base}=0,0,({ng}+{ex})/{base})'], fmt=PCT)

    ws.column_dimensions[C[0]].width = 36
    for x in C[1:]:
        ws.column_dimensions[x].width = 16
    return {'tipos_debito': [k for k, _ in debito], 'tipos_nc': [k for k, _ in credito]}


def limpiar(ws):
    for rango in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rango))
    ws.delete_rows(1, ws.max_row + 1)


def tiene_datos(ws):
    return any(c.value not in (None, '') for row in ws.iter_rows(max_row=min(ws.max_row, 10)) for c in row)


def copiar_hoja(origen, destino, max_row=None):
    """Copia valores, formatos, combinadas y anchos de una hoja a otra (de otro libro)."""
    for row in origen.iter_rows(max_row=max_row):
        for c in row:
            if c.value is None and not c.has_style:
                continue
            d = destino.cell(c.row, c.column, c.value)
            if c.has_style:
                d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                d.alignment, d.number_format = copy(c.alignment), c.number_format
    for rango in origen.merged_cells.ranges:
        if max_row is None or rango.max_row <= max_row:
            destino.merge_cells(str(rango))
    for k, v in origen.column_dimensions.items():
        destino.column_dimensions[k].width = v.width
    for k, v in origen.row_dimensions.items():
        if v.height and (max_row is None or k <= max_row):
            destino.row_dimensions[k].height = v.height
    destino.freeze_panes = origen.freeze_panes
    if origen.auto_filter.ref and max_row is not None:
        destino.auto_filter.ref = f'A1:{get_column_letter(origen.max_column)}{max_row}'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--zetti', required=True)
    ap.add_argument('--portal', required=True)
    ap.add_argument('--papel', required=True)
    ap.add_argument('--cuit', required=True)
    ap.add_argument('--periodo', required=True, help='AAAAMM')
    ap.add_argument('--si-existe', dest='si_existe', choices=['abortar', 'reemplazar'], default='abortar')
    ap.add_argument('--respaldos', default=os.path.join(os.path.expanduser('~'), 'Documents', 'BVA-salidas', '_respaldos', 'papeles'))
    a = ap.parse_args()

    # 1) Portal IVA: el CSV de ARCA pasado a Excel, todas las columnas
    ws_p, filas_p = filas_portal(a.portal)
    if filas_p < 2:
        fallar('El Libro IVA Ventas del Portal no tiene comprobantes.', archivo=a.portal)

    # 2) Zetti: la hoja ya viene con el formato del papel
    wb_z = leer(a.zetti)
    nz = hoja(wb_z, 'zetti')
    if not nz:
        fallar('El archivo de Zetti no tiene la hoja "Ventas Zetti".', archivo=a.zetti)
    ws_z = wb_z[nz]

    # 3) Cruce (modo dos archivos de cruzar_ventas.py) a un temporal
    tmp = tempfile.mkdtemp(prefix='bva-cruze-')
    try:
        salida_cruce = os.path.join(tmp, 'cruce.xlsx')
        p = subprocess.run([sys.executable, os.path.join(AQUI, 'cruzar_ventas.py'), f'--zetti={a.zetti}',
                            f'--portal={a.portal}', f'--salida={salida_cruce}'],
                           capture_output=True, text=True)
        try:
            cruce = json.loads(p.stdout)
        except ValueError:
            fallar('El cruce no devolvio un resultado legible.', stdout=p.stdout[-800:], stderr=p.stderr[-800:])
        if p.returncode != 0 or 'error' in cruce:
            fallar(f'El cruce fallo: {cruce.get("error", p.stderr[-300:])}', detalle=cruce)
        wb_c = leer(salida_cruce)

        # 4) Papel de trabajo
        nuevo = not os.path.exists(a.papel)
        respaldo = None
        if nuevo:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
        else:
            if not a.papel.lower().endswith('.xlsx'):
                fallar('El papel de trabajo no es .xlsx: no se puede escribir.', papel=a.papel)
            wb = leer(a.papel)

        destinos = [
            ('Ventas Zetti', hoja(wb, 'ventas', 'zetti')),
            ('Ventas Portal IVA', hoja(wb, 'ventas', 'portal')),
            # Solo la hoja que se llama exactamente "cruze": las del equipo ("cruze vtas
            # (control)") quedan como estan.
            ('cruze', next((n for n in wb.sheetnames if n.strip().lower() == 'cruze'), None)),
        ]
        hojas_cruce = [n for n in wb_c.sheetnames]
        for extra in hojas_cruce[1:]:
            destinos.append((extra, extra if extra in wb.sheetnames else None))

        con_datos = [n for _, n in destinos if n and tiene_datos(wb[n])]
        if con_datos and a.si_existe != 'reemplazar':
            fallar('El papel ya tiene datos en ' + ', '.join(f'"{n}"' for n in con_datos) +
                   '. No los piso: --si-existe=reemplazar para rehacerlas.', papel=a.papel)

        if not nuevo:
            os.makedirs(a.respaldos, exist_ok=True)
            respaldo = os.path.join(a.respaldos, datetime.now().strftime('%Y%m%d-%H%M%S-') + os.path.basename(a.papel))
            shutil.copy2(a.papel, respaldo)

        def preparar(nombre, existente):
            if existente:
                ws = wb[existente]
                limpiar(ws)
                return ws
            return wb.create_sheet(nombre)

        ws = preparar('Ventas Zetti', destinos[0][1])
        copiar_hoja(ws_z, ws)
        ws = preparar('Ventas Portal IVA', destinos[1][1])
        copiar_hoja(ws_p, ws, max_row=filas_p)
        cuadro_liquidacion(ws, filas_p)
        ws = preparar('cruze', destinos[2][1])
        copiar_hoja(wb_c[hojas_cruce[0]], ws)
        for nombre, existente in destinos[3:]:
            copiar_hoja(wb_c[nombre], preparar(nombre, existente))
        # Orden del modelo: Ventas Zetti, Ventas Portal IVA, cruze al principio.
        if nuevo:
            wb._sheets.sort(key=lambda s: ['Ventas Zetti', 'Ventas Portal IVA', 'cruze'].index(s.title)
                            if s.title in ('Ventas Zetti', 'Ventas Portal IVA', 'cruze') else 9)
        wb.save(a.papel)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    chequeo = cruce.get('chequeo_vs_totales_de_los_archivos', {})
    print(json.dumps({
        'ok': True, 'papel': a.papel, 'papel_nuevo': nuevo, 'respaldo': respaldo,
        'filas_portal': filas_p - 1, 'filas_zetti': max(ws_z.max_row - 2, 0),
        'hojas': [n for n in wb.sheetnames],
        'cruce': {
            'filas_diferencia': cruce.get('cantidad_filas_diferencia'),
            'total_zetti': chequeo.get('total_zetti', {}).get('total'),
            'total_portal': chequeo.get('total_portal', {}).get('total'),
            'diferencia': chequeo.get('diferencia_entre_archivos', {}).get('total'),
            'nc_controlador_fuera': cruce.get('nc_controlador_fuera_del_cruce'),
            'chequeo_coincide': chequeo.get('coincide'),
            'diferencias': cruce.get('diferencias'),
        },
    }, ensure_ascii=False, default=str))


if __name__ == '__main__':
    main()
