"""Vuelca las ventas del Portal IVA y de Zetti en el papel de trabajo del mes y
le agrega el cruce.

    python3 papel_ventas.py --zetti=<ZETTI - VENTAS.xlsx> --portal=<PORTAL IVA - VENTAS.xlsx>
                            --papel=<AAAAMM-papeldetrabajo-X.xlsx> --cuit=<CUIT> --periodo=AAAAMM
                            [--si-existe=abortar|reemplazar] [--respaldos=<carpeta>]

Hojas que escribe (formato compacto del papel de trabajo mensual):
  "Ventas Portal IVA"  A1 "Comprobantes de Ventas - CUIT <cuit> - Portal IVA ARCA - MM/AAAA",
                       fila 2 titulos, datos desde la fila 3: Fecha de Emision | Tipo de
                       Comprobante ("1 - Factura A") | Punto de Venta | Numero de Comprobante |
                       Importe No Gravado | Importe Exento | Total Neto Gravado | Total IVA |
                       Importe Total. NC en negativo, ordenado por tipo y fecha.
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
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))
CONTABLE = '_-"$"\\ * #,##0.00_-;\\-"$"\\ * #,##0.00_-;_-"$"\\ * "-"??_-;_-@_-'
CHICA = Font(size=10)

# Codigos de comprobante de ARCA, con el texto que usa el papel de trabajo.
TIPOS = {
    1: 'Factura A', 2: 'Nota de Débito A', 3: 'Nota de Crédito A', 4: 'Recibo A',
    5: 'Nota de Venta al contado A', 6: 'Factura B', 7: 'Nota de Débito B', 8: 'Nota de Crédito B',
    9: 'Recibo B', 10: 'Nota de Venta al contado B', 11: 'Factura C', 12: 'Nota de Débito C',
    13: 'Nota de Crédito C', 15: 'Recibo C', 19: 'Factura de Exportación E',
    20: 'Nota de Débito por Operaciones con el Exterior E',
    21: 'Nota de Crédito por Operaciones con el Exterior E', 51: 'Factura M',
    52: 'Nota de Débito M', 53: 'Nota de Crédito M', 81: 'Tique Factura A', 82: 'Tique Factura B',
    83: 'Tique', 109: 'Tique C', 110: 'Tique Nota de Crédito', 111: 'Tique Factura C',
    112: 'Tique Nota de Crédito A', 113: 'Tique Nota de Crédito B', 114: 'Tique Nota de Crédito C',
    115: 'Tique Nota de Débito A', 116: 'Tique Nota de Débito B', 117: 'Tique Nota de Débito C',
    118: 'Tique Factura M', 119: 'Tique Nota de Crédito M', 120: 'Tique Nota de Débito M',
    201: 'Factura de Crédito electrónica MiPyMEs (FCE) A',
    202: 'Nota de Débito electrónica MiPyMEs (FCE) A',
    203: 'Nota de Crédito electrónica MiPyMEs (FCE) A',
    206: 'Factura de Crédito electrónica MiPyMEs (FCE) B',
    207: 'Nota de Débito electrónica MiPyMEs (FCE) B',
    208: 'Nota de Crédito electrónica MiPyMEs (FCE) B',
    211: 'Factura de Crédito electrónica MiPyMEs (FCE) C',
    212: 'Nota de Débito electrónica MiPyMEs (FCE) C',
    213: 'Nota de Crédito electrónica MiPyMEs (FCE) C',
}
COLS_PORTAL = ['Fecha de Emisión', 'Tipo de Comprobante', 'Punto de Venta', 'Número de Comprobante',
               'Importe No Gravado', 'Importe Exento', 'Total Neto Gravado', 'Total IVA', 'Importe Total']
IMPORTES_PORTAL = COLS_PORTAL[4:]


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


def tipo_texto(v):
    if isinstance(v, str) and ' - ' in v:
        return v
    try:
        n = int(v)
    except (TypeError, ValueError):
        return v
    return f'{n} - {TIPOS.get(n, "Comprobante")}'


def filas_portal(ruta):
    """Filas del Libro IVA Ventas del Portal (xlsx de portal-iva-descarga) en el
    orden y las columnas de la hoja del papel."""
    wb = leer(ruta, data_only=True)
    n = hoja(wb, 'portal') or wb.sheetnames[0]
    ws = wb[n]
    filas = list(ws.iter_rows(values_only=True))
    tit = [str(c or '').strip() for c in filas[0]]
    faltan = [c for c in COLS_PORTAL if c not in tit]
    if faltan:
        fallar(f'Al Libro IVA Ventas del Portal le faltan columnas: {faltan}', archivo=ruta)
    idx = [tit.index(c) for c in COLS_PORTAL]
    # "Total IVA" del CSV de ARCA viene redondeado a un decimal (35.236,9); el
    # papel usa la suma de las columnas "Importe IVA x%" (35.236,85).
    alicuotas = [i for i, t in enumerate(tit) if t.startswith('Importe IVA')]
    out = []
    for f in filas[1:]:
        if not isinstance(f[idx[0]], datetime):
            continue  # TOTAL y nota al pie
        r = [f[i] for i in idx]
        if alicuotas:
            r[7] = round(sum(f[i] or 0 for i in alicuotas), 2)
        r[1] = tipo_texto(r[1])
        for j in range(4, 9):
            r[j] = r[j] or 0
        out.append(r)
    # Orden del papel: por codigo de comprobante y fecha (estable: dentro del dia
    # queda el orden de ARCA).
    out.sort(key=lambda r: (int(str(r[1]).split(' ')[0]) if str(r[1]).split(' ')[0].isdigit() else 999, r[0]))
    return out


def limpiar(ws):
    for rango in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rango))
    ws.delete_rows(1, ws.max_row + 1)


def tiene_datos(ws):
    return any(c.value not in (None, '') for row in ws.iter_rows(max_row=min(ws.max_row, 10)) for c in row)


def escribir_portal(ws, filas, cuit, periodo):
    ws['A1'] = f'Comprobantes de Ventas - CUIT {cuit} - Portal IVA ARCA - {periodo[4:]}/{periodo[:4]}'
    for j, t in enumerate(COLS_PORTAL, start=1):
        c = ws.cell(2, j, t)
        if t in IMPORTES_PORTAL:
            c.number_format, c.font = CONTABLE, CHICA
    for i, r in enumerate(filas, start=3):
        for j, v in enumerate(r, start=1):
            c = ws.cell(i, j, v)
            if j == 1:
                c.number_format = 'dd/mm/yyyy'
            elif j >= 5:
                c.number_format, c.font = CONTABLE, CHICA
    for col, w in zip('ABCDEFGHI', (11, 24, 8, 12, 17, 16, 17, 15, 17)):
        ws.column_dimensions[col].width = w


def copiar_hoja(origen, destino):
    """Copia valores, formatos, combinadas y anchos de una hoja a otra (de otro libro)."""
    for row in origen.iter_rows():
        for c in row:
            if c.value is None and not c.has_style:
                continue
            d = destino.cell(c.row, c.column, c.value)
            if c.has_style:
                d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                d.alignment, d.number_format = copy(c.alignment), c.number_format
    for rango in origen.merged_cells.ranges:
        destino.merge_cells(str(rango))
    for k, v in origen.column_dimensions.items():
        destino.column_dimensions[k].width = v.width
    for k, v in origen.row_dimensions.items():
        if v.height:
            destino.row_dimensions[k].height = v.height
    destino.freeze_panes = origen.freeze_panes


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

    # 1) Portal IVA -> filas del papel
    portal = filas_portal(a.portal)
    if not portal:
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
        # El cruce lee la misma hoja del Portal que va al papel (con el IVA sumado
        # por alicuota, no el "Total IVA" redondeado de ARCA).
        portal_papel = os.path.join(tmp, f'{a.periodo} - PORTAL IVA - VENTAS.xlsx')
        wb_pp = openpyxl.Workbook()
        escribir_portal(wb_pp.active, portal, a.cuit, a.periodo)
        wb_pp.active.title = 'Ventas Portal IVA'
        wb_pp.save(portal_papel)
        salida_cruce = os.path.join(tmp, 'cruce.xlsx')
        p = subprocess.run([sys.executable, os.path.join(AQUI, 'cruzar_ventas.py'), f'--zetti={a.zetti}',
                            f'--portal={portal_papel}', f'--salida={salida_cruce}'],
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
        escribir_portal(ws, portal, a.cuit, a.periodo)
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
        'filas_portal': len(portal), 'filas_zetti': max(ws_z.max_row - 2, 0),
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
