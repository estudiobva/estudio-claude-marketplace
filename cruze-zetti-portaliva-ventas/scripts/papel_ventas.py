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
