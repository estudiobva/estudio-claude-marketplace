#!/usr/bin/env python3
"""
csv_a_excel.py — convierte el CSV del Portal IVA de ARCA al Excel entregable.

    python csv_a_excel.py <entrada.csv> <salida.xlsx> --libro=ventas

El formato lo fija la skill portal-iva-descarga-ventas: una sola hoja, las
mismas columnas del CSV en el mismo orden, tipos de dato reales (no texto),
fila TOTAL con formulas =SUM() de verdad, y una fila en blanco entre el ultimo
comprobante y el total para no romper el filtro ni el copiado.

La hoja se llama "Ventas Portal IVA" porque es el nombre que espera la skill
cruce-ventas-zetti-portal-iva: asi se copia directo al papel de trabajo.

Imprime un JSON con el resumen por stdout.
"""

import csv
import json
import re
import sys
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FUENTE = 'Arial'
AZUL = '1F4E79'
FMT_NUM = '#,##0.00;-#,##0.00;-'
FMT_ENT = '0'
FMT_FECHA = 'DD/MM/YYYY'

# Columnas que van como entero. Se matchea por patron y no por nombre exacto
# para que sirva igual en compras (Vendedor) y en ventas (Comprador).
RE_ENTERO = re.compile(
    r'^(Tipo de Comprobante|Punto de Venta|N.mero de Comprobante( Hasta)?|'
    r'Tipo Doc\.|Nro\. Doc\.)', re.I)
RE_TEXTO  = re.compile(r'^(Denominaci.n|Moneda Original)', re.I)
RE_FECHA  = re.compile(r'^Fecha', re.I)


def a_numero(s):
    """'1.234,56' -> 1234.56 · '' -> None. Miles con punto, decimales con coma."""
    s = (s or '').strip()
    if not s:
        return None
    try:
        return float(s.replace('.', '').replace(',', '.'))
    except ValueError:
        return None


def a_entero(s):
    s = (s or '').strip()
    if not s:
        return None
    try:
        return int(float(s.replace('.', '').replace(',', '.')))
    except ValueError:
        return None


def a_fecha(s):
    s = (s or '').strip()
    if not s:
        return None
    for fmt in ('%Y-%m-%d', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def tipo_de(encabezado):
    if RE_FECHA.match(encabezado):
        return 'fecha'
    if RE_ENTERO.match(encabezado):
        return 'entero'
    if RE_TEXTO.match(encabezado):
        return 'texto'
    return 'numero'


def convertir(csv_path, xlsx_path, libro):
    with open(csv_path, encoding='latin-1', newline='') as fh:
        filas = list(csv.reader(fh, delimiter=';'))

    if not filas:
        raise SystemExit('El CSV esta vacio.')

    encabezados = [h.strip() for h in filas[0]]
    datos = [f for f in filas[1:] if any((c or '').strip() for c in f)]
    tipos = [tipo_de(h) for h in encabezados]

    wb = Workbook()
    ws = wb.active
    ws.title = f'{libro.capitalize()} Portal IVA'   # "Ventas Portal IVA"

    # ── Encabezado ────────────────────────────────────────────────────────────
    for j, h in enumerate(encabezados, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = Font(name=FUENTE, size=10, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=AZUL)
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.row_dimensions[1].height = 30

    # ── Datos ─────────────────────────────────────────────────────────────────
    for i, fila in enumerate(datos, start=2):
        for j, tipo in enumerate(tipos, start=1):
            crudo = fila[j - 1] if j - 1 < len(fila) else ''
            if tipo == 'fecha':
                v, fmt = a_fecha(crudo), FMT_FECHA
            elif tipo == 'entero':
                v, fmt = a_entero(crudo), FMT_ENT
            elif tipo == 'texto':
                v, fmt = ((crudo or '').strip() or None), None
            else:
                v, fmt = a_numero(crudo), FMT_NUM
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(name=FUENTE, size=10)
            if fmt:
                c.number_format = fmt

    ultima_datos = 1 + len(datos)
    # Una fila en blanco entre el ultimo comprobante y el TOTAL: sin eso el
    # autofiltro se come la fila de totales y el copiado al papel de trabajo
    # arrastra basura.
    fila_total = ultima_datos + 2

    # ── Fila TOTAL, con formulas reales ───────────────────────────────────────
    c = ws.cell(row=fila_total, column=1, value='TOTAL')
    c.font = Font(name=FUENTE, size=10, bold=True)
    formulas = 0
    for j, tipo in enumerate(tipos, start=1):
        if tipo != 'numero' or j == 1:
            continue
        col = get_column_letter(j)
        cell = ws.cell(row=fila_total, column=j,
                       value=f'=SUM({col}2:{col}{ultima_datos})')
        cell.font = Font(name=FUENTE, size=10, bold=True)
        cell.number_format = FMT_NUM
        formulas += 1

    nota = (f'{len(datos)} comprobantes · Fuente: Portal IVA de ARCA, Libro '
            f'{libro.capitalize()} · Las notas de credito vienen con importe '
            f'negativo, o sea que el total ya es neto.')
    n = ws.cell(row=fila_total + 1, column=1, value=nota)
    n.font = Font(name=FUENTE, size=9, italic=True)

    # ── Presentacion ──────────────────────────────────────────────────────────
    ws.freeze_panes = 'A2'
    # El autofiltro cubre SOLO los datos: incluir la fila TOTAL la haria
    # desaparecer al filtrar.
    ws.auto_filter.ref = f'A1:{get_column_letter(len(encabezados))}{ultima_datos}'

    for j, h in enumerate(encabezados, start=1):
        if RE_TEXTO.match(h) and 'Moneda' not in h:
            ancho = 55
        elif RE_FECHA.match(h):
            ancho = 13
        else:
            ancho = max(12, min(28, len(h) + 3))
        ws.column_dimensions[get_column_letter(j)].width = ancho

    wb.save(xlsx_path)
    return {
        'archivo': xlsx_path,
        'hoja': ws.title,
        'columnas': len(encabezados),
        'filas': len(datos),
        'filaTotal': fila_total,
        'formulas': formulas,
    }


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    opts = dict(a[2:].split('=', 1) for a in sys.argv[1:] if a.startswith('--') and '=' in a)
    if len(args) != 2:
        raise SystemExit('Uso: csv_a_excel.py <entrada.csv> <salida.xlsx> --libro=ventas')

    resumen = convertir(args[0], args[1], opts.get('libro', 'ventas'))

    # Control de la spec: si no se escribio ninguna formula, el Excel esta mal
    # aunque abra bien — el error tipico es comparar encabezados contra indices.
    if resumen['formulas'] == 0:
        raise SystemExit(json.dumps(
            {'error': 'No se escribio ninguna formula SUM: revisar el mapeo de columnas.',
             **resumen}))

    print(json.dumps(resumen, ensure_ascii=False))


if __name__ == '__main__':
    main()
