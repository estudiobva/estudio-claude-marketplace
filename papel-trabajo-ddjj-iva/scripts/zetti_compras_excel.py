"""Subdiario de IVA Compras de Zetti (reporte 5.6.4, CSV) -> Excel.

Mismo criterio que zetti_papel_trabajo.py (ventas), adaptado a las columnas de compras.

    python3 zetti_compras_excel.py <entrada.csv> <salida.xlsx> [--periodo=AAAAMM]

CSV de T&S Web: UTF-8 (a veces latin-1), separado por ";" con columnas vacias intercaladas.
    titulos  Fecha;;TC;M;Nro. Comp.; Proveedor;;;;CUIT;;;;R;Exen;;;Grav;;;;IVA;;P/R IIBB;;;P/R IVA;;;P/R IG;Total
    ultima   ";;...Total: ;;<Exen>;;;;<Grav>;;;<IVA>;;<P/R IIBB>;;;<P/R IVA>;;;<P/R IG>;<Total>"
    pie      "; :   -  ;...Página ;...;1"

Salida: hoja "Compras Zetti"
    fila 1  "<razon social>  -  CUIT NN-NNNNNNNN-N  -  Subdiario de IVA Compras  -  Desde: DD-MM-AAAA  Hasta: DD-MM-AAAA"
    fila 2  Fecha | TC | M | Nro. Comp. (D = punto de venta, E = numero) | Proveedor | CUIT | R
            | Exen | Grav | IVA | P/R IIBB | P/R IVA | P/R IG | Total
    datos   desde la fila 3, en el orden original de Zetti; al final fila TOTAL (I a O).

Imprime un JSON con razon social, CUIT, fechas, cantidad por TC, totales y control contra la
fila "Total:" del reporte. Sale con 1 si el Total no coincide (tolerancia $0,05 en las demas
columnas), si hay filas que no reconoce o fechas fuera del periodo.
"""
import csv
import json
import re
import sys
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter

from fuente import ZETTI, aplicar, letra

COLUMNAS = [
    ('Fecha', 'fecha'), ('TC', 'tc'), ('M', 'm'), ('Nro. Comp.', 'nro'),
    ('Proveedor', 'proveedor'), ('CUIT', 'cuit'), ('R', 'resp'),
    ('Exen', 'exen'), ('Grav', 'grav'), ('IVA', 'iva'),
    ('P/R IIBB', 'priibb'), ('P/R IVA', 'priva'), ('P/R IG', 'prig'), ('Total', 'total'),
]
IMPORTES = ['exen', 'grav', 'iva', 'priibb', 'priva', 'prig', 'total']
TEXTO = ['tc', 'm', 'nro', 'proveedor', 'cuit', 'resp']
TOLERANCIA = 0.05


def fallar(msg, **extra):
    print(json.dumps({'error': msg, **extra}, ensure_ascii=False))
    sys.exit(1)


def importe(s):
    s = (s or '').strip()
    if not s:
        return 0.0
    try:
        return float(s.replace('.', '').replace(',', '.'))
    except ValueError:
        fallar(f'Importe que no entiendo: "{s}"')


def fecha(s):
    s = (s or '').strip()
    for fmt in ('%d/%m/%y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def tras(fila, etiqueta):
    for i, c in enumerate(fila):
        if c.strip().lower().startswith(etiqueta.lower()):
            for v in fila[i + 1:]:
                if v.strip():
                    return v.strip()
    return ''


def main():
    pos = [a for a in sys.argv[1:] if not a.startswith('--')]
    opts = dict(a[2:].split('=', 1) for a in sys.argv[1:] if a.startswith('--') and '=' in a)
    if len(pos) != 2:
        sys.exit(__doc__)
    entrada, salida = pos
    periodo = opts.get('periodo')

    crudo = open(entrada, 'rb').read()
    try:
        texto = crudo.decode('utf-8-sig')
    except UnicodeDecodeError:
        texto = crudo.decode('latin-1')
    filas = list(csv.reader(texto.splitlines(), delimiter=';'))

    es_titulo = lambda f: f and f[0].strip() == 'Fecha' and 'TC' in [c.strip() for c in f]
    i_tit = next((i for i, f in enumerate(filas) if es_titulo(f)), None)
    if i_tit is None:
        fallar('No encontre la fila de titulos (Fecha;;TC;...) en el CSV.')
    titulos = [c.strip() for c in filas[i_tit]]
    idx = {}
    for nombre, clave in COLUMNAS:
        if nombre not in titulos:
            fallar(f'Falta la columna "{nombre}" en el CSV.', titulos=[t for t in titulos if t])
        idx[clave] = titulos.index(nombre)
    ancho = max(idx.values()) + 2
    # En el 5.6.4 el importe de Grav viene una columna a la derecha de su titulo
    # (titulo en 17, dato en 18). Si la celda del titulo esta vacia se toma la
    # siguiente, siempre que no sea la columna de otro titulo.
    ocupadas = set(idx.values())

    def celda(f, k):
        v = f[idx[k]]
        if not v.strip() and idx[k] + 1 not in ocupadas:
            v = f[idx[k] + 1]
        return v

    razon = cuit_soc = desde = hasta = nodo = ''
    for f in filas[:i_tit]:
        razon = razon or tras(f, 'Raz')
        cuit_soc = cuit_soc or tras(f, 'C.U.I.T')
        desde = desde or tras(f, 'Fecha des')
        hasta = hasta or tras(f, 'Fecha has')
        nodo = nodo or tras(f, 'Nodo')
    cuit_soc = re.sub(r'\D', '', cuit_soc)

    datos, total_reporte, raras = [], None, []
    for n, f in enumerate(filas[i_tit + 1:], start=i_tit + 2):
        if not any(c.strip() for c in f) or es_titulo(f):
            continue
        f = f + [''] * (ancho - len(f))
        d = fecha(f[idx['fecha']])
        if d is None:
            celdas = [c.strip().lower() for c in f]
            if any(c.startswith('total') for c in celdas):
                total_reporte = {k: importe(celda(f, k)) for k in IMPORTES}
            elif any(c.startswith('página') or c.startswith('pagina') for c in celdas):
                pass  # pie de pagina
            else:
                raras.append({'fila': n, 'texto': ';'.join(c for c in f if c.strip())[:120]})
            continue
        r = {'fecha': d}
        for k in TEXTO:
            r[k] = f[idx[k]].strip()
        for k in IMPORTES:
            r[k] = importe(celda(f, k))
        datos.append(r)

    if raras:
        fallar('Hay filas que no son comprobantes ni el total: revisar el CSV.', filas=raras[:20])
    if not datos:
        fallar('El reporte no trae comprobantes.', cuit=cuit_soc, desde=desde, hasta=hasta, sin_datos=True)
    if total_reporte is None:
        fallar('El CSV no trae la fila "Total:": puede haber bajado cortado.')

    fuera = []
    if periodo:
        fuera = sorted({r['fecha'].strftime('%d/%m/%Y') for r in datos if r['fecha'].strftime('%Y%m') != periodo})

    suma = {k: round(sum(r[k] for r in datos), 2) for k in IMPORTES}
    difs = {k: round(suma[k] - total_reporte[k], 2) for k in IMPORTES if abs(suma[k] - total_reporte[k]) >= 0.005}
    graves = {k: v for k, v in difs.items() if k == 'total' or abs(v) > TOLERANCIA}

    sin_separar = 0
    for r in datos:
        m = re.fullmatch(r'(\d+)-(\d+)', r['nro'])
        if m:
            r['pv'], r['numero'] = int(m.group(1)), int(m.group(2))
        else:
            r['pv'], r['numero'] = r['nro'], None
            sin_separar += 1

    cuit_fmt = f'{cuit_soc[:2]}-{cuit_soc[2:10]}-{cuit_soc[10:]}' if len(cuit_soc) == 11 else cuit_soc
    fd, fh = fecha(desde), fecha(hasta)
    d_txt = fd.strftime('%d-%m-%Y') if fd else desde
    h_txt = fh.strftime('%d-%m-%Y') if fh else hasta

    columnas = [('Fecha', 'fecha'), ('TC', 'tc'), ('M', 'm'), ('Nro. Comp.', 'pv'), (None, 'numero'),
                ('Proveedor', 'proveedor'), ('CUIT', 'cuit'), ('R', 'resp'),
                ('Exen', 'exen'), ('Grav', 'grav'), ('IVA', 'iva'), ('P/R IIBB', 'priibb'),
                ('P/R IVA', 'priva'), ('P/R IG', 'prig'), ('Total', 'total')]
    # Formato numero como el Zetti original (no contabilidad) y letra Aptos Narrow 11 (fuente.py).
    numero = '#,##0.00'
    anchos = {'fecha': 11, 'tc': 5, 'm': 4, 'pv': 6, 'numero': 10, 'proveedor': 40, 'cuit': 15, 'resp': 4,
              'exen': 17, 'grav': 17, 'iva': 16, 'priibb': 15, 'priva': 14, 'prig': 13, 'total': 17}
    ult = 2 + len(datos)
    ft = ult + 1
    c_ini, c_fin = 9, len(columnas)  # I a O

    wb = Workbook()
    ws = wb.active
    ws.title = 'Compras Zetti'
    ws['A1'] = f'{razon}  -  CUIT {cuit_fmt}  -  Subdiario de IVA Compras  -  Desde: {d_txt}  Hasta: {h_txt}'
    for j, (t, k) in enumerate(columnas, start=1):
        ws.cell(2, j, t)
    for i, r in enumerate(datos, start=3):
        for j, (_, k) in enumerate(columnas, start=1):
            v = r[k]
            c = ws.cell(i, j, v if v != '' else None)
            if k == 'fecha':
                c.number_format = 'dd/mm/yyyy'
            elif k in IMPORTES:
                c.number_format = numero
    ws.cell(ft, 8, 'TOTAL').font = letra(ZETTI, bold=True)
    for j in range(c_ini, c_fin + 1):
        col = get_column_letter(j)
        c = ws.cell(ft, j, f'=SUM({col}3:{col}{ult})')
        c.number_format = numero
        c.font = letra(ZETTI, bold=True)
        c.border = Border(top=Side(style='thin'))
    for j, (_, k) in enumerate(columnas, start=1):
        ws.column_dimensions[get_column_letter(j)].width = anchos[k]
    aplicar(ws, ZETTI)
    ws.freeze_panes = 'A3'
    ws.auto_filter.ref = f'A2:{get_column_letter(c_fin)}{ult}'
    wb.save(salida)

    por_tc = {}
    for r in datos:
        por_tc[r['tc']] = por_tc.get(r['tc'], 0) + 1
    res = {
        'razon_social': razon, 'cuit': cuit_soc, 'nodo': nodo, 'desde': desde, 'hasta': hasta,
        'comprobantes': len(datos), 'por_tc': por_tc, 'totales': suma,
        'control_total_reporte': 'ok' if not difs else ('redondeo' if not graves else 'no coincide'),
        **({'diferencias_total_reporte': difs} if difs else {}),
        'fechas_fuera_de_periodo': fuera, 'xlsx': salida,
    }
    if sin_separar:
        res['aviso_nro'] = f'{sin_separar} comprobantes con numero que no es PV-NUMERO: quedaron enteros en la columna D.'
    if graves:
        res['error'] = 'La suma de los comprobantes no coincide con el Total del reporte.'
    elif fuera:
        res['error'] = f'Hay comprobantes fuera del periodo {periodo}.'
    print(json.dumps(res, ensure_ascii=False))
    sys.exit(1 if 'error' in res else 0)


if __name__ == '__main__':
    main()
