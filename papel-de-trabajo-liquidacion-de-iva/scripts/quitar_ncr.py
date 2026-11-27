"""Resta las NCR (notas de credito de recupero) de Monroe y Suizo del Libro IVA Compras del Portal.

    python quitar_ncr.py --origen=<libro con NCR.xlsx> --ncr=<AAAAMM-COMPRAS-SIN-NCR-<SOC>.xlsx> --salida=<xlsx sin NCR>
    python quitar_ncr.py --origen=<libro.xlsx> --ncr=<conciliacion.xlsx> --verificar
    python quitar_ncr.py --origen=<libro con NCR.xlsx> --sin-ncr --salida=<xlsx>

Las NCR salen de la conciliacion del circuito conciliacion-ncr (hoja "Eliminados": drogueria, CUIT, PV, numero e
importe). Copia el libro tal cual, borra las filas de esas NCR (las busca por CUIT + PV + numero entre las notas de
credito y controla el importe), rehace la fila TOTAL (=SUM), el filtro y la nota al pie (cantidad de comprobantes y
cuantas NCR se eliminaron). El libro completo queda aparte, en la solapa "compras con NCR".

--verificar: no graba nada; dice si el libro trae TODAS las NCR de la conciliacion (sirve para elegir, entre los
  libros guardados en la carpeta, el que todavia las tiene: el de la DDJJ presentada o el borrador ya depurado no).
--sin-ncr: el periodo no tiene NCR para eliminar (sin conciliacion, o la sociedad no las tiene): copia el libro y lo
  aclara en la nota.

Si una NCR de la conciliacion no esta en el libro o tiene otro importe, corta sin grabar (codigo 1). Imprime un JSON.
"""
import argparse
import json
import re
import sys
import unicodedata
from collections import Counter

from openpyxl import load_workbook

TIPOS_NC = {3, 8, 13, 53, 112, 113, 114, 203}


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


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()
    return s.strip().lower()


def pesos(x):
    return f'$ {x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def leer_ncr(archivo):
    """{(cuit, pv, nro): {'drogueria', 'importe'}} de la hoja Eliminados de la conciliacion."""
    wn = load_workbook(archivo, read_only=True, data_only=True)
    if 'Eliminados' not in wn.sheetnames:
        fallar('La conciliacion NCR no tiene la hoja "Eliminados"', archivo=archivo, solapas=wn.sheetnames)
    filas = list(wn['Eliminados'].iter_rows(values_only=True))
    if not filas:
        return {}
    h = [norm(x) for x in filas[0]]
    try:
        ix = {k: h.index(k) for k in ('_drog', '_cuit', '_pv', '_nro', '_imp')}
    except ValueError:
        fallar('La hoja Eliminados no tiene las columnas _drog, _cuit, _pv, _nro, _imp', titulos=h)
    ncr = {}
    for r in filas[1:]:
        if not r or r[ix['_cuit']] in (None, ''):
            continue
        k = (cuit(r[ix['_cuit']]), num(r[ix['_pv']]), num(r[ix['_nro']]))
        ncr[k] = {'drogueria': str(r[ix['_drog']] or '').strip().upper(), 'importe': float(r[ix['_imp']] or 0)}
    return ncr


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--origen', required=True)
    ap.add_argument('--ncr')
    ap.add_argument('--salida')
    ap.add_argument('--verificar', action='store_true')
    ap.add_argument('--sin-ncr', action='store_true')
    a = ap.parse_args()
    if not a.sin_ncr and not a.ncr:
        fallar('Falta --ncr (conciliacion NCR) o --sin-ncr')
    if not a.verificar and not a.salida:
        fallar('Falta --salida')

    ncr = {} if a.sin_ncr else leer_ncr(a.ncr)

    wb = load_workbook(a.origen)
    ws = wb.worksheets[0]
    tit = [norm(c.value) for c in ws[1]]

    def col(pref):
        for i, t in enumerate(tit):
            if t.startswith(pref):
                return i + 1
        fallar(f'No encuentro la columna "{pref}" en el libro', archivo=a.origen, titulos=tit)
    cT, cPV, cN, cDoc, cH = (col('tipo de comprobante'), col('punto de venta'), col('numero de comprobante'),
                             col('nro. doc. vendedor'), col('importe total'))
    fila_total = next((r for r in range(2, ws.max_row + 1) if str(ws.cell(r, 1).value or '').strip().upper() == 'TOTAL'), None)
    if not fila_total:
        fallar('No encuentro la fila TOTAL del libro', archivo=a.origen)
    datos = [r for r in range(2, fila_total) if isinstance(ws.cell(r, cT).value, (int, float))]
    if not datos:
        fallar('El libro no trae comprobantes', archivo=a.origen)
    ultima = datos[-1]
    antes_n = len(datos)
    antes_h = sum(float(ws.cell(r, cH).value or 0) for r in datos)

    borrar, encontradas, mal_importe = [], set(), []
    for r in datos:
        if num(ws.cell(r, cT).value) not in TIPOS_NC:
            continue
        k = (cuit(ws.cell(r, cDoc).value), num(ws.cell(r, cPV).value), num(ws.cell(r, cN).value))
        if k in ncr:
            imp = float(ws.cell(r, cH).value or 0)
            if abs(imp - ncr[k]['importe']) > 0.01:
                mal_importe.append({'cuit': k[0], 'pv': k[1], 'nro': k[2], 'libro': imp, 'conciliacion': ncr[k]['importe']})
                continue
            borrar.append(r)
            encontradas.add(k)
    faltan = [{'cuit': k[0], 'pv': k[1], 'nro': k[2], **v} for k, v in ncr.items() if k not in encontradas]

    if a.verificar:
        print(json.dumps({'ok': True, 'origen': a.origen, 'comprobantes': antes_n, 'ncr_conciliacion': len(ncr),
                          'ncr_en_el_libro': len(encontradas), 'contiene_todas': not faltan and not mal_importe,
                          'faltan': len(faltan), 'importe_distinto': len(mal_importe)}, ensure_ascii=False))
        return
    if faltan or mal_importe:
        fallar('Hay NCR de la conciliacion que no estan en el libro o tienen otro importe: no grabo',
               origen=a.origen, no_encontradas=faltan[:30], cantidad_no_encontradas=len(faltan), importe_distinto=mal_importe)

    nota_fila = next((r for r in range(fila_total + 1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str)), None)
    nota_txt = ws.cell(nota_fila, 1).value if nota_fila else ''
    sumas = {c.column: c.value for c in ws[fila_total] if isinstance(c.value, str) and c.value.startswith('=SUM(')}

    for r in sorted(borrar, reverse=True):
        ws.delete_rows(r)
    n = len(borrar)
    ultima -= n
    fila_total -= n
    if nota_fila:
        nota_fila -= n
    for c, f in sumas.items():
        ws.cell(fila_total, c).value = re.sub(r'(\$?[A-Z]+\$?)2:(\$?[A-Z]+\$?)\d+', rf'\g<1>2:\g<2>{ultima}', f)
    if ws.auto_filter.ref:
        ws.auto_filter.ref = re.sub(r'\d+$', str(ultima), ws.auto_filter.ref)

    total_ncr = sum(ncr[k]['importe'] for k in encontradas)
    por_drog = Counter(ncr[k]['drogueria'] for k in encontradas)
    partes = nota_txt.split(' · ', 1) if nota_txt else ['', '']
    nota = re.sub(r'^\d+ comprobantes', f'{antes_n - n} comprobantes', partes[0] or f'{antes_n - n} comprobantes')
    if a.sin_ncr:
        nota += ' · Sin NCR para eliminar (sin conciliacion NCR del periodo): igual a la solapa "compras con NCR"'
    else:
        detalle = ', '.join(f'{d.title()} {c}' for d, c in sorted(por_drog.items())) or 'ninguna'
        nota += (f' · Sin las {n} NCR eliminadas de las droguerias ({detalle}) por {pesos(total_ncr)}: el libro completo '
                 'esta en la solapa "compras con NCR" y el detalle de las NCR, debajo del cuadro')
    if len(partes) > 1 and partes[1]:
        nota += ' · ' + partes[1]
    ws.cell(nota_fila or fila_total + 1, 1).value = nota

    despues = [r for r in range(2, fila_total) if isinstance(ws.cell(r, cT).value, (int, float))]
    despues_h = sum(float(ws.cell(r, cH).value or 0) for r in despues)
    wb.save(a.salida)
    print(json.dumps({'ok': True, 'salida': a.salida, 'sin_ncr': a.sin_ncr, 'comprobantes_con_ncr': antes_n,
                      'importe_total_con_ncr': round(antes_h, 2), 'ncr_eliminadas': n, 'ncr_por_drogueria': dict(por_drog),
                      'importe_ncr': round(total_ncr, 2), 'comprobantes_sin_ncr': len(despues),
                      'importe_total_sin_ncr': round(despues_h, 2), 'control': round(antes_h - total_ncr - despues_h, 2)},
                     ensure_ascii=False))


if __name__ == '__main__':
    main()
