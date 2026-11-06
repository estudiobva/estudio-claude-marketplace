"""Etapa 2 (DDJJ de IVA), paso 7 parte 1: clasifica los proveedores de "Compras portal iva".

    python clasificar_proveedores.py --papel=<papel de trabajo.xlsx> [--maestro=<Maestro de Proveedores.xlsx>]

Padron: solapa "MAESTRO" del Maestro de Proveedores (unidad compartida, carpeta 0 - ESTUDIO; se busca
en BVA_UNIDAD_PATH, I: y G:, o se pasa con --maestro),
titulos en la fila 2, datos desde la 3; A = CUIT, B = razon social, E = CUENTA CONTABLE.

En "Compras portal iva" toma el CUIT de la columna F (Nro. Doc. Vendedor), lo busca en el padron
(CUIT sin guiones ni ceros a la izquierda) y escribe la cuenta contable en AG (PROVEEDORES). Los que
no estan, o estan sin cuenta, van como NO CLASIFICADO. Solo escribe la columna AG de las filas de
comprobantes. Imprime un JSON con la cantidad por cuenta y los NO CLASIFICADOS (CUIT, nombre,
cantidad de comprobantes, total).
"""
import json
import os
import re
import sys
import warnings
from collections import Counter, OrderedDict
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font

warnings.filterwarnings('ignore', category=UserWarning)

MAESTRO_REL = os.path.join('0 - ESTUDIO - 00000000000', 'Maestro de Proveedores.xlsx')
# I: en las maquinas del estudio (rutas-bva.js); G: donde Drive monta con otra letra
UNIDADES = [r'I:\Unidades compartidas\BVA - Sociedades Farmaceuticas',
            r'G:\Unidades compartidas\BVA - Sociedades Farmaceuticas']
HOJA = 'Compras portal iva'
NO_CLASIFICADO = 'NO CLASIFICADO'
COL_CUIT, COL_NOMBRE, COL_TOTAL, COL_PROV = 6, 7, 8, 33  # F, G, H, AG


def unidad_env():
    """BVA_UNIDAD_PATH del entorno o de los .env del estudio (~/.fisco-ar/.env y el .env del plugin)."""
    if os.environ.get('BVA_UNIDAD_PATH', '').strip():
        return os.environ['BVA_UNIDAD_PATH'].strip()
    raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for env in (os.path.join(raiz, '.env'), os.path.join(os.path.expanduser('~'), '.fisco-ar', '.env')):
        try:
            for linea in open(env, encoding='utf-8'):
                k, _, v = linea.strip().partition('=')
                if k.strip() == 'BVA_UNIDAD_PATH' and v.strip():
                    return v.strip().strip('"')
        except OSError:
            pass
    return None


def ubicar_maestro():
    if os.environ.get('BVA_MAESTRO_PROVEEDORES'):
        return os.environ['BVA_MAESTRO_PROVEEDORES']
    bases = [b for b in [unidad_env()] + UNIDADES if b]
    for base in bases:
        if os.path.exists(os.path.join(base, MAESTRO_REL)):
            return os.path.join(base, MAESTRO_REL)
    sys.exit(json.dumps({'error': 'No encuentro el Maestro de Proveedores en la unidad compartida: pasarlo con --maestro=',
                         'buscado_en': bases}, ensure_ascii=False))


def cuit_norm(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        v = str(int(v))
    return re.sub(r'\D', '', str(v or '')).lstrip('0')


def main():
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel = opts.get('papel')
    if not papel:
        sys.exit(__doc__)
    maestro = opts.get('maestro') or ubicar_maestro()

    wm = load_workbook(maestro, read_only=True, data_only=True)['MAESTRO']
    padron, conflictos = {}, {}
    for r in wm.iter_rows(min_row=3, values_only=True):
        c = cuit_norm(r[0])
        if not c:
            continue
        cuenta = str(r[4] or '').strip()
        if c in padron and cuenta and padron[c] and padron[c] != cuenta:
            conflictos.setdefault(c, {padron[c]}).add(cuenta)
        if cuenta or c not in padron:
            padron[c] = cuenta

    wb = load_workbook(papel)
    if HOJA not in wb.sheetnames:
        sys.exit(json.dumps({'error': f'Falta la solapa "{HOJA}"', 'solapas': wb.sheetnames}, ensure_ascii=False))
    ws = wb[HOJA]
    titulo = str(ws.cell(1, COL_PROV).value or '').strip().upper()
    if titulo != 'PROVEEDORES' or str(ws.cell(1, COL_CUIT).value or '').strip().lower() != 'nro. doc. vendedor':
        sys.exit(json.dumps({'error': 'La solapa no tiene la distribucion esperada (F = Nro. Doc. Vendedor, AG = PROVEEDORES)',
                             'F1': ws.cell(1, COL_CUIT).value, 'AG1': ws.cell(1, COL_PROV).value}, ensure_ascii=False))

    por_cuenta = Counter()
    sin = OrderedDict()
    filas = 0
    for r in range(2, ws.max_row + 1):
        if not isinstance(ws.cell(r, 1).value, datetime):
            continue
        filas += 1
        c = cuit_norm(ws.cell(r, COL_CUIT).value)
        cuenta = padron.get(c) or NO_CLASIFICADO
        celda = ws.cell(r, COL_PROV, cuenta)
        celda.font = Font(name='Arial', size=10)
        por_cuenta[cuenta] += 1
        if cuenta == NO_CLASIFICADO:
            e = sin.setdefault(c, {'cuit': c, 'nombre': ws.cell(r, COL_NOMBRE).value, 'comprobantes': 0, 'total': 0.0,
                                   'motivo': 'en el padron sin cuenta contable' if c in padron else 'no esta en el padron'})
            e['comprobantes'] += 1
            e['total'] = round(e['total'] + float(ws.cell(r, COL_TOTAL).value or 0), 2)
    wb.save(papel)

    print(json.dumps({'papel': papel, 'maestro': maestro, 'comprobantes': filas,
                      'por_cuenta': dict(por_cuenta.most_common()),
                      'no_clasificados': list(sin.values()),
                      **({'cuits_con_dos_cuentas_en_padron': {k: sorted(v) for k, v in conflictos.items()}} if conflictos else {})},
                     ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
