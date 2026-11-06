"""Proceso TSH / IIBB, paso 6 (despues del cruze del paso 5, se corre siempre,
haya o no cruze): cruce de ventas Mis Comprobantes / Portal IVA / Zetti.

    python cruce_3_fuentes.py <carpeta IVA del mes> [--excepcion="2:NC B"] [--salida=<xlsx>]
                              [--mc=<xlsx>] [--portal=<xlsx>] [--zetti=<xlsx>] [--tolerancia=0.5]
    python cruce_3_fuentes.py --papel=<papel de trabajo.xlsx> [--excepcion="2:NC B"]

Con --papel (el uso del proceso, paso 5) lee las solapas "Ventas MCR", "Ventas portal
iva" y "Ventas Zetti" (la original: la depuracion de PV se repite aca con la misma
regla, para poder listar lo eliminado) y escribe o reemplaza la solapa "cruce ventas".

Lee de la carpeta IVA del periodo (o de las rutas que se pasen):
    AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx   (NC en positivo: se pasan a negativo)
    AAAAMM - PORTAL IVA - VENTAS.xlsx
    AAAAMM - ZETTI - VENTAS.xlsx                (reporte 5.6.5)
y deja `AAAAMM - CRUCE VENTAS MC PORTAL ZETTI.xlsx` en 01-Impuestos Mensuales/AAAAMM/
(la carpeta de arriba de IVA) con UNA sola solapa "Cruce": resumen (puntos 1 a 3),
detalle de las diferencias (punto 4) y, abajo, "comprobantes de Zetti ELIMINADOS".

Reconstruido a partir del plugin cruce-ventas-3-fuentes, con un cambio de fondo:
antes de cruzar, Zetti se depura con la regla del paso 4 (depurar_pv.py): se
eliminan TODOS los comprobantes cuyo punto de venta no figura en el Portal IVA
(fallidos de Zetti: PV 1, 2, 9999...), no solo las NC B a consumidor final. El
plugin dejaba en el cruce, como "Solo en Zetti", las facturas/NC a clientes con
nombre de esos PV (caso real: una FA A y una NC B del PV 2).
--excepcion funciona igual que en depurar_pv.py.

Como se cruza:
  * Facturas, NC y ND: por (tipo ARCA, PV, numero); TC + letra de Zetti -> codigo ARCA.
  * Tiques: sumados por (fecha, PV) contra el Z de Zetti. Dos dias que se compensan
    = corrimiento de fecha (el Z quedo con la fecha del dia de al lado).
  * Exento = No gravado + Exento; Gravado = neto de todas las alicuotas;
    IVA = suma de los IVA por alicuota (no el "Total IVA" de ARCA).
"""
import json
import os
import re
import sys
import unicodedata
from collections import OrderedDict, defaultdict
from datetime import date, datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FUENTES = ('mc', 'portal', 'zetti')
NOMBRE = {'mc': 'Mis Comprobantes', 'portal': 'Portal IVA', 'zetti': 'Zetti'}
COMP = ('exen', 'grav', 'iva', 'total')

TIPOS = {
    1: 'Factura A', 2: 'Nota de Debito A', 3: 'Nota de Credito A',
    6: 'Factura B', 7: 'Nota de Debito B', 8: 'Nota de Credito B',
    11: 'Factura C', 12: 'Nota de Debito C', 13: 'Nota de Credito C',
    51: 'Factura M', 52: 'Nota de Debito M', 53: 'Nota de Credito M',
    81: 'Tique Factura A', 82: 'Tique Factura B', 83: 'Tique', 110: 'Tique Nota de Credito',
    111: 'Tique Factura C', 112: 'Tique NC A', 113: 'Tique NC B', 114: 'Tique NC C',
    115: 'Tique ND A', 116: 'Tique ND B', 117: 'Tique ND C',
    201: 'FCE MiPyME A', 203: 'NC MiPyME A', 206: 'FCE MiPyME B', 208: 'NC MiPyME B',
}
NOTAS_CREDITO = {3, 8, 13, 21, 53, 90, 110, 112, 113, 114, 119, 203, 208, 213}
TIQUES = {81, 82, 83, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120}
ZETTI_A_ARCA = {
    ('FV', 'A'): 1, ('ND', 'A'): 2, ('NC', 'A'): 3,
    ('FV', 'B'): 6, ('ND', 'B'): 7, ('NC', 'B'): 8,
    ('FV', 'C'): 11, ('ND', 'C'): 12, ('NC', 'C'): 13,
    ('FV', 'M'): 51, ('ND', 'M'): 52, ('NC', 'M'): 53,
}

F_TIT = Font(name='Arial', size=12, bold=True)
F_B = Font(name='Arial', size=10, bold=True)
F_N = Font(name='Arial', size=10)
F_H = Font(name='Arial', size=10, bold=True, color='FFFFFF')
FILL_H = PatternFill('solid', fgColor='1F4E79')
FILL = {'mc': PatternFill('solid', fgColor='DDEBF7'), 'portal': PatternFill('solid', fgColor='E2EFDA'),
        'zetti': PatternFill('solid', fgColor='FFF2CC'), 'dif': PatternFill('solid', fgColor='FCE4D6')}
FILL_ROJO = PatternFill('solid', fgColor='F8CBAD')
FMT = '#,##0.00;-#,##0.00;-'
FMT_F = 'DD/MM/YYYY'


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()
    return ' '.join(s.lower().split())


def r2(x):
    return round(x + 0.0, 2)


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def to_int(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(v)
    s = str(v or '').strip()
    return int(s) if s.isdigit() else (s or None)


def como_fecha(v):
    if isinstance(v, datetime):
        return v.date()
    return v if isinstance(v, date) else None


def pesos(x):
    return f'{x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def vacio():
    return {k: 0.0 for k in COMP}


def sumar(acc, fila):
    for k in COMP:
        acc[k] += num(fila[k])


# ---------------------------------------------------------------------------
# Lectura de las fuentes
# ---------------------------------------------------------------------------

def leer_arca(ws, fuente):
    filas_ws = list(ws.iter_rows(values_only=True))
    tit = [norm(v) for v in filas_ws[0]]
    if fuente == 'mc':
        mapa = {'fecha de emision': 'fecha', 'tipo de comprobante': 'tipo', 'punto de venta': 'punto',
                'numero desde': 'numero', 'numero hasta': 'hasta', 'denominacion receptor': 'cliente',
                'imp. neto gravado total': 'grav', 'imp. neto no gravado': 'no_grav',
                'imp. op. exentas': 'exento', 'imp. total': 'total'}
        es_iva = re.compile(r'^iva \d')
    else:
        mapa = {'fecha de emision': 'fecha', 'tipo de comprobante': 'tipo', 'punto de venta': 'punto',
                'numero de comprobante': 'numero', 'numero de comprobante hasta': 'hasta',
                'denominacion comprador': 'cliente', 'total neto gravado': 'grav',
                'importe no gravado': 'no_grav', 'importe exento': 'exento', 'importe total': 'total'}
        es_iva = re.compile(r'^importe iva \d')
    col = {mapa[h]: i for i, h in enumerate(tit) if h in mapa}
    faltan = set(mapa.values()) - set(col)
    if faltan:
        sys.exit(json.dumps({'error': f'{os.path.basename(ruta)}: faltan columnas {sorted(faltan)}'}))
    cols_iva = [i for i, h in enumerate(tit) if es_iva.match(h)]
    out = []
    for vals in filas_ws[1:]:
        g = lambda k: vals[col[k]]  # noqa: E731
        fecha = como_fecha(g('fecha'))
        if fecha is None:
            continue
        tipo = to_int(g('tipo'))
        f = {'fecha': fecha, 'tipo': tipo, 'punto': to_int(g('punto')), 'numero': to_int(g('numero')),
             'hasta': to_int(g('hasta')), 'cliente': str(g('cliente') or '').strip(),
             'exen': num(g('no_grav')) + num(g('exento')), 'grav': num(g('grav')),
             'iva': sum(num(vals[i]) for i in cols_iva), 'total': num(g('total'))}
        if fuente == 'mc' and tipo in NOTAS_CREDITO and f['total'] > 0:
            for k in COMP:
                f[k] = -f[k]
        out.append(f)
    return out


def leer_zetti(ws):
    """Acepta el Excel de Zetti de la carpeta IVA (Nro. Comp. '0012-00000581') y la
    solapa "Ventas Zetti" del papel (punto de venta en D y numero en E)."""
    filas_ws = list(ws.iter_rows(values_only=True))
    enc = next(i for i, v in enumerate(filas_ws) if norm(v[0]) == 'fecha' and norm(v[1]) == 'tc')
    tit = [norm(v) for v in filas_ws[enc]]
    c = {k: tit.index(k) for k in ('fecha', 'tc', 'm', 'nro. comp.', 'cliente', 'exen', 'grav', 'iva', 'total')}
    cn = c['nro. comp.']
    razon = next((v[0] for v in filas_ws[:enc] if isinstance(v[0], str) and v[0].strip() and 'T&S' not in v[0]), '')
    razon = razon.split('  -  ')[0].strip()
    out = []
    for v in filas_ws[enc + 1:]:
        fecha = como_fecha(v[c['fecha']])
        if fecha is None:
            continue
        if isinstance(v[cn], (int, float)) and not isinstance(v[cn], bool):
            pv, nro = v[cn], v[cn + 1]
        else:
            pv, _, nro = str(v[cn] or '').partition('-')
        tc = str(v[c['tc']] or '').strip().upper()
        m = str(v[c['m']] or '').strip().upper()
        out.append({'fecha': fecha, 'tc': tc, 'm': m, 'tc_zetti': f'{tc} {m}'.strip(),
                    'tipo': ZETTI_A_ARCA.get((tc, m)), 'punto': to_int(pv), 'numero': to_int(nro),
                    'cliente': str(v[c['cliente']] or '').strip(),
                    **{k: num(v[c[k]]) for k in COMP}})
    return out, razon


def periodo_vecino(per, delta):
    a, m = int(per[:4]), int(per[4:]) + delta
    if m == 0:
        a, m = a - 1, 12
    elif m == 13:
        a, m = a + 1, 1
    return f'{a}{m:02d}'


def leer_zetti_mes(mes, delta):
    """Ventas de Zetti del mes anterior (delta=-1) o siguiente (+1): del papel de trabajo de
    ese mes (solapa Ventas Zetti) o, si no hay, de <AAAAMM>/IVA/<AAAAMM> - ZETTI - VENTAS.xlsx.
    Solo se usa para aclarar (primer tique del mes, comprobantes en otro periodo): el cuadro
    de correlatividad sale del Zetti del mes. Devuelve (periodo, filas, origen) o (periodo, [], None)."""
    per = periodo_vecino(os.path.basename(mes), delta)
    dir_v = os.path.join(os.path.dirname(mes), per)
    if not os.path.isdir(dir_v):
        return per, [], None
    candidatos = sorted(os.path.join(dir_v, f) for f in os.listdir(dir_v)
                        if f.lower().endswith('.xlsx') and not f.startswith('~$') and 'papel' in f.lower())
    candidatos.append(os.path.join(dir_v, 'IVA', f'{per} - ZETTI - VENTAS.xlsx'))
    for ruta in candidatos:
        if not os.path.exists(ruta):
            continue
        try:
            wb = load_workbook(ruta, read_only=True, data_only=True)
            hoja = next((h for h in wb.sheetnames if norm(h) == 'ventas zetti'), None) \
                or next((h for h in wb.sheetnames if 'zetti' in norm(h) and 'compra' not in norm(h)
                         and 'depurad' not in norm(h)), None)
            if hoja is None:
                continue
            filas_z, _ = leer_zetti(wb[hoja])
        except Exception:  # noqa: BLE001 - un papel viejo con otro formato no debe cortar el cruce
            continue
        if filas_z:
            return per, filas_z, f'{os.path.basename(ruta)} ({hoja})'
    return per, [], None


def z_con_ti(filas_z):
    """{pv: [(fecha, nro Z, ultimo tique)]} de los cierres Z con 'TI nnn', ordenados por tique."""
    out = defaultdict(list)
    for r in filas_z:
        mt = re.search(r'TI\s*(\d+)\s*-\s*(\d+)', r['cliente']) if r['tc'] == 'Z' else None
        if mt:
            out[r['punto']].append((r['fecha'], r['numero'], int(mt.group(2))))
    return {pv: sorted(v, key=lambda x: x[2]) for pv, v in out.items()}


def otros_periodos(datos, filas, vecinos):
    """Comprobantes de los listados de ARCA del mes que Zetti tiene en otro periodo.
    vecinos: {'anterior': (per, filas), 'siguiente': (per, filas)}. Devuelve {per: [items]}.
      * tiques: numero <= ultimo tique del ultimo Z del mes anterior (Zetti los cerro el mes
        anterior) o > ultimo tique del mes (Zetti los cierra en el siguiente);
      * facturas, NC y ND que faltan en el Zetti del mes y estan en el de otro mes."""
    out = defaultdict(list)
    z_mes = z_con_ti(datos['zetti'])
    tiq = defaultdict(dict)
    for f in ('portal', 'mc'):
        for r in datos[f]:
            if r['tipo'] == 83 and isinstance(r['numero'], int):
                tiq[r['punto']].setdefault(r['numero'], r)
    for lado, (per, filas_v) in vecinos.items():
        if not filas_v:
            continue
        z_v = z_con_ti(filas_v)
        nros_z_mes = defaultdict(set)
        for r in datos['zetti']:
            if r['tc'] == 'Z' and isinstance(r['numero'], int):
                nros_z_mes[r['punto']].add(r['numero'])
        for pv, zs_mes in z_mes.items():
            zs_v = z_v.get(pv)
            if not zs_v:
                continue
            if tiq[pv] and set(tiq[pv]) <= nros_z_mes[pv] | {x[1] for x in zs_v}:
                # Esquema viejo (2025): ARCA numera por cierre Z. Se buscan los Z de ARCA que
                # no estan en el Zetti del mes y si en el del otro periodo.
                # Solo si el importe coincide: una sociedad informo en ARCA el Z 1685 como "1695",
                # que es otro Z (de octubre, con otro importe).
                zv = {}
                for x in filas_v:
                    if x['tc'] == 'Z' and x['punto'] == pv and isinstance(x['numero'], int):
                        zv.setdefault(x['numero'], []).append((x['fecha'], x['total']))
                for n, r in sorted(tiq[pv].items()):
                    igual = next((x for x in zv.get(n, []) if abs(x[1] - r['total']) <= 1), None)
                    if n not in nros_z_mes[pv] and igual:
                        out[per].append({'comp': f'Cierre Z {pv}-{n}', 'fecha_arca': r['fecha'],
                                         'total_arca': r['total'], 'fecha_zetti': igual[0],
                                         'en_zetti': f'cierre Z {pv}-{n}' + (f' (cargado {len(zv[n])} veces)'
                                                                            if len(zv[n]) > 1 else ''),
                                         'total_zetti': igual[1]})
                continue
            # Limites por el ultimo Z por numero, no por el TI mas alto: un TI mal cargado en un Z
            # (caso real: un Z con un tique mayor que el del Z siguiente) no debe mover el corte.
            lim_ant = max(zs_v, key=lambda x: x[1])[2]
            lim_mes = max(zs_mes, key=lambda x: x[1])[2]
            for n, r in sorted(tiq[pv].items()):
                if (lado == 'anterior' and n <= lim_ant) or (lado == 'siguiente' and n > lim_mes):
                    z = next((x for x in zs_v if x[2] >= n), None)
                    if z:
                        out[per].append({'comp': f'Tique {pv:05d}-{n:08d}', 'fecha_arca': r['fecha'],
                                         'total_arca': r['total'], 'fecha_zetti': z[0],
                                         'en_zetti': f'dentro del cierre Z {pv}-{z[1]} (TI hasta {z[2]})',
                                         'total_zetti': None})
        idx = {(r['tipo'], r['punto'], r['numero']): r for r in filas_v if r['tc'] != 'Z'}
        for rg in filas:
            if rg['modo'] != 'comp' or 'zetti' in rg['fuentes']:
                continue
            z = idx.get((rg['tipo'], rg['punto'], rg['numero']))
            if z:
                v = rg['valores']['portal'] if 'portal' in rg['fuentes'] else rg['valores']['mc']
                out[per].append({'comp': concepto(rg), 'fecha_arca': rg['fecha'], 'total_arca': v['total'],
                                 'fecha_zetti': z['fecha'], 'en_zetti': f"{z['tc_zetti']} {z['punto']}-{z['numero']}",
                                 'total_zetti': z['total']})
    return out


def bloque_otros_periodos(ws, r0, otros):
    """Un listado por periodo, debajo de la correlatividad. Devuelve la fila siguiente libre."""
    r = r0
    for per in sorted(otros):
        items = otros[per]
        ws.cell(r, 1, f'comprobantes de Zetti que figuran en el periodo {per}').font = F_TIT
        h = r + 1
        cab = ['Comprobante (ARCA)', 'Fecha en ARCA', 'Total en ARCA', 'Fecha en Zetti', 'En Zetti', 'Total en Zetti']
        for j, t in enumerate(cab, start=1):
            c = ws.cell(h, j, t)
            c.font, c.fill = F_H, FILL_H
            c.alignment = Alignment(horizontal='center')
        r = h + 1
        for x in items:
            vals = [x['comp'], x['fecha_arca'], r2(x['total_arca']), x['fecha_zetti'], x['en_zetti'],
                    r2(x['total_zetti']) if x['total_zetti'] is not None else None]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(r, j, v)
                c.font = F_N
                c.number_format = FMT_F if j in (2, 4) else FMT if j in (3, 6) else 'General'
            r += 1
        ws.cell(r, 1, f'{len(items)} comprobantes').font = F_B
        r += 3
    return r


def excepciones(arg):
    out = {}
    for parte in filter(None, (p.strip() for p in (arg or '').split(';'))):
        pv, _, tipos = parte.partition(':')
        out[to_int(pv)] = {' '.join(t.split()).upper() for t in tipos.split(',') if t.strip()}
    return out


def depurar_zetti(zetti, portal, exc):
    """Regla del paso 4: fuera todo comprobante de Zetti cuyo PV no esta en el Portal."""
    pv_portal = {r['punto'] for r in portal}
    # Cierre Z con PV mal tipeado (caso real: Z cargado con PV 41 en vez de 14): si el PV
    # no esta en el Portal y el numero sigue la secuencia de Z de un PV que si esta, se toma ese PV.
    z_por_pv = defaultdict(set)
    for r in zetti:
        if r['tc'] == 'Z' and r['punto'] in pv_portal and r['numero']:
            z_por_pv[r['punto']].add(r['numero'])
    depurar_zetti.corregidos = []
    for r in zetti:
        if r['tc'] == 'Z' and r['punto'] not in pv_portal and r['punto'] not in exc and r['numero']:
            cand = [p for p, ns in z_por_pv.items() if r['numero'] - 1 in ns or r['numero'] + 1 in ns]
            if len(cand) == 1:
                depurar_zetti.corregidos.append((r['punto'], cand[0], r['numero'], r['fecha']))
                r['punto'] = cand[0]
    quedan, fuera = [], []
    for r in zetti:
        borrar = (r['tc_zetti'] in exc[r['punto']]) if r['punto'] in exc else (r['punto'] not in pv_portal)
        (fuera if borrar else quedan).append(r)
    return quedan, fuera


# ---------------------------------------------------------------------------
# Cruce
# ---------------------------------------------------------------------------

def es_esquema_viejo(arca_nums, z_nums):
    """Hasta 2025 ARCA informa un renglon tipo 83 por cierre Z, con el numero del Z. Se toma
    como esquema viejo si al menos el 80% de esos numeros son Z de Zetti del mes (puede haber
    alguno de otro mes, como una sociedad que informo otra vez un Z del mes anterior)."""
    arca_nums = set(arca_nums)
    return bool(arca_nums) and len(arca_nums & set(z_nums)) >= 0.8 * len(arca_nums)


def alinear_fechas_esquema_viejo(datos):
    """En el esquema viejo cruza los tiques por numero de Z y no por fecha: ARCA puede fechar el
    cierre Z uno o dos dias antes que Zetti (caso real). Los renglones tipo 83 de ARCA cuyo
    numero es un Z de Zetti toman la fecha de Zetti (la original queda en 'fecha_arca').
    Devuelve {pv: cantidad de cierres con fecha distinta}."""
    z_fecha = defaultdict(dict)
    for r in datos['zetti']:
        if r['tc'] == 'Z' and isinstance(r['numero'], int):
            z_fecha[r['punto']].setdefault(r['numero'], r['fecha'])
    out = defaultdict(int)
    for pv, zs in z_fecha.items():
        nums = {r['numero'] for f in ('mc', 'portal') for r in datos[f]
                if r['tipo'] == 83 and r['punto'] == pv and isinstance(r['numero'], int)}
        if not es_esquema_viejo(nums, zs):
            continue
        cambiados = set()
        for f in ('mc', 'portal'):
            for r in datos[f]:
                if r['tipo'] == 83 and r['punto'] == pv and r['numero'] in zs and r['fecha'] != zs[r['numero']]:
                    r['fecha_arca'] = r['fecha']
                    r['fecha'] = zs[r['numero']]
                    cambiados.add(r['numero'])
                elif r['tipo'] == 83 and r['punto'] == pv and isinstance(r['numero'], int) and r['numero'] not in zs:
                    r['z_otro_mes'] = 'anterior' if r['numero'] < min(zs) else 'siguiente' if r['numero'] > max(zs) else ''
        if cambiados:
            out[pv] = len(cambiados)
    return dict(out)


def cruzar(datos, tol):
    pv_z = {r['punto'] for r in datos['zetti'] if r['tc'] == 'Z'}
    tipos_pv = defaultdict(set)
    for f in ('mc', 'portal'):
        for r in datos[f]:
            tipos_pv[r['punto']].add(r['tipo'])
    pv_tiq = pv_z | {pv for pv, t in tipos_pv.items() if t and t <= TIQUES}

    renglones = OrderedDict()

    def renglon(clave, base):
        if clave not in renglones:
            renglones[clave] = {**base, 'fuentes': {}, 'clientes': {}}
        return renglones[clave]

    indice = defaultdict(list)  # (pv, nro) -> claves ARCA, por si Zetti trae otra letra
    for f in ('mc', 'portal'):
        for r in datos[f]:
            if r['punto'] in pv_tiq:
                continue
            clave = ('c', r['tipo'], r['punto'], r['numero'])
            rg = renglon(clave, {'modo': 'comp', 'fecha': r['fecha'], 'tipo': r['tipo'],
                                 'punto': r['punto'], 'numero': r['numero']})
            acc = rg['fuentes'].setdefault(f, {**vacio(), 'n': 0})
            sumar(acc, r); acc['n'] += 1
            if r['cliente']:
                rg['clientes'].setdefault(f, r['cliente'])
            if clave not in indice[(r['punto'], r['numero'])]:
                indice[(r['punto'], r['numero'])].append(clave)
    for r in datos['zetti']:
        if r['punto'] in pv_tiq:
            continue
        clave = ('c', r['tipo'], r['punto'], r['numero'])
        if clave not in renglones and len(indice.get((r['punto'], r['numero']), [])) == 1:
            clave = indice[(r['punto'], r['numero'])][0]
        rg = renglon(clave, {'modo': 'comp', 'fecha': r['fecha'], 'tipo': r['tipo'],
                             'punto': r['punto'], 'numero': r['numero'], 'tc_zetti': r['tc_zetti']})
        acc = rg['fuentes'].setdefault('zetti', {**vacio(), 'n': 0})
        sumar(acc, r); acc['n'] += 1
        if r['cliente']:
            rg['clientes'].setdefault('zetti', r['cliente'])

    for f in FUENTES:
        for r in datos[f]:
            if r['punto'] not in pv_tiq:
                continue
            rg = renglon(('z', r['fecha'], r['punto']), {'modo': 'tiques', 'fecha': r['fecha'], 'tipo': None,
                                                         'punto': r['punto'], 'numero': None})
            acc = rg['fuentes'].setdefault(f, {**vacio(), 'n': 0, 'rangos': {}, 'zs': []})
            sumar(acc, r); acc['n'] += 1
            if f == 'zetti':
                acc['zs'].append((r['numero'], r['total']))
            elif r.get('z_otro_mes') is not None:
                rg['z_otro_mes'] = (r['numero'], r['z_otro_mes'])
            lo, hi = r['numero'], r.get('hasta') or r['numero']
            if f != 'zetti' and isinstance(lo, int) and isinstance(hi, int):
                x = acc['rangos'].setdefault(r['tipo'], [lo, hi, 0])
                x[0], x[1], x[2] = min(x[0], lo), max(x[1], hi), x[2] + 1

    filas = list(renglones.values())
    for rg in filas:
        clasificar(rg, tol)
    marcar_corrimientos(filas, tol)
    marcar_corte_de_mes(filas)
    filas.sort(key=lambda x: (x['fecha'], 0 if x['modo'] == 'comp' else 1, x['punto'] or 0,
                              x['tipo'] or 0, str(x['numero'] or '')))
    return filas, sorted(pv_tiq)


def difiere(a, b, tol):
    return any(abs(r2(a[k] - b[k])) > tol for k in COMP)


def clasificar(rg, tol):
    fu = rg['fuentes']
    en = [f for f in FUENTES if f in fu]
    faltan = [f for f in FUENTES if f not in fu]
    v = rg['valores'] = {f: fu.get(f, vacio()) for f in FUENTES}
    rg['dif_mc_portal'] = r2(v['mc']['total'] - v['portal']['total'])
    rg['dif_zetti_portal'] = r2(v['zetti']['total'] - v['portal']['total'])
    rg['dif_zetti_mc'] = r2(v['zetti']['total'] - v['mc']['total'])
    distintas = [f for f in en if any(difiere(fu[f], fu[g], tol) for g in en if g != f)]
    tiq = rg['modo'] == 'tiques'
    nombres = lambda fs: ' y '.join(NOMBRE[f] for f in fs)  # noqa: E731

    if not faltan and not distintas:
        rg['estado'], rg['obs'] = 'OK', ''
        return
    if len(en) == 1:
        solo = en[0]
        rg['estado'] = f'Solo en {NOMBRE[solo]}'
        rg['obs'] = {
            'zetti': 'Z de Zetti sin tiques en ARCA ese dia' if tiq else
                     'Registrado en Zetti sin comprobante en ARCA: sin CAE, anulado o mal numerado en Zetti',
            'mc': 'Tiques en ARCA que no estan en el libro ni en Zetti' if tiq else
                  'Autorizado en ARCA pero no esta en el libro IVA ni en Zetti',
            'portal': 'Tiques en el libro que no estan en ARCA ni en Zetti' if tiq else
                      'Solo en el libro IVA: cargado a mano en el Portal',
        }[solo]
        return
    if faltan:
        rg['estado'] = f'Falta en {nombres(faltan)}'
        rg['obs'] = {
            'portal': 'Autorizado en ARCA y registrado en Zetti, pero no esta en el libro IVA: reimportar en el Portal',
            'mc': 'Esta en el libro y en Zetti pero no en Mis Comprobantes: cargado a mano en el Portal',
            'zetti': 'Tiques del dia en ARCA y en el libro sin Z en Zetti' if tiq else
                     'Emitido en ARCA y en el libro, pero no registrado en Zetti',
        }.get(faltan[0], '') if len(faltan) == 1 else ''
        if distintas:
            rg['obs'] += ' · ademas difieren los importes entre ' + nombres(en)
        return
    if len(distintas) == 3:
        for f in en:
            otras = [g for g in en if g != f]
            if not difiere(fu[otras[0]], fu[otras[1]], tol):
                distintas = [f]
                break
    rg['estado'] = f'Importe distinto en {NOMBRE[distintas[0]]}' if len(distintas) == 1 else 'Importes distintos'
    comps = [k for k in COMP if len({r2(fu[f][k]) for f in en}) > 1]
    dif_total = r2(max(fu[f]['total'] for f in en) - min(fu[f]['total'] for f in en))
    rg['obs'] = 'Difiere: ' + ', '.join({'exen': 'exento/no gravado', 'grav': 'gravado',
                                         'iva': 'IVA', 'total': 'total'}[k] for k in comps)
    if dif_total >= 1000 and dif_total % 1000 == 0:
        rg['obs'] += f' · diferencia redonda de $ {pesos(dif_total)}: posible error de carga'


def marcar_corrimientos(filas, tol):
    tiq = [f for f in filas if f['modo'] == 'tiques' and f['estado'] != 'OK']
    for f in tiq:
        f['_d'] = r2(f['valores']['zetti']['total'] - f['valores']['portal']['total'])
        if abs(f['_d']) <= tol * 4:
            f['obs'] = (f['obs'] + ' · ' if f['obs'] else '') + 'redondeo'
    for f in tiq:
        if abs(f['_d']) <= tol * 4:
            continue
        compensa = [g for g in tiq if g is not f and g['punto'] == f['punto'] and g['fecha'] != f['fecha']
                    and abs(g['_d'] + f['_d']) <= max(2.0, tol * 4) and abs((g['fecha'] - f['fecha']).days) <= 3]
        if compensa:
            f['estado'] = 'Corrimiento de fecha (tiques)'
            f['compensa'] = compensa
    # Corrimientos de mas de un dia o de varios cierres juntos (caso real: el Z del 21/02
    # cargado el 25/02; 202607: los del 18 y 19/07 cargados el 20/07). Un dia de Zetti de mas
    # compensa con 1 a 3 dias sueltos o con varios dias seguidos de Zetti de menos, dentro de 10 dias.
    from itertools import combinations
    libres = [f for f in tiq if not f.get('compensa') and abs(f['_d']) > tol * 4]
    for f in sorted(libres, key=lambda x: -abs(x['_d'])):
        if f.get('compensa'):
            continue
        cand = [g for g in libres if g is not f and not g.get('compensa') and g['punto'] == f['punto']
                and (g['_d'] > 0) != (f['_d'] > 0) and abs((g['fecha'] - f['fecha']).days) <= 10]
        cand.sort(key=lambda g: g['fecha'])
        ok = lambda c: abs(sum(g['_d'] for g in c) + f['_d']) <= max(2.0, tol * 4)  # noqa: E731
        # dias seguidos (caso real: del 10 al 17/07 cargados todos el 18/07), o hasta 3 sueltos
        tramos = [cand[i:j] for i in range(len(cand)) for j in range(i + 2, len(cand) + 1)]
        for k in (1, 2, 3, 0):
            hallado = (next((c for c in tramos if ok(c)), None) if k == 0 else
                       next((c for c in combinations(cand, k) if ok(c)), None))
            if hallado:
                f['estado'] = 'Corrimiento de fecha (tiques)'
                f['compensa'] = sorted(hallado, key=lambda g: g['fecha'])
                for g in hallado:
                    g['estado'] = 'Corrimiento de fecha (tiques)'
                    g['compensa'] = [f]
                break
    # Tramo de dias seguidos con diferencias de los dos signos que suman 0 (caso real:
    # del 15 al 18/02 Zetti reparte distinto entre los dias los mismos cierres).
    libres = sorted([f for f in tiq if not f.get('compensa') and abs(f['_d']) > tol * 4],
                    key=lambda x: (str(x['punto']), x['fecha']))
    for i in range(len(libres)):
        for j in range(i + 2, len(libres) + 1):
            tramo = libres[i:j]
            if any(g.get('compensa') for g in tramo) or len({g['punto'] for g in tramo}) > 1                     or (tramo[-1]['fecha'] - tramo[0]['fecha']).days > 10:
                break
            if abs(sum(g['_d'] for g in tramo)) <= max(2.0, tol * 4):
                for g in tramo:
                    g['estado'] = 'Corrimiento de fecha (tiques)'
                    g['compensa'] = [x for x in tramo if x is not g]
                    g['tramo'] = (tramo[0]['fecha'], tramo[-1]['fecha'])
                break


def marcar_corte_de_mes(filas):
    """Comprobante que le falta a Zetti (o que solo tiene Zetti) cuyo numero queda
    antes del primero / despues del ultimo que la otra parte tiene en el mes para
    ese tipo y PV: esta en el mes de al lado (Zetti fecha la venta, ARCA el CAE)."""
    rango = defaultdict(lambda: [None, None])
    for rg in filas:
        if rg['modo'] != 'comp' or not isinstance(rg['numero'], int):
            continue
        for f in rg['fuentes']:
            x = rango[(f, rg['tipo'], rg['punto'])]
            x[0] = rg['numero'] if x[0] is None else min(x[0], rg['numero'])
            x[1] = rg['numero'] if x[1] is None else max(x[1], rg['numero'])
    for rg in filas:
        if rg['modo'] != 'comp' or not isinstance(rg['numero'], int) or not rg['estado'].startswith(('Falta', 'Solo')):
            continue
        faltan = [f for f in FUENTES if f not in rg['fuentes']]
        if faltan != ['zetti'] and list(rg['fuentes']) != ['zetti']:
            continue
        lados = set()
        for f in faltan:
            lo, hi = rango.get((f, rg['tipo'], rg['punto']), (None, None))
            if lo is None:
                lados = {None}
                break
            lados.add('anterior' if rg['numero'] < lo else 'siguiente' if rg['numero'] > hi else None)
        if len(lados) == 1 and None not in lados:
            rg['estado'] = 'Corte de mes'
            rg['corte'] = (lados.pop(), faltan)


# ---------------------------------------------------------------------------
# Observaciones (mismo vocabulario que la solapa "cruze" del papel de trabajo)
# ---------------------------------------------------------------------------

def monto(x):
    x = abs(x)
    if x >= 1_000_000:
        return f'${x / 1_000_000:.1f} M'.replace('.', ',')
    return f'${pesos(x)}'


def ddmm(d):
    return d.strftime('%d/%m')


def nro_z(pv, nro):
    return f'{pv}-{nro}' if isinstance(nro, int) else str(nro)


def redactar_observaciones(filas, dif_mes, tol):
    COMPS = (('exen', 'Exento'), ('grav', 'Gravado'), ('iva', 'IVA'))
    # Z de Zetti con numero fuera de la secuencia del mes (caso real: "13-952" entre
    # el 1759 y el 1760, probable Z 1754 mal cargado)
    nros = defaultdict(list)
    for rg in filas:
        if rg['modo'] == 'tiques' and 'zetti' in rg['fuentes']:
            for n, t in rg['fuentes']['zetti'].get('zs', []):
                if isinstance(n, int):
                    nros[rg['punto']].append((n, rg['fecha'], t))
    fuera_seq = {}
    for pv, xs in nros.items():
        orden = sorted(n for n, _, _ in xs)
        med = orden[len(orden) // 2]
        raros = [x for x in xs if abs(x[0] - med) > max(200, 3 * len(orden))]
        if raros and len(raros) < len(orden) / 4:
            fuera_seq[pv] = (raros, orden)
    for rg in filas:
        if rg['estado'] == 'OK':
            continue
        v = rg['valores']
        z, p = v['zetti'], v['portal']
        d = rg['dif_zetti_portal']
        tiq = rg['modo'] == 'tiques'

        if rg.get('tramo'):
            d0, d1 = rg['tramo']
            rg['obs'] = (f"Corrimiento de fecha: entre el {ddmm(d0)} y el {ddmm(d1)} Zetti reparte los cierres Z en "
                         f"otros dias que Portal IVA (este dia {monto(d)} de {'mas' if d > 0 else 'menos'}); "
                         f"esos dias juntos suman lo mismo. No afecta el total del mes.")
            continue
        if rg.get('compensa'):
            otra = rg['compensa'][0]['fecha']
            if len(rg['compensa']) > 1 and d > 0:
                fechas = [ddmm(g['fecha']) for g in rg['compensa']]
                fechas_txt = ', '.join(fechas[:-1]) + ' y ' + fechas[-1]
                rg['obs'] = (f"Corrimiento de fecha: Zetti carga el {ddmm(rg['fecha'])} los cierres que Portal IVA "
                             f"registra el {fechas_txt} (diferencia {monto(d)}). Compensa con las filas del "
                             f"{fechas_txt}, no afecta el total del mes.")
                continue
            if d < 0:
                rg['obs'] = (f"Corrimiento de fecha: Portal IVA registra este cierre Z ({monto(d)}) el {ddmm(rg['fecha'])}; "
                             f"en Zetti esta cargado el {ddmm(otra)}. Compensa con la fila del {ddmm(otra)}, "
                             f"no afecta el total del mes.")
            elif p['total']:
                rg['obs'] = (f"Corrimiento de fecha: Zetti carga el cierre del {ddmm(otra)} sumado al del "
                             f"{ddmm(rg['fecha'])} (diferencia {monto(d)}). Compensa con la fila del {ddmm(otra)}.")
            else:
                rg['obs'] = (f"Corrimiento de fecha: Zetti carga el {ddmm(rg['fecha'])} el cierre que Portal IVA "
                             f"registra el {ddmm(otra)} ({monto(d)}). Compensa con la fila del {ddmm(otra)}.")
            continue

        if rg['estado'] == 'Corte de mes':
            lado, faltan = rg['corte']
            if faltan == ['zetti']:
                rg['obs'] = (f"Corte de mes: esta en Portal IVA y Mis Comprobantes con fecha {ddmm(rg['fecha'])} pero no en "
                             f"el Ventas Zetti del mes; por la numeracion corresponde al mes {lado}. Confirmarlo en la "
                             f"solapa Ventas Zetti del papel de trabajo del mes {lado}.")
            else:
                rg['obs'] = (f"Corte de mes: Zetti lo tiene el {ddmm(rg['fecha'])} pero ARCA lo informa en el mes {lado} "
                             f"(el numero queda {'antes del primero' if lado == 'anterior' else 'despues del ultimo'} "
                             f"del mes en ARCA). No es un faltante; revisar el libro del mes {lado}.")
            continue

        if not tiq and rg['estado'] == 'Falta en Zetti' and all(
                abs(v[f]['total']) <= tol for f in ('mc', 'portal') if f in rg['fuentes']):
            rg['obs'] = ("Comprobante en $0 en ARCA (" + ' y '.join(NOMBRE[f] for f in ('mc', 'portal') if f in rg['fuentes'])
                         + ") que no esta en Zetti: posible comprobante anulado o emitido en cero. No afecta los totales.")
            continue

        if abs(d) <= tol * 4 and not rg['dif_mc_portal']:
            rg['obs'] = ('Diferencia minima de redondeo en el cierre Z, sin relevancia' if tiq else
                         'Diferencia minima de redondeo entre Zetti y Portal IVA, sin relevancia')
            continue

        if tiq and 'mc' not in rg['fuentes'] and 'zetti' in rg['fuentes'] and 'portal' in rg['fuentes'] \
                and not difiere(z, p, tol):
            rg['obs'] = (f"Cierre Z del controlador fiscal ({monto(p['total'])}): esta en Zetti y en Portal IVA con el mismo "
                         f"importe, pero no en Mis Comprobantes, que hasta 2025 no informa los tiques del controlador "
                         f"fiscal. No es una diferencia.")
            continue

        if tiq and 'zetti' in rg['fuentes'] and ('portal' in rg['fuentes'] or 'mc' in rg['fuentes']):
            difs = [(n, r2(z[k] - p[k])) for k, n in COMPS if abs(z[k] - p[k]) > tol * 4]
            if abs(d) <= tol * 4:
                partes = '; '.join(f"{n}: Zetti ${pesos(z[k])}, Portal IVA ${pesos(p[k])}"
                                   for k, n in COMPS if abs(z[k] - p[k]) > tol * 4)
                zs = rg['fuentes']['zetti'].get('zs', [])
                con = list(dict.fromkeys(nro_z(rg['punto'], n) for n, t in zs if abs(t) > tol))
                rg['obs'] = (f"Mismo total que Portal IVA ({monto(p['total'])}) pero distinta composicion: {partes}. "
                             f"No cambia el total del mes, si el exento y el gravado. "
                             f"Revisar la carga del cierre Z {', '.join(con) or ''} en Zetti.")
                continue
            if len(difs) == 1:
                n, x = difs[0]
                redondo = abs(x) >= 1000 and abs(round(x)) % 1000 == 0
                txt = (f"Diferencia real sin compensar: Zetti tiene ${pesos(abs(round(x) if redondo else x))}"
                       f"{' exactos' if redondo else ''} de {'mas' if x > 0 else 'menos'} en {n} que Portal IVA.")
            else:
                txt = (f"Diferencia real sin compensar: Zetti tiene {monto(d)} de {'mas' if d > 0 else 'menos'} "
                       f"que Portal IVA en el cierre del dia ({', '.join(n for n, _ in difs) or 'Total'}).")
                redondo = abs(d) >= 1000 and abs(round(d)) % 1000 == 0
            if redondo:
                txt += ' Importe redondo, tipico de un ajuste o carga manual sin respaldo en un ticket Z.'
            if abs(d - dif_mes) <= tol * 4:
                txt += ' Es el origen de toda la diferencia del mes.'
            zs = rg['fuentes']['zetti'].get('zs', [])
            vistos = defaultdict(list)
            for n, t in zs:
                vistos[n].append(t)
            dup = [(n, ts[0]) for n, ts in vistos.items() if len(ts) > 1 and abs(ts[0]) > tol]
            if dup and abs(sum(t * (len(vistos[n]) - 1) for n, t in dup) - d) <= tol * 4:
                txt = (f"Diferencia real sin compensar: Zetti tiene cargado dos veces el cierre Z "
                       f"{', '.join(nro_z(rg['punto'], n) for n, _ in dup)} del {ddmm(rg['fecha'])} "
                       f"(mismo numero e importe, {monto(dup[0][1])}); la diferencia con Portal IVA es exactamente "
                       f"ese duplicado.")
                if abs(d - dif_mes) <= tol * 4:
                    txt += ' Es el origen de toda la diferencia del mes.'
                rg['obs'] = txt + ' Hay que eliminar el renglon repetido en Zetti.'
                continue
            rng = (rg['fuentes'].get('portal') or rg['fuentes'].get('mc') or {}).get('rangos', {}).get(83)
            if rng and rng[2] == rng[1] - rng[0] + 1 and all(isinstance(n, int) for n, _ in zs)                     and any(rng[0] <= n <= rng[1] for n, _ in zs):
                # esquema 2025 (ARCA numera por Z): Z de Zetti del dia que ARCA no informa
                sin_arca = [(n, t) for n, t in zs if not rng[0] <= n <= rng[1] and abs(t) > tol]
                if sin_arca and abs(sum(t for _, t in sin_arca) - d) <= tol * 4:
                    txt = (f"Diferencia real sin compensar: el cierre Z "
                           f"{', '.join(nro_z(rg['punto'], n) for n, _ in sin_arca)} de Zetti "
                           f"({monto(sum(t for _, t in sin_arca))}) no esta en Portal IVA; ARCA informa ese dia "
                           f"solo el {nro_z(rg['punto'], rng[0]) if rng[0] == rng[1] else f'{rng[0]}-{rng[1]}'}.")
                    if abs(d - dif_mes) <= tol * 4:
                        txt += ' Es el origen de toda la diferencia del mes.'
                    rg['obs'] = txt
                    continue
            con = list(dict.fromkeys(nro_z(rg['punto'], n) for n, t in zs if abs(t) > tol))
            cero = [nro_z(rg['punto'], n) for n, t in zs if abs(t) <= tol]
            raros = [n for n, _, _ in fuera_seq.get(rg['punto'], ([], []))[0] if n in {m for m, _ in zs}]
            if raros:
                orden = fuera_seq[rg['punto']][1]
                txt += (f" El cierre Z {', '.join(nro_z(rg['punto'], n) for n in raros)} tiene numero fuera de secuencia "
                        f"(los Z del mes van del {orden[0] if orden[0] not in raros else min(x for x in orden if x not in raros)} "
                        f"al {max(x for x in orden if x not in raros)}): posible Z de otro dia mal cargado en Zetti.")
            if con:
                txt += f" Revisar el cierre Z {', '.join(con)}"
                txt += (f" (y el {', '.join(cero)}, en cero, del mismo dia)." if cero else '.')
            rg['obs'] = txt
            continue

        if tiq:
            solo = rg['estado']
            if 'zetti' not in rg['fuentes'] and rg.get('z_otro_mes'):
                nz, lado = rg['z_otro_mes']
                rg['obs'] = (f"Diferencia real sin compensar: Portal IVA informa el {ddmm(rg['fecha'])} el cierre Z "
                             f"{nro_z(rg['punto'], nz)} ({monto(p['total'] or v['mc']['total'])}), que no es un Z de Zetti "
                             f"de este mes" + (f": el numero es del mes {lado}, donde Zetti ya lo tiene (ver el listado "
                                               f"de comprobantes de Zetti que figuran en otro periodo)." if lado else '.')
                             + " Posible cierre Z informado dos veces en ARCA.")
            elif 'zetti' not in rg['fuentes']:
                rg['obs'] = (f"Diferencia real sin compensar: Portal IVA registra tiques por {monto(p['total'] or v['mc']['total'])} "
                             f"el {ddmm(rg['fecha'])} y Zetti no tiene cierre Z ese dia ni otro que lo compense.")
                if rg['punto'] in fuera_seq:
                    raros, orden = fuera_seq[rg['punto']]
                    rg['obs'] += (' Zetti tiene en el mes un cierre Z con numero fuera de secuencia: '
                                  + ', '.join(f"{nro_z(rg['punto'], n)} del {ddmm(f)} ({monto(t)})" for n, f, t in raros)
                                  + '; puede ser el Z de este dia mal cargado.')
            elif solo.startswith('Solo en Zetti'):
                rg['obs'] = (f"Diferencia real sin compensar: Zetti tiene un cierre Z por {monto(z['total'])} el "
                             f"{ddmm(rg['fecha'])} sin tiques en ARCA ese dia ni otro que lo compense.")
            continue

        # comprobantes con numero propio
        if rg['estado'] == 'Solo en Zetti':
            rg['obs'] = (f"Solo en Zetti: registrado en Zetti por {monto(z['total'])} sin comprobante en ARCA "
                         f"(no esta en Mis Comprobantes ni en Portal IVA). Puede no tener CAE, estar anulado o mal "
                         f"numerado en Zetti; revisarlo en el sistema.")
        elif rg['estado'] == 'Falta en Zetti':
            rg['obs'] = (f"Falta en Zetti: emitido en ARCA y en el libro IVA por {monto(p['total'] or v['mc']['total'])}, "
                         f"pero no esta registrado en Zetti (facturado por fuera del sistema: RCEL u otro).")
        elif rg['estado'] == 'Importe distinto en Zetti':
            partes = [f"{n} {monto(z[k] - p[k])}" for k, n in COMPS + (('total', 'Total'),) if abs(z[k] - p[k]) > tol]
            rg['obs'] = f"Mismo comprobante con distinto importe en Zetti y Portal IVA ({', '.join(partes)})."


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def concepto(rg):
    if rg['modo'] == 'tiques':
        rangos = next((v['rangos'] for f, v in rg['fuentes'].items() if f != 'zetti' and v.get('rangos')), {})
        n = {f: v['n'] for f, v in rg['fuentes'].items()}
        txt = f"Tiques PV {rg['punto']}"
        if rangos:
            tipo, (lo, hi, _) = max(rangos.items(), key=lambda x: x[1][2])
            txt += f" ({TIPOS.get(tipo, tipo)} {lo}-{hi}" + (f' y {len(rangos) - 1} tipo(s) mas' if len(rangos) > 1 else '') + ')'
        partes = []
        if 'mc' in n or 'portal' in n:
            partes.append(f"{max(n.get('mc', 0), n.get('portal', 0))} tiques")
        if 'zetti' in n:
            partes.append(f"{n['zetti']} Z")
        return txt + (' · ' + ', '.join(partes) if partes else '')
    t = TIPOS.get(rg['tipo']) or rg.get('tc_zetti') or ''
    if isinstance(rg['punto'], int) and isinstance(rg['numero'], int):
        return f"{t} {rg['punto']:05d}-{rg['numero']:08d}"
    return f"{t} {rg['punto']}-{rg['numero']}"


def encabezado(ws, fila, textos):
    for j, h in enumerate(textos, start=1):
        c = ws.cell(fila, j, h)
        c.font, c.fill = F_H, FILL_H
        c.alignment = Alignment(horizontal='center', wrap_text=True)


def totales(datos):
    out = {}
    for f, filas in datos.items():
        g = OrderedDict((k, {**vacio(), 'n': 0}) for k in ('Facturas y ND', 'Notas de credito', 'Tiques'))
        tot = {**vacio(), 'n': 0}
        for r in filas:
            if f == 'zetti':
                k = 'Tiques' if r['tc'] == 'Z' else 'Notas de credito' if r['tc'] == 'NC' else 'Facturas y ND'
            else:
                k = 'Tiques' if r['tipo'] in TIQUES else 'Notas de credito' if r['tipo'] in NOTAS_CREDITO else 'Facturas y ND'
            sumar(g[k], r); g[k]['n'] += 1
            sumar(tot, r); tot['n'] += 1
        out[f] = {'grupos': g, 'total': tot}
    return out


def bloque_resumen(ws, titulo, archivos, tot, filas, fuera, avisos):
    ws.cell(1, 1, titulo).font = F_TIT
    r = 2
    for f in FUENTES:
        ws.cell(r, 1, f'{NOMBRE[f]}: {archivos[f]}').font = F_N
        r += 1
    r += 1
    ws.cell(r, 1, '1. Totales de cada fuente (NC en negativo)').font = F_B
    r += 1
    encabezado(ws, r, ['Fuente', 'Grupo', 'Cantidad', 'Exento/No grav.', 'Gravado', 'IVA', 'Total'])
    r += 1
    for f in FUENTES:
        for gname, v in list(tot[f]['grupos'].items()) + [('TOTAL', tot[f]['total'])]:
            ws.cell(r, 1, NOMBRE[f]).fill = FILL[f]
            ws.cell(r, 2, gname)
            ws.cell(r, 3, v['n'])
            for k, comp in enumerate(COMP):
                ws.cell(r, 4 + k, r2(v[comp])).number_format = FMT
            for j in range(1, 8):
                ws.cell(r, j).font = F_B if gname == 'TOTAL' else F_N
            r += 1
    if fuera:
        s = vacio()
        for x in fuera:
            sumar(s, x)
        pvs = ', '.join(str(p) for p in sorted({x['punto'] for x in fuera}, key=str))
        ws.cell(r, 1, 'Zetti').fill = FILL['zetti']
        ws.cell(r, 2, f'Eliminados: PV {pvs} (no informado a ARCA)')
        ws.cell(r, 3, len(fuera))
        for k, comp in enumerate(COMP):
            ws.cell(r, 4 + k, r2(s[comp])).number_format = FMT
        for j in range(1, 8):
            ws.cell(r, j).font = F_N
        r += 1
    r += 1

    ws.cell(r, 1, '2. Diferencias de total entre fuentes').font = F_B
    r += 1
    encabezado(ws, r, ['Comparacion', '', '', 'Exento/No grav.', 'Gravado', 'IVA', 'Total'])
    r += 1
    for a, b in (('mc', 'portal'), ('zetti', 'portal'), ('zetti', 'mc')):
        ws.cell(r, 1, f'{NOMBRE[a]} - {NOMBRE[b]}')
        if a == 'zetti' and fuera:
            ws.cell(r, 2, '(Zetti depurado)')
        for k, comp in enumerate(COMP):
            ws.cell(r, 4 + k, r2(tot[a]['total'][comp] - tot[b]['total'][comp])).number_format = FMT
        for j in range(1, 8):
            ws.cell(r, j).font = F_N
        r += 1
    r += 1

    ws.cell(r, 1, '3. Que explica las diferencias (renglones del punto 4)').font = F_B
    r += 1
    encabezado(ws, r, ['Estado', 'Renglones', '', 'MC - Portal', 'Zetti - Portal', 'Zetti - MC'])
    r += 1
    por_estado = OrderedDict()
    for rg in filas:
        if rg['estado'] == 'OK':
            continue
        e = por_estado.setdefault(rg['estado'], {'n': 0, 'dif_mc_portal': 0.0, 'dif_zetti_portal': 0.0, 'dif_zetti_mc': 0.0})
        e['n'] += 1
        for k in ('dif_mc_portal', 'dif_zetti_portal', 'dif_zetti_mc'):
            e[k] += rg[k]
    for estado, e in sorted(por_estado.items(), key=lambda x: -x[1]['n']):
        ws.cell(r, 1, estado)
        ws.cell(r, 2, e['n'])
        for k, clave in enumerate(('dif_mc_portal', 'dif_zetti_portal', 'dif_zetti_mc')):
            ws.cell(r, 4 + k, r2(e[clave])).number_format = FMT
        for j in range(1, 8):
            ws.cell(r, j).font = F_N
        r += 1
    if not por_estado:
        ws.cell(r, 1, 'Sin diferencias: las fuentes coinciden comprobante por comprobante.').font = F_N
        r += 1
    ok = sum(1 for rg in filas if rg['estado'] == 'OK')
    ws.cell(r, 1, f'OK (coinciden): {ok} renglones').font = F_N
    r += 2
    if avisos:
        ws.cell(r, 1, 'Avisos').font = F_B
        r += 1
        for a in avisos:
            ws.cell(r, 1, a).font = F_N
            r += 1
    return r + 1


def bloque_diferencias(ws, r0, filas):
    """Punto 4: solo los renglones con diferencia. Encabezado en r0+1 y r0+2."""
    ws.cell(r0, 1, '4. Detalle de las diferencias').font = F_B
    g, h = r0 + 1, r0 + 2
    fijas = ['Fecha', 'Comprobante', 'Cliente', 'Estado', 'Observacion']
    for j, t in enumerate(fijas, start=1):
        c = ws.cell(h, j, t)
        c.font, c.fill = F_H, FILL_H
        ws.cell(g, j).fill = FILL_H
    col = {}
    for i, f in enumerate(FUENTES):
        col[f] = 6 + 4 * i
        ws.merge_cells(start_row=g, start_column=col[f], end_row=g, end_column=col[f] + 3)
        c = ws.cell(g, col[f], NOMBRE[f].upper())
        c.font, c.fill, c.alignment = F_B, FILL[f], Alignment(horizontal='center')
        for k, t in enumerate(('Exento/No grav.', 'Gravado', 'IVA', 'Total')):
            c = ws.cell(h, col[f] + k, t)
            c.font, c.fill, c.alignment = F_B, FILL[f], Alignment(horizontal='center', wrap_text=True)
    difs = (('dif_mc_portal', 'MC - Portal'), ('dif_zetti_portal', 'Zetti - Portal'), ('dif_zetti_mc', 'Zetti - MC'))
    cd = 18
    ws.merge_cells(start_row=g, start_column=cd, end_row=g, end_column=cd + 2)
    c = ws.cell(g, cd, 'DIFERENCIA DE TOTAL')
    c.font, c.fill, c.alignment = F_B, FILL['dif'], Alignment(horizontal='center')
    for k, (_, t) in enumerate(difs):
        c = ws.cell(h, cd + k, t)
        c.font, c.fill, c.alignment = F_B, FILL['dif'], Alignment(horizontal='center')
    ult = cd + 2

    r = d0 = h + 1
    for rg in (x for x in filas if x['estado'] != 'OK'):
        ws.cell(r, 1, rg['fecha']).number_format = FMT_F
        ws.cell(r, 2, concepto(rg))
        cl = rg['clientes']
        ws.cell(r, 3, cl.get('mc') or cl.get('portal') or cl.get('zetti') or '')
        ws.cell(r, 4, rg['estado']).fill = FILL_ROJO
        ws.cell(r, 5, rg['obs'])
        for f in FUENTES:
            if f in rg['fuentes']:
                for k, comp in enumerate(COMP):
                    ws.cell(r, col[f] + k, r2(rg['valores'][f][comp])).number_format = FMT
        for k, (clave, _) in enumerate(difs):
            ws.cell(r, cd + k, rg[clave]).number_format = FMT
        for j in range(1, ult + 1):
            ws.cell(r, j).font = F_N
        r += 1
    ws.cell(r, 1, 'TOTAL').font = F_B
    for j in range(6, ult + 1):
        L = get_column_letter(j)
        c = ws.cell(r, j, f'=SUBTOTAL(9,{L}{d0}:{L}{r - 1})' if r > d0 else 0)
        c.number_format, c.font = FMT, F_B
    if r > d0:
        ws.auto_filter.ref = f'A{h}:{get_column_letter(ult)}{r - 1}'
    return r + 3


def rangos_txt(nums, maximo=25):
    """[1,2,3,7,9,10] -> '1-3, 7, 9-10' (corta si son muchos tramos)."""
    tramos = []
    for n in sorted(nums):
        if tramos and n == tramos[-1][1] + 1:
            tramos[-1][1] = n
        else:
            tramos.append([n, n])
    txt = [f'{a}' if a == b else f'{a}-{b}' for a, b in tramos]
    return ', '.join(txt[:maximo]) + (f' y {len(txt) - maximo} tramos mas' if len(txt) > maximo else '')


def control_numeros(nums):
    """Desde, hasta, cantidad, faltantes dentro del rango y repetidos."""
    if not nums:
        return None
    vistos, rep = set(), set()
    for n in nums:
        (rep if n in vistos else vistos).add(n)
    lo, hi = min(vistos), max(vistos)
    return {'desde': lo, 'hasta': hi, 'n': len(nums),
            'faltan': sorted(set(range(lo, hi + 1)) - vistos), 'rep': sorted(rep)}


def correlatividad(datos, ult_ti_anterior=None, estado_dia=None):
    """Un renglon por (tipo, PV) con el control de numeracion de cada fuente.
    ARCA (MC y Portal): cada comprobante por su numero; los renglones desde-hasta se
    expanden. Zetti (depurado): FC/NC/ND por numero; los tiques no vienen uno por uno,
    asi que se controla la numeracion de los cierres Z y que el ultimo tique de cada Z
    ('TI nnn') no retroceda; el ultimo tique de Zetti se compara con el de ARCA."""
    nums = defaultdict(lambda: defaultdict(list))
    for f in ('mc', 'portal'):
        for r in datos[f]:
            lo, hi = r['numero'], r.get('hasta') or r['numero']
            if isinstance(lo, int) and isinstance(hi, int):
                nums[(r['tipo'], r['punto'])][f].extend(range(lo, max(lo, hi) + 1))
    z_nums, z_ti, z_sin_ti, z_cero = defaultdict(list), defaultdict(list), defaultdict(list), defaultdict(list)
    for r in datos['zetti']:
        if r['tc'] == 'Z':
            if isinstance(r['numero'], int):
                z_nums[r['punto']].append(r['numero'])
                if abs(r['total']) <= 0.5:
                    z_cero[r['punto']].append(r['numero'])
            m = re.search(r'TI\s*(\d+)\s*-\s*(\d+)', r['cliente'])
            if m:
                z_ti[r['punto']].append((r['fecha'], r['numero'], int(m.group(2))))
            else:
                z_sin_ti[r['punto']].append(r['numero'])
        elif isinstance(r['numero'], int):
            nums[(r['tipo'] or r['tc_zetti'], r['punto'])]['zetti'].append(r['numero'])
    estado_dia = estado_dia or {}
    tiq_arca = defaultdict(dict)  # pv -> numero de tique -> {fecha, total}
    for f in ('mc', 'portal'):
        for r in datos[f]:
            if r['tipo'] == 83 and isinstance(r['numero'], int):
                tiq_arca[r['punto']].setdefault(r['numero'], {'fecha': r['fecha'], 'total': r['total']})

    def orden(k):
        tipo, pv = k
        t = tipo if isinstance(tipo, int) else 999
        g = 2 if t in TIQUES else 1 if t in NOTAS_CREDITO else 0
        return (g, t, str(pv))

    renglones = []
    for clave in sorted(nums, key=orden):
        tipo, pv = clave
        ctl = {f: control_numeros(nums[clave].get(f, [])) for f in FUENTES}
        obs = []
        info = []      # aclaraciones que no son un problema (no pasan el renglon a Revisar)
        rg_falt = None
        viejo = False
        arca = {f: set(nums[clave].get(f, [])) for f in ('mc', 'portal')}
        if tipo == 83 and z_nums.get(pv) and any(arca.values()) \
                and all(es_esquema_viejo(v, z_nums[pv]) for v in arca.values() if v):
            # Esquema viejo (hasta 2025): ARCA informa un renglon por cierre Z con el numero
            # del Z, no los tiques. Se compara contra los Z de Zetti; los Z en cero no se informan.
            viejo = True
            ctl['zetti'] = control_numeros(z_nums[pv])
            cero = set(z_cero.get(pv, []))
            con_importe = set(z_nums[pv]) - cero
            for f in ('mc', 'portal'):
                if ctl[f]:
                    ctl[f]['faltan'] = [n for n in ctl[f]['faltan'] if n not in cero]
            info.append('ARCA informa un renglon por cierre Z (numero del Z, no de los tiques): se compara contra los Z de Zetti')
            if cero:
                info.append(f"Z en cero en Zetti, no informados a ARCA: {rangos_txt(sorted(cero))}")
            for f in ('mc', 'portal'):
                if arca[f] and arca[f] != con_importe:
                    sobran, faltan = sorted(arca[f] - con_importe), sorted(con_importe - arca[f])
                    obs.append(f"{NOMBRE[f]}: los cierres Z no coinciden con los de Zetti"
                               + (f"; faltan {rangos_txt(faltan)}" if faltan else '')
                               + (f"; de mas {rangos_txt(sobran)}" if sobran else ''))
        elif tipo == 83 and z_ti.get(pv):
            # Zetti no lista los tiques: cada Z trae solo el ultimo ("TI nnn"). En el cuadro va
            # solo lo que sale del Zetti del mes (el Hasta); el primer tique se aclara en la
            # observacion con el ultimo Z del mes anterior.
            ult = max(t for _, _, t in z_ti[pv])
            z1 = min(z_ti[pv], key=lambda x: x[2])
            # Desde = el primer numero que muestra el Zetti del mes: el ultimo tique del primer Z
            ctl['zetti'] = {'desde': z1[2], 'hasta': ult, 'n': None, 'faltan': [], 'rep': []}
            ref = ctl['portal'] or ctl['mc']
            ant = (ult_ti_anterior or {}).get(pv)
            txt = (f"Zetti: el Desde ({z1[2]}) es el ultimo tique del primer cierre Z del mes ({pv}-{z1[1]} del "
                   f"{ddmm(z1[0])}); Zetti solo muestra el ultimo numero de cada Z")
            if ref and ref['desde'] <= z1[2]:
                txt += (f". El {ref['desde']} (primer tique de ARCA) es el primer tique de ese mismo Z {pv}-{z1[1]}: "
                        f"esta dentro del Zetti del periodo, pero el Z solo muestra su ultimo numero ({z1[2]})")
            elif ref and ant and ant[1] < z1[2] < ref['desde']:
                # Caso real: el Z 13-1662 (tiques 258722 a 258732) no esta en ARCA
                obs.append(f"El primer cierre Z del mes ({pv}-{z1[1]} del {ddmm(z1[0])}, tiques {ant[1] + 1} a {z1[2]}) "
                           f"no esta en ARCA: el primer tique de ARCA es el {ref['desde']}, del Z siguiente")
            if ant and ref and not (ant[1] < z1[2] < ref['desde']):
                if ref['desde'] == ant[1] + 1:
                    txt += (f" (el ultimo Z del mes anterior, {pv}-{ant[0]}, termina en el {ant[1]})")
                elif ref['desde'] <= ant[1]:
                    txt += (f". El ultimo Z del mes anterior ({pv}-{ant[0]}) llega hasta el tique {ant[1]}: los tiques "
                            f"{ref['desde']} a {ant[1]} de ARCA Zetti los tiene en el periodo anterior (ver listado abajo)")
                else:
                    obs.append(f"Entre el ultimo Z del mes anterior ({pv}-{ant[0]}, hasta {ant[1]}) y el primer tique "
                               f"de ARCA ({ref['desde']}) faltan los tiques {ant[1] + 1} a {ref['desde'] - 1}")
            elif not ant:
                txt += '; no esta el papel del mes anterior para confirmar donde empieza'
            info.append(txt)
            if ref and ult != ref['hasta']:
                obs.append(f"Ultimo tique de Zetti ({ult}) distinto del de ARCA ({ref['hasta']})")
            # Cuadro de los tiques que ARCA no tiene: vecinos y cierre Z de Zetti que los incluye
            falt = sorted(set((ctl['portal'] or {}).get('faltan', [])) | set((ctl['mc'] or {}).get('faltan', [])))
            if falt:
                zs = sorted(z_ti[pv], key=lambda x: x[2])
                faltantes = []
                for n in falt:
                    a = max((k for k in tiq_arca[pv] if k < n), default=None)
                    s = min((k for k in tiq_arca[pv] if k > n), default=None)
                    fa, fs = (tiq_arca[pv].get(a) or {}).get('fecha'), (tiq_arca[pv].get(s) or {}).get('fecha')
                    z = next((x for x in zs if x[2] >= n), None)
                    faltantes.append({
                        'n': n, 'fecha': fa if fa == fs else None,
                        'fecha_txt': ddmm(fa) if fa == fs else f"entre {ddmm(fa) if fa else '?'} y {ddmm(fs) if fs else '?'}",
                        'ant': a, 'sig': s,
                        'z': f'{pv}-{z[1]} ({ddmm(z[0])}, TI hasta {z[2]})' if z else '',
                        'dia': estado_dia.get((fa, pv), 'OK') if fa == fs else '',
                    })
                rg_falt = faltantes
        for clave_ctl, texto in (('faltan', 'faltan'), ('rep', 'repetidos')):
            grupos = OrderedDict()  # mismas faltas en varias fuentes -> una sola observacion
            for f in FUENTES:
                if ctl[f] and ctl[f][clave_ctl]:
                    grupos.setdefault(tuple(ctl[f][clave_ctl]), []).append(NOMBRE[f])
            for lista, fuentes in grupos.items():
                obs.append(f"{' y '.join(fuentes)}: {texto} {len(lista)} ({rangos_txt(lista)})")
        presentes = {f: c for f, c in ctl.items() if c and c['n'] is not None}
        if tipo == 83 and not viejo:
            presentes.pop('zetti', None)  # Zetti cubre rangos de tiques: ya se controlo desde/hasta arriba
        if not viejo and len({(c['desde'], c['hasta'], c['n']) for c in presentes.values()}) > 1:
            obs.append('Desde/hasta/cantidad no coinciden entre fuentes (ver punto 4)')
        nombre = TIPOS.get(tipo, tipo) if isinstance(tipo, int) else f'{tipo} (Zetti, sin tipo ARCA)'
        if viejo:
            nombre += ' (por cierre Z)'
        renglones.append({'comp': nombre, 'pv': pv, 'ctl': ctl, 'obs': obs, 'info': info, 'faltantes': rg_falt})

    for pv, zs in sorted(z_nums.items(), key=lambda x: str(x[0])):
        c = control_numeros(zs)
        obs = []
        if c['faltan']:
            obs.append(f"Zetti: faltan {len(c['faltan'])} cierres Z ({rangos_txt(c['faltan'])})")
        if c['rep']:
            obs.append(f"Zetti: cierres Z repetidos {rangos_txt(c['rep'])}")
        ti = sorted(z_ti.get(pv, []), key=lambda x: x[1])
        atras = [b[1] for a, b in zip(ti, ti[1:]) if b[2] < a[2]]
        if atras:
            obs.append(f"El ultimo tique retrocede en los Z {rangos_txt(atras)}")
        info = []
        cero = set(z_cero.get(pv, []))
        sin_ti = [n for n in z_sin_ti.get(pv, []) if n not in cero]
        if sin_ti:
            obs.append(f"Z sin rango de tiques: {rangos_txt(sin_ti)}")
        if cero:
            info.append(f"Z en cero (sin ventas): {rangos_txt(sorted(cero))}")
        renglones.append({'comp': 'Cierre Z (solo Zetti)', 'pv': pv,
                          'ctl': {'mc': None, 'portal': None, 'zetti': c}, 'obs': obs, 'info': info})
    return renglones


def bloque_correlatividad(ws, r0, renglones):
    ws.cell(r0, 1, '5. Correlatividad de comprobantes (FC, NC, ND y tiques) por fuente').font = F_B
    ws.cell(r0 + 1, 1, 'Zetti ya depurado. Faltan = numeros que no estan dentro del rango desde-hasta de esa fuente.').font = F_N
    g, h = r0 + 2, r0 + 3
    for j, t in enumerate(['Comprobante', 'PV', 'Estado'], start=1):
        c = ws.cell(h, j, t)
        c.font, c.fill = F_H, FILL_H
        ws.cell(g, j).fill = FILL_H
    col = {}
    for i, f in enumerate(FUENTES):
        col[f] = 4 + 4 * i
        ws.merge_cells(start_row=g, start_column=col[f], end_row=g, end_column=col[f] + 3)
        c = ws.cell(g, col[f], NOMBRE[f].upper())
        c.font, c.fill, c.alignment = F_B, FILL[f], Alignment(horizontal='center')
        for k, t in enumerate(('Desde', 'Hasta', 'Cantidad', 'Faltan')):
            c = ws.cell(h, col[f] + k, t)
            c.font, c.fill, c.alignment = F_B, FILL[f], Alignment(horizontal='center')
    co = 16
    c = ws.cell(h, co, 'Observacion')
    c.font, c.fill = F_H, FILL_H
    ws.cell(g, co).fill = FILL_H
    r = h + 1
    for rg in renglones:
        ws.cell(r, 1, rg['comp'])
        ws.cell(r, 2, rg['pv'])
        ce = ws.cell(r, 3, 'Revisar' if rg['obs'] else 'OK')
        if rg['obs']:
            ce.fill = FILL_ROJO
        for f in FUENTES:
            c = rg['ctl'][f]
            if c:
                for k, v in enumerate((c['desde'], c['hasta'], c['n'], len(c['faltan']) if c['n'] is not None else None)):
                    ws.cell(r, col[f] + k, v).number_format = '0'
        txt = rg['obs'] + rg.get('info', [])
        if rg.get('faltantes'):
            txt.append('detalle de los faltantes en el cuadro de la columna U')
        ws.cell(r, co, ' · '.join(txt))
        for j in range(1, co + 1):
            ws.cell(r, j).font = F_N
        if rg.get('faltantes'):
            cuadro_faltantes(ws, r, rg['faltantes'])
        r += 1
    if not renglones:
        ws.cell(r, 1, 'Sin comprobantes para controlar.').font = F_N
        r += 1
    return r + 3


CU = 21  # columna U: cuadro de tiques faltantes, en la fila del renglon de correlatividad


def cuadro_faltantes(ws, r0, faltantes):
    """Tiques que faltan en ARCA dentro del rango del mes. No estan en ningun listado (ni
    Mis Comprobantes ni Portal IVA, y Zetti solo trae cierres Z), asi que se ubican por los
    tiques vecinos y por el cierre Z de Zetti que cubre ese numero."""
    cab = ['Tique faltante', 'Fecha (por tiques vecinos)', 'Tique anterior', 'Tique siguiente',
           'Cierre Z de Zetti que lo incluye', 'Cierre del dia en el cruce', 'Figura en algun listado']
    for j, t in enumerate(cab):
        c = ws.cell(r0, CU + j, t)
        c.font, c.fill = F_H, FILL_H
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    r = r0 + 1
    for x in faltantes:
        vals = [x['n'], x['fecha'] or x['fecha_txt'], x['ant'], x['sig'], x['z'],
                x['dia'] or '-', 'No (ni Mis Comprobantes, ni Portal IVA; Zetti solo trae el Z)']
        for j, v in enumerate(vals):
            c = ws.cell(r, CU + j, v)
            c.font = F_N
            c.number_format = FMT_F if j == 1 and x['fecha'] else '0'
        r += 1
    ok = all(x['dia'] in ('OK', 'Corrimiento de fecha (tiques)') for x in faltantes)
    nota = ('Los cierres Z de esos dias coinciden con ARCA (o solo tienen corrimiento de fecha): los tiques '
            'faltantes no tienen importe. Posiblemente anulados en el controlador fiscal; revisar la memoria '
            'fiscal / cinta testigo.' if ok else
            'Algun dia con tiques faltantes tiene diferencia en el cruce: revisar si el importe del tique falta en ARCA.')
    c = ws.cell(r, CU, nota)
    c.font = Font(name='Arial', size=10, italic=True)
    for j, w in enumerate([12, 14, 12, 12, 30, 26, 44]):
        ws.column_dimensions[get_column_letter(CU + j)].width = w


def bloque_eliminados(ws, r0, fuera):
    """Listado de los comprobantes que se sacaron de Zetti antes de cruzar."""
    c = ws.cell(r0, 1, 'comprobantes de Zetti ELIMINADOS')
    c.font = F_TIT
    h = r0 + 1
    for j, t in enumerate(['Fecha', 'Comprobante', 'Cliente', 'Exento/No grav.', 'Gravado', 'IVA', 'Total'], start=1):
        c = ws.cell(h, j, t)
        c.font, c.fill = F_H, FILL_H
        c.alignment = Alignment(horizontal='center')
    r = h + 1
    for x in sorted(fuera, key=lambda x: (x['fecha'], str(x['punto']), x['tc_zetti'], str(x['numero']))):
        pv, nro = x['punto'], x['numero']
        nro_txt = f'{pv:05d}-{nro:08d}' if isinstance(pv, int) and isinstance(nro, int) else f'{pv}-{nro}'
        ws.cell(r, 1, x['fecha']).number_format = FMT_F
        ws.cell(r, 2, f"{x['tc_zetti']} {nro_txt}")
        ws.cell(r, 3, x['cliente'])
        for k, comp in enumerate(COMP):
            ws.cell(r, 4 + k, r2(x[comp])).number_format = FMT
        for j in range(1, 8):
            ws.cell(r, j).font = F_N
        r += 1
    ws.cell(r, 1, 'TOTAL').font = F_B
    ws.cell(r, 3, f'{len(fuera)} comprobantes').font = F_B
    for j in range(4, 8):
        L = get_column_letter(j)
        c = ws.cell(r, j, f'=SUM({L}{h + 1}:{L}{r - 1})' if r > h + 1 else 0)
        c.number_format, c.font = FMT, F_B


# ---------------------------------------------------------------------------

HOJAS_PAPEL = {'mc': 'Ventas MCR', 'portal': 'Ventas portal iva', 'zetti': 'Ventas Zetti'}
HOJA_CRUCE = 'cruce ventas'

MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
         'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']


def main():
    pos = [a for a in sys.argv[1:] if not a.startswith('--')]
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel = opts.get('papel')
    if bool(papel) == bool(pos) or len(pos) > 1:
        sys.exit(__doc__)
    tol = float(opts.get('tolerancia') or 0.5)
    exc = excepciones(opts.get('excepcion'))
    if papel:
        # modo papel de trabajo: lee las solapas y escribe la solapa "cruce ventas" en el mismo libro
        mes = os.path.dirname(os.path.abspath(papel))
        wb = load_workbook(papel)
        hojas = {}
        for f, nombre in HOJAS_PAPEL.items():
            hoja = next((h for h in wb.sheetnames if norm(h) == norm(nombre)), None)
            if hoja is None:
                sys.exit(json.dumps({'error': f'Falta la solapa "{nombre}" en el papel', 'solapas': wb.sheetnames}, ensure_ascii=False))
            hojas[f] = wb[hoja]
        archivos = {f: f'solapa "{hojas[f].title}"' for f in FUENTES}
    else:
        carpeta = pos[0].rstrip('/\\')
        mes = os.path.dirname(carpeta)
        per = os.path.basename(mes)
        rutas = {
            'mc': opts.get('mc') or os.path.join(carpeta, f'{per} - MIS COMPROBANTES - EMITIDOS.xlsx'),
            'portal': opts.get('portal') or os.path.join(carpeta, f'{per} - PORTAL IVA - VENTAS.xlsx'),
            'zetti': opts.get('zetti') or os.path.join(carpeta, f'{per} - ZETTI - VENTAS.xlsx'),
        }
        for f, p in rutas.items():
            if not os.path.exists(p):
                sys.exit(json.dumps({'error': f'No encontre {NOMBRE[f]}: {p}'}, ensure_ascii=False))
        hojas = {f: load_workbook(p, read_only=True, data_only=True).worksheets[0] for f, p in rutas.items()}
        archivos = {f: os.path.basename(p) for f, p in rutas.items()}
    periodo = os.path.basename(mes)
    if not re.fullmatch(r'\d{6}', periodo):
        sys.exit(json.dumps({'error': f'No reconozco el periodo AAAAMM en la ruta: {mes}'}))

    datos = {'mc': leer_arca(hojas['mc'], 'mc'), 'portal': leer_arca(hojas['portal'], 'portal')}
    zetti, razon = leer_zetti(hojas['zetti'])
    datos['zetti'], fuera = depurar_zetti(zetti, datos['portal'], exc)

    redatados = alinear_fechas_esquema_viejo(datos)
    filas, pv_tiq = cruzar(datos, tol)
    tot = totales(datos)
    redactar_observaciones(filas, tot['zetti']['total']['total'] - tot['portal']['total']['total'], tol)
    sociedad = os.path.basename(os.path.dirname(os.path.dirname(mes)))
    sociedad = re.sub(r'^\d+\s*-\s*|\s*-\s*\d{11}$', '', sociedad)
    titulo = (f'CRUCE DE VENTAS - Mis Comprobantes / Portal IVA / Zetti - {sociedad}'
              + (f' ({razon})' if razon else '') + f' - {MESES[int(periodo[4:]) - 1]} {periodo[:4]}')
    avisos = [f"Puntos de venta de tiques (cruzados por dia): {', '.join(str(p) for p in pv_tiq)}"]
    for pv, n in sorted(redatados.items(), key=lambda x: str(x[0])):
        avisos.append(f"PV {pv}: ARCA informa los cierres Z con su numero (esquema 2025) y {n} tienen otra fecha que en "
                      f"Zetti: se cruzan por numero de Z, con la fecha de Zetti.")
    tiq_mc = tot['mc']['grupos']['Tiques']
    tiq_portal = tot['portal']['grupos']['Tiques']
    if not tiq_mc['n'] and tiq_portal['n']:
        avisos.append(f"Mis Comprobantes no trae los tiques del controlador fiscal (hasta 2025 no se informaban uno por uno): "
                      f"la diferencia Mis Comprobantes - Portal IVA son los cierres Z ({tiq_portal['n']} renglones, "
                      f"$ {pesos(tiq_portal['total'])}).")
    for mal, bien, nro, f in depurar_zetti.corregidos:
        avisos.append(f"Zetti tiene el cierre Z {nro} del {f.strftime('%d/%m') if hasattr(f, 'strftime') else f} cargado con PV {mal}: "
                      f"por la numeracion es del PV {bien} y se cruza como {bien}-{nro}. Corregir el punto de venta en Zetti.")
    if fuera:
        por_pv = defaultdict(lambda: defaultdict(int))
        for x in fuera:
            por_pv[x['punto']][x['tc_zetti']] += 1
        for pv, tipos in sorted(por_pv.items(), key=lambda x: str(x[0])):
            avisos.append(f"Eliminados de Zetti PV {pv}: " + ', '.join(f'{n} {t}' for t, n in sorted(tipos.items()))
                          + (' (excepcion)' if pv in exc else ' (PV que no figura en el Portal IVA)'))

    if papel:
        if HOJA_CRUCE not in wb.sheetnames and 'cruce' in wb.sheetnames:
            wb['cruce'].title = HOJA_CRUCE  # papeles armados antes de que la solapa se llamara "cruce ventas"
        if HOJA_CRUCE in wb.sheetnames:
            pos_h = wb.sheetnames.index(HOJA_CRUCE)
            del wb[HOJA_CRUCE]
            ws = wb.create_sheet(HOJA_CRUCE, pos_h)
        else:
            ws = wb.create_sheet(HOJA_CRUCE)
        salida = papel
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = 'Cruce'
        salida = opts.get('salida') or os.path.join(mes, f'{periodo} - CRUCE VENTAS MC PORTAL ZETTI.xlsx')
    r = bloque_resumen(ws, titulo, archivos, tot, filas, fuera, avisos)
    r = bloque_diferencias(ws, r, filas)
    estado_dia = {(rg['fecha'], rg['punto']): rg['estado'] for rg in filas if rg['modo'] == 'tiques'}
    per_ant, zetti_ant, origen_ant = leer_zetti_mes(mes, -1)
    per_sig, zetti_sig, origen_sig = leer_zetti_mes(mes, +1)
    ult_ant = {pv: (zs[-1][1], zs[-1][2]) for pv, zs in z_con_ti(zetti_ant).items()}
    correl = correlatividad(datos, ult_ant, estado_dia)
    otros = otros_periodos(datos, filas, {'anterior': (per_ant, zetti_ant), 'siguiente': (per_sig, zetti_sig)})
    r = bloque_correlatividad(ws, r, correl)
    r = bloque_otros_periodos(ws, r, otros)
    bloque_eliminados(ws, r, fuera)
    for j, w in enumerate([34, 44, 34, 26, 26, 16, 16], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for j in range(8, 21):
        ws.column_dimensions[get_column_letter(j)].width = 14
    wb.save(salida)

    res = {
        'salida': salida,
        'totales': {f: r2(tot[f]['total']['total']) for f in FUENTES},
        'eliminados_zetti': {'cantidad': len(fuera), 'total': r2(sum(x['total'] for x in fuera))},
        'zetti_menos_portal': r2(tot['zetti']['total']['total'] - tot['portal']['total']['total']),
        'mc_menos_portal': r2(tot['mc']['total']['total'] - tot['portal']['total']['total']),
        'zetti_otros_meses': {'anterior': origen_ant, 'siguiente': origen_sig},
        'otros_periodos': {per: len(v) for per, v in otros.items()},
        'correlatividad': [{'comprobante': c['comp'], 'pv': c['pv'], 'obs': c['obs'], 'info': c.get('info', [])}
                           for c in correl],
        'diferencias': [{'fecha': rg['fecha'].isoformat(), 'comprobante': concepto(rg), 'estado': rg['estado'],
                         'zetti_menos_portal': rg['dif_zetti_portal']} for rg in filas if rg['estado'] != 'OK'],
    }
    print(json.dumps(res, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
