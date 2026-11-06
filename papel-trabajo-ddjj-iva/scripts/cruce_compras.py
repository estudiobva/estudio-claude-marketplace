"""Etapa 2 (DDJJ de IVA), paso 3: cruce de COMPRAS Zetti / Portal IVA.

    python cruce_compras.py --papel=<papel de trabajo.xlsx> [--tolerancia=0.5]

Lee las solapas "Compras Zetti" (reporte 5.6.4, paso 1) y "Compras portal iva" (Libro IVA
Compras, paso 2) del papel y escribe o reemplaza la solapa "cruce compras" a continuacion
de "Compras portal iva". Mismo esquema que el cruce de ventas (cruce_3_fuentes.py), con dos
diferencias pedidas por la usuaria: de Zetti NO se elimina nada y no se usa Mis Comprobantes.

Como se cruza:
  1. Por CUIT del proveedor + punto de venta + numero (el tipo no se exige: Zetti carga
     FC C lo que en ARCA es "Tique Factura C").
  2. Lo que sobra, por CUIT + misma clase (factura / NC) + mismo total: "Numero distinto"
     (Zetti a veces guarda el numero mal, ej. NC de Monroe como "00NC-A-3116-").
  3. Lo que queda: "Solo en Portal IVA" / "Solo en Zetti", con observacion.

Columnas comparadas (Zetti / Portal):
  Exento/No grav.  Exen      / No Gravado + Exento
  Gravado          Grav      / Total Neto Gravado
  IVA              IVA       / suma de los IVA por alicuota (no "Total IVA")
  Perc. IIBB       P/R IIBB  / Percepciones de Ingresos Brutos
  Perc. IVA        P/R IVA   / Percepciones o Pagos a Cuenta de IVA
  Otros            P/R IG    / Per. Otros Imp. Nac. + Imp. Municipales + Imp. Internos + Otros Tributos
  Total            Total     / Importe Total

Solapa: 1 totales de cada fuente, 2 diferencias de total, 3 que explica las diferencias (la
suma de los estados da el punto 2), 4 detalle comprobante por comprobante, 5 percepciones con
el mismo total pero en otra columna o sin cerrar (agrupado por proveedor) y 6 diferencias por
proveedor. Imprime un JSON con el resumen.
"""
import json
import os
import re
import sys
import unicodedata
from collections import OrderedDict, defaultdict
from datetime import date, datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

COMP = ('exen', 'grav', 'iva', 'piibb', 'piva', 'otros', 'total')
TIT_COMP = ('Exento/No grav.', 'Gravado', 'IVA', 'Perc. IIBB', 'Perc. IVA', 'Otros', 'Total')
PERC = ('piibb', 'piva', 'otros')
NOMBRE = {'portal': 'Portal IVA', 'zetti': 'Zetti'}
FUENTES = ('portal', 'zetti')

TIPOS = {
    1: 'Factura A', 2: 'Nota de Debito A', 3: 'Nota de Credito A', 4: 'Recibo A',
    6: 'Factura B', 7: 'Nota de Debito B', 8: 'Nota de Credito B', 9: 'Recibo B',
    11: 'Factura C', 12: 'Nota de Debito C', 13: 'Nota de Credito C', 15: 'Recibo C',
    51: 'Factura M', 52: 'Nota de Debito M', 53: 'Nota de Credito M',
    60: 'Cta. de Vta. y Liq. A', 61: 'Cta. de Vta. y Liq. B', 63: 'Liquidacion A', 64: 'Liquidacion B',
    81: 'Tique Factura A', 82: 'Tique Factura B', 83: 'Tique', 111: 'Tique Factura C',
    112: 'Tique NC A', 113: 'Tique NC B', 114: 'Tique NC C',
    201: 'FCE MiPyME A', 202: 'ND MiPyME A', 203: 'NC MiPyME A',
    206: 'FCE MiPyME B', 207: 'ND MiPyME B', 208: 'NC MiPyME B',
    211: 'FCE MiPyME C', 212: 'ND MiPyME C', 213: 'NC MiPyME C',
}
NOTAS_CREDITO = {3, 8, 13, 21, 53, 90, 110, 112, 113, 114, 119, 203, 208, 213}
LIQUIDACIONES = {60, 61, 63, 64}
# Comprobantes C: no tienen IVA. El Portal IVA los expone solo en Importe Total (defecto del sistema, no es
# un error) y Zetti los carga en Exento/No gravado: se toman como Exento/No grav. y quedan OK (pedido 2026-11-02).
TIPOS_C = {11, 12, 13, 15, 111, 114, 117, 211, 212, 213}
DROGUERIAS_NCR = {'30516968431': 'Suizo Argentina', '30517059095': 'Monroe Americana'}

F_TIT = Font(name='Arial', size=12, bold=True)
F_B = Font(name='Arial', size=10, bold=True)
F_N = Font(name='Arial', size=10)
F_H = Font(name='Arial', size=10, bold=True, color='FFFFFF')
FILL_H = PatternFill('solid', fgColor='1F4E79')
FILL = {'portal': PatternFill('solid', fgColor='E2EFDA'), 'zetti': PatternFill('solid', fgColor='FFF2CC'),
        'dif': PatternFill('solid', fgColor='FCE4D6')}
FILL_ROJO = PatternFill('solid', fgColor='F8CBAD')
FMT = '#,##0.00;-#,##0.00;-'
FMT_F = 'DD/MM/YYYY'

HOJA_ZETTI = 'Compras Zetti'
HOJA_PORTAL = 'Compras portal iva'
HOJA_CRUCE = 'cruce compras'
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
         'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']


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


def solo_digitos(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return str(int(v))
    return re.sub(r'\D', '', str(v or ''))


def pesos(x):
    return '$' + f'{x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def ddmm(d):
    return d.strftime('%d/%m')


def vacio():
    return {k: 0.0 for k in COMP}


def sumar(acc, fila, signo=1):
    for k in COMP:
        acc[k] += signo * num(fila[k])


def fallar(msg, **extra):
    sys.exit(json.dumps({'error': msg, **extra}, ensure_ascii=False))


# ---------------------------------------------------------------------------
# Lectura
# ---------------------------------------------------------------------------

def leer_portal(ws):
    filas = list(ws.iter_rows(values_only=True))
    tit = [norm(v) for v in filas[0]]
    mapa = {'fecha de emision': 'fecha', 'tipo de comprobante': 'tipo', 'punto de venta': 'punto',
            'numero de comprobante': 'numero', 'nro. doc. vendedor': 'cuit',
            'denominacion vendedor': 'proveedor', 'importe total': 'total',
            'importe no gravado': 'no_grav', 'importe exento': 'exento', 'total neto gravado': 'grav',
            'importe de percepciones de ingresos brutos': 'piibb',
            'importe de percepciones o pagos a cuenta de iva': 'piva',
            'importe de per. o pagos a cta. de otros imp. nac.': 'otros_nac',
            'importe de impuestos municipales': 'municipales',
            'importe de impuestos internos': 'internos', 'importe otros tributos': 'otros_trib'}
    col = {mapa[h]: i for i, h in enumerate(tit) if h in mapa}
    faltan = set(mapa.values()) - set(col)
    if faltan:
        fallar(f'Solapa "{ws.title}": faltan columnas {sorted(faltan)}')
    cols_iva = [i for i, h in enumerate(tit) if re.match(r'^importe iva \d', h)]
    out = []
    for v in filas[1:]:
        g = lambda k: v[col[k]]  # noqa: E731
        fecha = como_fecha(g('fecha'))
        tipo = to_int(g('tipo'))
        if fecha is None or not isinstance(tipo, int):
            continue
        out.append({
            'fuente': 'portal', 'fecha': fecha, 'tipo': tipo, 'punto': to_int(g('punto')),
            'numero': to_int(g('numero')), 'nro_txt': None, 'cuit': solo_digitos(g('cuit')),
            'proveedor': str(g('proveedor') or '').strip(),
            'clase': 'NC' if tipo in NOTAS_CREDITO else 'LQ' if tipo in LIQUIDACIONES else 'FC',
            'exen': num(g('no_grav')) + num(g('exento')), 'grav': num(g('grav')),
            'iva': sum(num(v[i]) for i in cols_iva), 'piibb': num(g('piibb')), 'piva': num(g('piva')),
            'otros': num(g('otros_nac')) + num(g('municipales')) + num(g('internos')) + num(g('otros_trib')),
            'total': num(g('total')),
            'otros_nac': num(g('otros_nac')),
        })
        f = out[-1]
        if tipo in TIPOS_C and abs(f['total']) > 0.005 and all(abs(f[k]) <= 0.005 for k in COMP if k != 'total'):
            f['exen'] = f['total']
            f['c_solo_total'] = True
    return out


def leer_zetti(ws):
    filas = list(ws.iter_rows(values_only=True))
    enc = next((i for i, v in enumerate(filas) if norm(v[0]) == 'fecha' and norm(v[1]) == 'tc'), None)
    if enc is None:
        fallar(f'Solapa "{ws.title}": no encuentro la fila de titulos (Fecha, TC, ...)')
    tit = [norm(v) for v in filas[enc]]
    claves = {'fecha': 'fecha', 'tc': 'tc', 'm': 'm', 'nro. comp.': 'nro', 'proveedor': 'proveedor',
              'cuit': 'cuit', 'exen': 'exen', 'grav': 'grav', 'iva': 'iva', 'p/r iibb': 'piibb',
              'p/r iva': 'piva', 'p/r ig': 'otros', 'total': 'total'}
    faltan = [k for k in claves if k not in tit]
    if faltan:
        fallar(f'Solapa "{ws.title}": faltan columnas {faltan}')
    c = {claves[k]: tit.index(k) for k in claves}
    razon = str(filas[0][0] or '').split('  -  ')[0].strip() if enc else ''
    out = []
    for v in filas[enc + 1:]:
        fecha = como_fecha(v[c['fecha']])
        if fecha is None:
            continue
        pv, nro = v[c['nro']], v[c['nro'] + 1]
        nro_txt = None
        if not isinstance(pv, (int, float)) or not isinstance(nro, (int, float)):
            nro_txt = ' '.join(str(x) for x in (pv, nro) if x not in (None, ''))
            pv = nro = None
        tc = str(v[c['tc']] or '').strip().upper()
        m = str(v[c['m']] or '').strip().upper()
        out.append({
            'fuente': 'zetti', 'fecha': fecha, 'tc': tc, 'm': m, 'punto': to_int(pv), 'numero': to_int(nro),
            'nro_txt': nro_txt, 'cuit': solo_digitos(v[c['cuit']]), 'proveedor': str(v[c['proveedor']] or '').strip(),
            'clase': 'NC' if tc == 'NC' else 'FC',
            **{k: num(v[c[k]]) for k in COMP},
        })
    return out, razon


# ---------------------------------------------------------------------------
# Cruce
# ---------------------------------------------------------------------------

def descuadre(f):
    """Lo que le falta (o sobra) a las columnas para llegar al Total del comprobante."""
    return r2(f['total'] - sum(f[k] for k in COMP if k != 'total'))


def cruzar(portal, zetti, tol):
    pares, solo_p, solo_z = [], [], []
    idx = defaultdict(list)
    for p in portal:
        idx[(p['cuit'], p['punto'], p['numero'])].append(p)
    usados = set()
    resto_z = []
    for z in zetti:
        cand = [p for p in idx.get((z['cuit'], z['punto'], z['numero']), []) if id(p) not in usados] \
            if z['numero'] is not None else []
        if cand:
            p = min(cand, key=lambda x: abs(x['total'] - z['total']))
            usados.add(id(p))
            pares.append((p, z, 'numero'))
        else:
            resto_z.append(z)
    resto_p = [p for p in portal if id(p) not in usados]
    # Segunda vuelta: mismo proveedor, misma clase y mismo total (preferir la misma fecha).
    for z in resto_z:
        cand = [p for p in resto_p if id(p) not in usados and p['cuit'] == z['cuit']
                and (p['clase'] == z['clase'] or p['clase'] == 'LQ') and abs(p['total'] - z['total']) <= tol]
        if cand:
            p = min(cand, key=lambda x: abs((x['fecha'] - z['fecha']).days))
            usados.add(id(p))
            pares.append((p, z, 'total'))
        else:
            solo_z.append(z)
    solo_p = [p for p in resto_p if id(p) not in usados]
    return pares, solo_p, solo_z


def comprobante(f):
    if f['fuente'] == 'portal':
        t = TIPOS.get(f['tipo'], f"Tipo {f['tipo']}")
        return f"{t} {f['punto']:05d}-{f['numero']:08d}" if isinstance(f['punto'], int) and isinstance(f['numero'], int) \
            else f"{t} {f['punto']}-{f['numero']}"
    t = f"{f['tc']} {f['m']}".strip()
    if f['nro_txt']:
        return f"{t} {f['nro_txt']}"
    return f"{t} {f['punto']:05d}-{f['numero']:08d}"


def nombres_cols(ks, fuente):
    z = {'exen': 'Exen', 'grav': 'Grav', 'iva': 'IVA', 'piibb': 'P/R IIBB', 'piva': 'P/R IVA', 'otros': 'P/R IG'}
    p = {'exen': 'Exento/No grav.', 'grav': 'Neto gravado', 'iva': 'IVA', 'piibb': 'Perc. IIBB',
         'piva': 'Perc. IVA', 'otros': 'Per. Otros Imp. Nac./otros tributos'}
    m = z if fuente == 'zetti' else p
    return ' y '.join(m[k] for k in ks)


def clasificar(pares, solo_p, solo_z, tol):
    """Un renglon por comprobante (o par). estado/obs para los que no coinciden."""
    renglones = []
    for p, z, como in pares:
        rg = {'fecha': p['fecha'], 'p': p, 'z': z, 'dif': {k: r2(z[k] - p[k]) for k in COMP}}
        d = rg['dif']
        if abs(d['total']) > tol:
            rg['estado'] = 'Importe distinto'
            cols = [TIT_COMP[i] for i, k in enumerate(COMP[:-1]) if abs(d[k]) > tol]
            rg['obs'] = (f"Mismo comprobante con distinto total: Zetti {pesos(z['total'])}, Portal IVA {pesos(p['total'])} "
                         f"(diferencia {pesos(d['total'])}" + (f"; columnas: {', '.join(cols)}" if cols else '') + ').')
        elif como == 'total':
            rg['estado'] = 'Numero distinto'
            if z['nro_txt']:
                rg['obs'] = (f"Zetti tiene el numero mal cargado (\"{z['nro_txt']}\"); en el Portal IVA es "
                             f"{comprobante(p)}. Mismo proveedor y mismo total: es el mismo comprobante.")
            else:
                rg['obs'] = (f"Zetti lo tiene como {comprobante(z)} del {ddmm(z['fecha'])}; en el Portal IVA es "
                             f"{comprobante(p)} del {ddmm(p['fecha'])}. Mismo proveedor y mismo total: revisar el numero cargado en Zetti.")
        elif any(abs(d[k]) > tol for k in ('exen', 'grav', 'iva')):
            if all(abs(p[k]) <= tol for k in COMP if k != 'total'):
                rg['estado'] = 'Portal IVA sin desglose'
                rg['obs'] = (f"El Portal IVA trae solo el Importe Total ({pesos(p['total'])}) sin abrirlo en columnas; "
                             f"Zetti lo tiene en {nombres_cols([k for k in COMP[:-1] if abs(z[k]) > tol], 'zetti')}. "
                             'Revisar como queda en el Libro IVA Compras.')
            else:
                rg['estado'] = 'Distinta apertura'
                cols = [TIT_COMP[i] for i, k in enumerate(COMP[:-1]) if abs(d[k]) > tol]
                rg['obs'] = f"Mismo total, distinta apertura entre {', '.join(cols)}."
        elif any(abs(d[k]) > tol for k in PERC):
            if abs(descuadre(z)) > tol:
                rg['estado'] = 'Percepcion de Zetti no cierra'
            else:
                rg['estado'] = 'Percepcion en otra columna'
            rg['agrupado'] = True
        else:
            rg['estado'] = 'OK'
        renglones.append(rg)

    # Solo en Portal: pares que se anulan dentro del Portal (FC y NC/liquidacion del mismo proveedor e importe).
    anulados = set()
    for a in solo_p:
        if id(a) in anulados:
            continue
        for b in solo_p:
            if b is not a and id(b) not in anulados and b['cuit'] == a['cuit'] and abs(a['total'] + b['total']) <= 0.005 \
                    and abs(a['total']) > 0.005:
                anulados.update((id(a), id(b)))
                a['_par'], b['_par'] = b, a
                break
    for p in solo_p:
        rg = {'fecha': p['fecha'], 'p': p, 'z': None, 'dif': {k: r2(-p[k]) for k in COMP}, 'estado': 'Solo en Portal IVA'}
        if '_par' in p:
            rg['obs'] = (f"No esta en Zetti, pero se anula con {comprobante(p['_par'])} del mismo proveedor en el Portal IVA "
                         '(mismo importe con signo contrario): no afecta el total.')
        elif abs(p['total']) <= 0.005:
            rg['obs'] = 'Comprobante en $0: no esta en Zetti y no afecta importes.'
        elif p['tipo'] in LIQUIDACIONES:
            rg['obs'] = (f"{TIPOS.get(p['tipo'])} ({pesos(p['total'])}) que no esta cargada en Zetti: suele ser la "
                         'liquidacion de comisiones/gastos del banco o la tarjeta. Revisar si va al Libro IVA Compras.')
        elif abs(p['total']) < 1:
            rg['obs'] = f"Comprobante por centavos ({pesos(p['total'])}): ajuste o redondeo del proveedor, no esta en Zetti."
        elif p['clase'] == 'NC' and p['cuit'] in DROGUERIAS_NCR:
            rg['obs'] = (f"NC de {DROGUERIAS_NCR[p['cuit']]} ({pesos(p['total'])}) que no esta en Zetti. Revisar si es una "
                         'Nota de Credito de Recupero (NCR) que debia eliminarse del Libro IVA Compras.')
        else:
            rg['obs'] = f"No esta cargado en Zetti ({pesos(p['total'])}). Revisar si corresponde cargarlo."
        renglones.append(rg)
    cruzados = {(z['cuit'], z['punto'], z['numero']) for _, z, _ in pares if z['numero'] is not None}
    for z in solo_z:
        rg = {'fecha': z['fecha'], 'p': None, 'z': z, 'dif': {k: r2(z[k]) for k in COMP}, 'estado': 'Solo en Zetti'}
        if z['numero'] is not None and (z['cuit'], z['punto'], z['numero']) in cruzados:
            rg['estado'] = 'Duplicado en Zetti'
            rg['obs'] = (f"Zetti tiene este comprobante dos veces ({pesos(z['total'])} cada uno); en el Portal IVA esta una "
                         'sola vez. Hay que borrar el duplicado en Zetti.')
        elif z['m'] == 'X':
            rg['obs'] = (f"Comprobante X (no fiscal) de {pesos(z['total'])}: no se informa a ARCA, por eso no esta en el "
                         'Portal IVA. No computa para IVA.')
        else:
            rg['obs'] = (f"No esta en el Portal IVA ({pesos(z['total'])}). Revisar si el proveedor lo informo a ARCA o si "
                         'falta agregarlo al Libro IVA Compras.')
        renglones.append(rg)
    renglones.sort(key=lambda rg: (rg['estado'] == 'OK', rg['fecha'], rg_proveedor(rg)))
    return renglones


def rg_proveedor(rg):
    return (rg['p'] or rg['z'])['proveedor']


def rg_cuit(rg):
    return (rg['p'] or rg['z'])['cuit']


def obs_percepciones(filas):
    """Observacion de un proveedor con percepciones en otra columna / sin cerrar (con importes)."""
    z, p = vacio(), vacio()
    desc = 0.0
    for rg in filas:
        sumar(z, rg['z'])
        sumar(p, rg['p'])
        desc += descuadre(rg['z'])
    d = {k: z[k] - p[k] for k in PERC}
    partes = []
    if abs(desc) > 0.5:
        # lo que sobra en Zetti para llegar al total esta en P/R IIBB
        d['piibb'] += desc
        partes.append(f"En Zetti las columnas no suman el Total: el P/R IIBB tiene {pesos(-desc)} de mas "
                      f"(Zetti {pesos(z['piibb'])}, Portal IVA {pesos(p['piibb'])}). El total coincide con el Portal.")
    nz = {'piibb': 'P/R IIBB', 'piva': 'P/R IVA', 'otros': 'P/R IG'}
    np_ = {'piibb': 'Perc. IIBB', 'piva': 'Perc. IVA', 'otros': 'Per. Otros Imp. Nac'}
    mas_z = [k for k in PERC if d[k] > 0.5]
    mas_p = [k for k in PERC if d[k] < -0.5]
    if mas_z and mas_p:
        txt = ('Mismo total, distinta columna: Zetti tiene ' + ' y '.join(f'{pesos(d[k])} en {nz[k]}' for k in mas_z)
               + '; el Portal IVA tiene ' + ' y '.join(f'{pesos(-d[k])} en {np_[k]}' for k in mas_p) + '.')
        if 'piva' in mas_z:
            txt += (' Ojo: en la DDJJ de IVA solo computan las Percepciones de IVA del Libro Compras; '
                    'revisar que las percepciones de IVA no queden en Otros Imp. Nac.')
        partes.append(txt)
    return ' '.join(partes) or 'Mismo total, distinta apertura de percepciones.'


# ---------------------------------------------------------------------------
# Solapa
# ---------------------------------------------------------------------------

def encabezado(ws, fila, textos, col0=1):
    for j, h in enumerate(textos, start=col0):
        c = ws.cell(fila, j, h)
        c.font, c.fill = F_H, FILL_H
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def importes(ws, r, c0, vals, font=F_N):
    for k, comp in enumerate(COMP):
        c = ws.cell(r, c0 + k, r2(vals[comp]))
        c.number_format, c.font = FMT, font


def grupo_de(f):
    if f['fuente'] == 'zetti':
        return 'Comprobantes X' if f['m'] == 'X' else 'Notas de credito' if f['clase'] == 'NC' else 'Facturas y ND'
    return {'NC': 'Notas de credito', 'LQ': 'Liquidaciones'}.get(f['clase'], 'Facturas y ND')


def totales(datos):
    out = {}
    for f, filas in datos.items():
        g = OrderedDict()
        tot = {**vacio(), 'n': 0}
        for x in filas:
            k = grupo_de(x)
            g.setdefault(k, {**vacio(), 'n': 0})
            sumar(g[k], x); g[k]['n'] += 1
            sumar(tot, x); tot['n'] += 1
        orden = ['Facturas y ND', 'Notas de credito', 'Liquidaciones', 'Comprobantes X']
        out[f] = {'grupos': OrderedDict((k, g[k]) for k in orden if k in g), 'total': tot}
    return out


def bloque_resumen(ws, titulo, tot, renglones, avisos):
    ws.cell(1, 1, titulo).font = F_TIT
    ws.cell(2, 1, f'Zetti: solapa "{HOJA_ZETTI}" (reporte 5.6.4, sin depurar)').font = F_N
    ws.cell(3, 1, f'Portal IVA: solapa "{HOJA_PORTAL}" (Libro IVA Compras)').font = F_N
    r = 5
    ws.cell(r, 1, '1. Totales de cada fuente (NC en negativo)').font = F_B
    r += 1
    encabezado(ws, r, ['Fuente', 'Grupo', 'Cantidad', *TIT_COMP])
    r += 1
    for f in FUENTES:
        for gname, v in list(tot[f]['grupos'].items()) + [('TOTAL', tot[f]['total'])]:
            fnt = F_B if gname == 'TOTAL' else F_N
            ws.cell(r, 1, NOMBRE[f]).fill = FILL[f]
            ws.cell(r, 2, gname)
            ws.cell(r, 3, v['n'])
            for j in range(1, 4):
                ws.cell(r, j).font = fnt
            importes(ws, r, 4, v, fnt)
            r += 1
    r += 1
    ws.cell(r, 1, '2. Diferencias de total entre fuentes').font = F_B
    r += 1
    encabezado(ws, r, ['Comparacion', '', 'Cantidad', *TIT_COMP])
    r += 1
    ws.cell(r, 1, 'Zetti - Portal IVA').font = F_B
    ws.cell(r, 3, tot['zetti']['total']['n'] - tot['portal']['total']['n']).font = F_B
    importes(ws, r, 4, {k: tot['zetti']['total'][k] - tot['portal']['total'][k] for k in COMP}, F_B)
    r += 2

    ws.cell(r, 1, '3. Que explica las diferencias (Zetti - Portal IVA, suma de los renglones de los puntos 4 y 5)').font = F_B
    r += 1
    encabezado(ws, r, ['Estado', 'Detalle en', 'Renglones', *TIT_COMP])
    r += 1
    por_estado = OrderedDict()
    for rg in renglones:
        e = por_estado.setdefault(rg['estado'], {**vacio(), 'n': 0, 'agrupado': rg.get('agrupado', False)})
        e['n'] += 1
        for k in COMP:
            e[k] += rg['dif'][k]
    orden = ['Solo en Portal IVA', 'Solo en Zetti', 'Duplicado en Zetti', 'Importe distinto', 'Numero distinto', 'Portal IVA sin desglose',
             'Distinta apertura', 'Percepcion de Zetti no cierra', 'Percepcion en otra columna', 'OK']
    s = vacio()
    for estado in orden + [e for e in por_estado if e not in orden]:
        if estado not in por_estado:
            continue
        e = por_estado[estado]
        ws.cell(r, 1, 'OK (coinciden; solo redondeos)' if estado == 'OK' else estado)
        ws.cell(r, 2, '' if estado == 'OK' else 'punto 5' if e['agrupado'] else 'punto 4')
        ws.cell(r, 3, e['n'])
        for j in range(1, 4):
            ws.cell(r, j).font = F_N
        if estado != 'OK':
            ws.cell(r, 1).fill = FILL_ROJO
        importes(ws, r, 4, e)
        sumar(s, e)
        r += 1
    ws.cell(r, 1, 'TOTAL (= punto 2)').font = F_B
    importes(ws, r, 4, s, F_B)
    r += 2
    if avisos:
        ws.cell(r, 1, 'Avisos').font = F_B
        r += 1
        for a in avisos:
            ws.cell(r, 1, a).font = F_N
            r += 1
    return r + 1


def bloque_detalle(ws, r0, renglones):
    """Punto 4: comprobante por comprobante (todo lo que no es OK ni va agrupado al punto 5)."""
    ws.cell(r0, 1, '4. Detalle de las diferencias').font = F_B
    g, h = r0 + 1, r0 + 2
    fijas = ['Fecha', 'Comprobante (Portal IVA)', 'Comprobante (Zetti)', 'Proveedor', 'CUIT', 'Estado', 'Observacion']
    for j, t in enumerate(fijas, start=1):
        c = ws.cell(h, j, t)
        c.font, c.fill = F_H, FILL_H
        ws.cell(g, j).fill = FILL_H
    c0 = len(fijas) + 1
    col = {}
    for i, (f, nombre) in enumerate((('portal', 'PORTAL IVA'), ('zetti', 'ZETTI'), ('dif', 'DIFERENCIA ZETTI - PORTAL IVA'))):
        col[f] = c0 + i * len(COMP)
        ws.merge_cells(start_row=g, start_column=col[f], end_row=g, end_column=col[f] + len(COMP) - 1)
        c = ws.cell(g, col[f], nombre)
        c.font, c.fill, c.alignment = F_B, FILL[f], Alignment(horizontal='center')
        for k, t in enumerate(TIT_COMP):
            c = ws.cell(h, col[f] + k, t)
            c.font, c.fill, c.alignment = F_B, FILL[f], Alignment(horizontal='center', wrap_text=True)
    ult = col['dif'] + len(COMP) - 1
    r = d0 = h + 1
    for rg in renglones:
        if rg['estado'] == 'OK' or rg.get('agrupado'):
            continue
        ws.cell(r, 1, rg['fecha']).number_format = FMT_F
        ws.cell(r, 2, comprobante(rg['p']) if rg['p'] else '')
        ws.cell(r, 3, comprobante(rg['z']) if rg['z'] else '')
        ws.cell(r, 4, rg_proveedor(rg))
        ws.cell(r, 5, rg_cuit(rg))
        ws.cell(r, 6, rg['estado']).fill = FILL_ROJO
        ws.cell(r, 7, rg['obs']).alignment = Alignment(wrap_text=True, vertical='top')
        for j in range(1, 7):
            ws.cell(r, j).font = F_N
        ws.cell(r, 7).font = F_N
        if rg['p']:
            importes(ws, r, col['portal'], rg['p'])
        if rg['z']:
            importes(ws, r, col['zetti'], rg['z'])
        importes(ws, r, col['dif'], rg['dif'])
        r += 1
    if r == d0:
        ws.cell(r, 1, 'Sin diferencias comprobante por comprobante.').font = F_N
        r += 1
    ws.cell(r, 1, 'TOTAL').font = F_B
    for j in range(c0, ult + 1):
        L = get_column_letter(j)
        c = ws.cell(r, j, f'=SUBTOTAL(9,{L}{d0}:{L}{r - 1})' if r > d0 else 0)
        c.number_format, c.font = FMT, F_B
    if r > d0:
        ws.auto_filter.ref = f'A{h}:{get_column_letter(ult)}{r - 1}'
    return r + 3


def bloque_percepciones(ws, r0, renglones):
    """Punto 5: mismo comprobante y mismo total, percepciones en otra columna o sin cerrar, por proveedor."""
    ws.cell(r0, 1, '5. Percepciones con el mismo total pero en otra columna (agrupado por proveedor)').font = F_B
    por = OrderedDict()
    for rg in renglones:
        if rg.get('agrupado'):
            por.setdefault((rg_cuit(rg), rg['estado']), []).append(rg)
    if not por:
        ws.cell(r0 + 1, 1, 'No hay: las percepciones coinciden columna por columna.').font = F_N
        return r0 + 4
    g, h = r0 + 1, r0 + 2
    fijas = ['Proveedor', 'CUIT', 'Comprobantes', '', '', 'Estado', 'Observacion']
    for j, t in enumerate(fijas, start=1):
        c = ws.cell(h, j, t)
        c.font, c.fill = F_H, FILL_H
        ws.cell(g, j).fill = FILL_H
    c0 = len(fijas) + 1
    tits = ('Perc. IIBB', 'Perc. IVA', 'Otros', 'Total')
    ks = ('piibb', 'piva', 'otros', 'total')
    col = {}
    for i, (f, nombre) in enumerate((('portal', 'PORTAL IVA'), ('zetti', 'ZETTI'), ('dif', 'ZETTI - PORTAL IVA'))):
        col[f] = c0 + i * len(ks)
        ws.merge_cells(start_row=g, start_column=col[f], end_row=g, end_column=col[f] + len(ks) - 1)
        c = ws.cell(g, col[f], nombre)
        c.font, c.fill, c.alignment = F_B, FILL[f], Alignment(horizontal='center')
        for k, t in enumerate(tits):
            c = ws.cell(h, col[f] + k, t)
            c.font, c.fill, c.alignment = F_B, FILL[f], Alignment(horizontal='center', wrap_text=True)
    ult = col['dif'] + len(ks) - 1
    r = d0 = h + 1
    for (cuit, estado), filas in sorted(por.items(), key=lambda x: (rg_proveedor(x[1][0]), x[0][1])):
        p, z, d = vacio(), vacio(), vacio()
        for rg in filas:
            sumar(p, rg['p']); sumar(z, rg['z'])
            for k in COMP:
                d[k] += rg['dif'][k]
        ws.cell(r, 1, rg_proveedor(filas[0]))
        ws.cell(r, 2, cuit)
        ws.cell(r, 3, len(filas))
        ws.cell(r, 6, estado).fill = FILL_ROJO
        ws.cell(r, 7, obs_percepciones(filas)).alignment = Alignment(wrap_text=True, vertical='top')
        for j in (1, 2, 3, 6, 7):
            ws.cell(r, j).font = F_N
        for f, v in (('portal', p), ('zetti', z), ('dif', d)):
            for k, comp in enumerate(ks):
                c = ws.cell(r, col[f] + k, r2(v[comp]))
                c.number_format, c.font = FMT, F_N
        r += 1
    ws.cell(r, 1, 'TOTAL').font = F_B
    c = ws.cell(r, 3, f'=SUM(C{d0}:C{r - 1})')
    c.font = F_B
    for j in range(c0, ult + 1):
        L = get_column_letter(j)
        c = ws.cell(r, j, f'=SUM({L}{d0}:{L}{r - 1})')
        c.number_format, c.font = FMT, F_B
    return r + 3


def bloque_proveedores(ws, r0, renglones, tol):
    """Punto 6: proveedores con diferencia de total."""
    ws.cell(r0, 1, '6. Diferencias de total por proveedor').font = F_B
    h = r0 + 1
    encabezado(ws, h, ['Proveedor', 'CUIT', 'Comprobantes Portal IVA', 'Comprobantes Zetti', '', '', '',
                       'Total Portal IVA', 'Total Zetti', 'Zetti - Portal IVA'])
    por = OrderedDict()
    for rg in renglones:
        e = por.setdefault(rg_cuit(rg), {'prov': rg_proveedor(rg), 'np': 0, 'nz': 0, 'tp': 0.0, 'tz': 0.0})
        if rg['p']:
            e['np'] += 1; e['tp'] += rg['p']['total']
        if rg['z']:
            e['nz'] += 1; e['tz'] += rg['z']['total']
    r = d0 = h + 1
    for cuit, e in sorted(por.items(), key=lambda x: -abs(x[1]['tz'] - x[1]['tp'])):
        if abs(e['tz'] - e['tp']) <= tol:
            continue
        ws.cell(r, 1, e['prov'])
        ws.cell(r, 2, cuit)
        ws.cell(r, 3, e['np'])
        ws.cell(r, 4, e['nz'])
        for j, v in ((8, e['tp']), (9, e['tz']), (10, e['tz'] - e['tp'])):
            ws.cell(r, j, r2(v)).number_format = FMT
        for j in range(1, 11):
            ws.cell(r, j).font = F_N
        r += 1
    if r == d0:
        ws.cell(r, 1, 'Sin diferencias de total por proveedor.').font = F_N
        return r + 2
    ws.cell(r, 1, 'TOTAL').font = F_B
    for j in (3, 4, 8, 9, 10):
        L = get_column_letter(j)
        c = ws.cell(r, j, f'=SUM({L}{d0}:{L}{r - 1})')
        c.font = F_B
        if j >= 8:
            c.number_format = FMT
    return r + 2


def main():
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel = opts.get('papel')
    if not papel:
        sys.exit(__doc__)
    tol = float(opts.get('tolerancia') or 0.5)
    mes = os.path.dirname(os.path.abspath(papel))
    periodo = os.path.basename(mes)
    if not re.fullmatch(r'\d{6}', periodo):
        fallar(f'No reconozco el periodo AAAAMM en la ruta: {mes}')
    wb = load_workbook(papel)
    hojas = {}
    for f, nombre in (('zetti', HOJA_ZETTI), ('portal', HOJA_PORTAL)):
        hoja = next((x for x in wb.sheetnames if norm(x) == norm(nombre)), None)
        if hoja is None:
            fallar(f'Falta la solapa "{nombre}" en el papel', solapas=wb.sheetnames)
        hojas[f] = wb[hoja]
    portal = leer_portal(hojas['portal'])
    zetti, razon = leer_zetti(hojas['zetti'])
    if not portal or not zetti:
        fallar('Una de las solapas no trae comprobantes', portal=len(portal), zetti=len(zetti))

    pares, solo_p, solo_z = cruzar(portal, zetti, tol)
    renglones = clasificar(pares, solo_p, solo_z, tol)
    tot = totales({'portal': portal, 'zetti': zetti})

    # Nombre del papel ("AAAAMM- papeles de trabajo- <Sociedad>.xlsx"); si no, la carpeta de la sociedad
    # (sirve tambien con el papel en 01-Impuestos Mensuales/controles/AAAAMM o en una carpeta de anio).
    m_papel = re.match(r'^\d{6}\s*-\s*papeles de trabajo\s*-\s*(.+)\.xlsx$', os.path.basename(papel), re.I)
    if m_papel:
        sociedad = m_papel.group(1).strip()
    else:
        d = mes
        while d and os.path.dirname(d) != d and not re.search(r'-\s*\d{11}\s*$', os.path.basename(d)):
            d = os.path.dirname(d)
        sociedad = re.sub(r'^\d+\s*-\s*|\s*-\s*\d{11}$', '', os.path.basename(d))
    titulo = (f'CRUCE DE COMPRAS - Portal IVA / Zetti - {sociedad}' + (f' ({razon})' if razon else '')
              + f' - {MESES[int(periodo[4:]) - 1]} {periodo[:4]}')
    avisos = [f'Coinciden por CUIT + punto de venta + numero: {sum(1 for x in pares if x[2] == "numero")} comprobantes.']
    mal = [z for z in zetti if z['nro_txt']]
    if mal:
        avisos.append(f'Zetti tiene {len(mal)} comprobantes con el numero mal cargado (no PV-numero): se cruzaron por '
                      'proveedor y total (estado "Numero distinto").')
    no_cierran = [z for z in zetti if abs(descuadre(z)) > tol]
    if no_cierran:
        provs = sorted({z['proveedor'] for z in no_cierran})
        avisos.append(f'En Zetti hay {len(no_cierran)} comprobantes cuyas columnas no suman el Total ({", ".join(provs)}): '
                      'la suma de columnas de la solapa Compras Zetti no da el TOTAL.')
    solo_total_c = [p for p in portal if p.get('c_solo_total')]
    if solo_total_c:
        avisos.append(f'Comprobantes C: {len(solo_total_c)} vienen del Portal IVA solo con el Importe Total '
                      f'({pesos(sum(p["total"] for p in solo_total_c))}); como no tienen IVA se toman como Exento/No grav., '
                      'igual que en Zetti. No es diferencia.')
    if any(p['tipo'] not in TIPOS for p in portal):
        avisos.append('Tipos de comprobante del Portal no reconocidos: '
                      + ', '.join(sorted({str(p['tipo']) for p in portal if p['tipo'] not in TIPOS})))

    pos = wb.sheetnames.index(hojas['portal'].title) + 1
    if HOJA_CRUCE in wb.sheetnames:
        pos = wb.sheetnames.index(HOJA_CRUCE)
        del wb[HOJA_CRUCE]
    ws = wb.create_sheet(HOJA_CRUCE, pos)
    r = bloque_resumen(ws, titulo, tot, renglones, avisos)
    r = bloque_detalle(ws, r, renglones)
    r = bloque_percepciones(ws, r, renglones)
    bloque_proveedores(ws, r, renglones, tol)
    for j, w in enumerate([30, 30, 28, 34, 14, 26, 70], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for j in range(8, 30):
        ws.column_dimensions[get_column_letter(j)].width = 15
    wb.save(papel)

    estados = OrderedDict()
    for rg in renglones:
        e = estados.setdefault(rg['estado'], {'renglones': 0, 'zetti_menos_portal': 0.0})
        e['renglones'] += 1
        e['zetti_menos_portal'] += rg['dif']['total']
    for e in estados.values():
        e['zetti_menos_portal'] = r2(e['zetti_menos_portal'])
    print(json.dumps({
        'papel': papel, 'solapas': wb.sheetnames,
        'comprobantes': {'portal': len(portal), 'zetti': len(zetti)},
        'totales': {f: {k: r2(tot[f]['total'][k]) for k in COMP} for f in FUENTES},
        'zetti_menos_portal': {k: r2(tot['zetti']['total'][k] - tot['portal']['total'][k]) for k in COMP},
        'estados': estados,
        'avisos': avisos,
    }, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
