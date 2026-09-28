#!/usr/bin/env python3
"""
planes_a_excel.py — arma los Excel de planes de facilidades de pago (Mis Facilidades).

    python planes_a_excel.py <sociedad.json> <salida.xlsx>                 (una sociedad)
    python planes_a_excel.py --consolidado <lote.json> <salida.xlsx>       (informe de todas)

Por sociedad (formato de la skill planes-facilidades-arca): hojas RESUMEN,
DETALLE, OBSERVACIONES y PRESENTACIONES.

Informe consolidado: INFORME (una fila por sociedad), CUOTAS IMPAGAS (una fila
por cuota vencida sin pagar, lo que pone en riesgo el plan), PLANES VIGENTES,
DETALLE CUOTAS y CONTROL (errores, omitidas y observaciones).

Las cuotas impagas van en rojo en todas las hojas: son lo que hay que mirar.
"""

import json
import re
import sys
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

FUENTE = 'Arial'
AZUL = '1F4E79'
ROJO = '9C0006'
ROJO_BG = 'FFC7CE'
AMARILLO_BG = 'FFEB9C'
VERDE_BG = 'C6EFCE'
GRIS_BG = 'E7E6E6'
FMT_PESOS = '#,##0.00;-#,##0.00;-'
FMT_FECHA = 'DD/MM/YYYY'


def a_fecha(s):
    m = re.search(r'(\d{2})/(\d{2})/(\d{4})', str(s or ''))
    if not m:
        return None
    return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))).date()


def hoja(wb, titulo, columnas, anchos=None, primera=False):
    ws = wb.active if primera else wb.create_sheet()
    ws.title = titulo
    for j, h in enumerate(columnas, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = Font(name=FUENTE, size=10, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=AZUL)
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = 'A2'
    for j, w in enumerate(anchos or [], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    return ws


def fila(ws, i, vals, pesos=(), fechas=(), fondo=None, rojo=False, negrita=False, wrap=()):
    for j, v in enumerate(vals, start=1):
        if j in fechas and isinstance(v, str):
            v = a_fecha(v) or (v or None)
        c = ws.cell(row=i, column=j, value=v)
        c.font = Font(name=FUENTE, size=10, bold=negrita or rojo, color=ROJO if rojo else None)
        if fondo:
            c.fill = PatternFill('solid', fgColor=fondo)
        if j in pesos:
            c.number_format = FMT_PESOS
        if j in fechas:
            c.number_format = FMT_FECHA
            c.alignment = Alignment(horizontal='center')
        if j in wrap:
            c.alignment = Alignment(wrap_text=True, vertical='top')


def texto_impagas(p):
    return '; '.join(
        f"cuota {i['cuota']} vto {i['vencimiento']} ({i['diasAtraso']} días"
        + (', 2° vto ' + i['segundoVenc'] + ' pendiente' if i.get('segundoVencPendiente') else '')
        + ')'
        for i in p.get('impagas', []))


# Codigo de entidad = primeros 3 digitos del CBU.
BANCOS = {'007': 'Galicia', '011': 'Nación', '014': 'Provincia', '015': 'ICBC', '017': 'BBVA',
          '027': 'Supervielle', '029': 'Ciudad', '034': 'Patagonia', '044': 'Hipotecario',
          '072': 'Santander', '150': 'HSBC / Galicia Más', '191': 'Credicoop', '285': 'Macro',
          '299': 'Comafi', '322': 'Industrial', '389': 'Columbia'}


def banco(cbu):
    return BANCOS.get(str(cbu or '')[:3], '')


def cbu_txt(p):
    d, v = p.get('cbuDeclarado') or '', p.get('cbuVigente') or ''
    t = d + (f' ({banco(d)})' if banco(d) else '')
    if v and v != d:
        t += f' · vigente hoy: {v}' + (f' ({banco(v)})' if banco(v) else '')
    return t or None


def conceptos_txt(p):
    return '\n'.join(p.get('conceptos') or []) or None


def vigentes(planes):
    return [p for p in planes if re.search('vigente', p.get('situacion') or '', re.I)]


# ── Por sociedad ─────────────────────────────────────────────────────────────

def escribir_sociedad(d, salida):
    wb = Workbook()
    planes = d.get('planes', [])

    cols = ['Nro plan', 'Fecha presentación', 'Tipo', 'Tipo de plan', 'Cant. cuotas', 'Consolidado',
            'Situación', 'Cuotas pagadas', 'Cuotas IMPAGAS', 'Detalle impagas', 'Capital pagado',
            'Intereses pagados', 'Total pagado', 'Saldo de capital', 'Próx. cuota N°',
            'Próx. vencimiento', 'Próx. importe', 'Pagadas fuera de término', 'Alerta',
            'CBU declarado', 'Obligaciones incluidas (concepto)']
    ws = hoja(wb, 'RESUMEN', cols, [11, 12, 40, 9, 8, 15, 12, 9, 9, 45, 15, 14, 15, 15, 9, 12, 14, 10, 20, 34, 55], primera=True)
    i = 2
    for p in planes:
        pr = p.get('proxima') or {}
        imp = p.get('cuotasImpagas') or 0
        fila(ws, i, [p.get('nroPlan'), p.get('presentacion'), p.get('tipo'), p.get('tipoPlan'),
                     p.get('cantCuotas'), p.get('consolidado'), p.get('situacion'), p.get('cuotasPagadas'),
                     imp, texto_impagas(p), p.get('capitalPagado'), p.get('interesesPagados'),
                     p.get('totalPagado'), p.get('saldoCapital'), pr.get('cuota'), pr.get('vencimiento'),
                     pr.get('importe'), p.get('fueraDeTermino'), p.get('alerta'),
                     cbu_txt(p), conceptos_txt(p)],
             pesos=(6, 11, 12, 13, 14, 17), fechas=(2, 16), wrap=(3, 10, 20, 21),
             fondo=ROJO_BG if imp else None, rojo=bool(imp))
        i += 1
    if not planes:
        fila(ws, 2, ['Sin planes vigentes'])
        i = 3
    else:
        # Totales con formulas
        fila(ws, i, ['TOTAL'], negrita=True)
        for col in (6, 11, 12, 13, 14):
            L = get_column_letter(col)
            c = ws.cell(row=i, column=col, value=f'=SUM({L}2:{L}{i - 1})')
            c.number_format = FMT_PESOS
            c.font = Font(name=FUENTE, size=10, bold=True)
        c = ws.cell(row=i, column=9, value=f'=SUM(I2:I{i - 1})')
        c.font = Font(name=FUENTE, size=10, bold=True)

    cols = ['Nro plan', 'Cuota N°', 'Capital', 'Interés financiero', 'Total 1° vto', 'Fecha venc.',
            '2° vto', 'Total 2° vto', 'Estado', 'Estado ARCA', 'Fecha de pago', 'Total pagado',
            'Int. resarcitorio pagado', 'Fuera de término', 'Días de atraso', 'Intentos de débito fallidos']
    ws = hoja(wb, 'DETALLE', cols, [11, 12, 15, 14, 15, 12, 12, 15, 11, 16, 12, 15, 12, 9, 9, 50])
    i = 2
    for p in planes:
        for c in p.get('cuotas', []):
            impaga = c.get('estado') == 'IMPAGA'
            fila(ws, i, [p.get('nroPlan'), c.get('cuota'), c.get('capital'), c.get('intFinanciero'),
                         c.get('total'), c.get('venc1'), c.get('venc2') or None, c.get('totalVenc2'),
                         c.get('estado'), c.get('estadoArca'), c.get('fechaPago') or None, c.get('pagadoTotal'),
                         c.get('intResarcitorioPagado'), 'S' if c.get('fueraDeTermino') else 'N',
                         c.get('diasAtraso'), '; '.join(c.get('intentosFallidos') or [])],
                 pesos=(3, 4, 5, 8, 12, 13), fechas=(6, 7, 11), wrap=(16,),
                 fondo=ROJO_BG if impaga else None, rojo=impaga)
            i += 1

    ws = hoja(wb, 'OBSERVACIONES', ['Nro plan', 'Observación'], [12, 110])
    i = 2
    for p in planes:
        if p.get('cuotasImpagas'):
            fila(ws, i, [p.get('nroPlan'), f"{p['cuotasImpagas']} CUOTA(S) IMPAGA(S) VENCIDA(S): {texto_impagas(p)}"],
                 fondo=ROJO_BG, rojo=True, wrap=(2,))
            i += 1
        for o in p.get('observaciones', []):
            fila(ws, i, [p.get('nroPlan'), o], wrap=(2,))
            i += 1
    if i == 2:
        fila(ws, 2, ['', 'Sin observaciones.'])

    escribir_obligaciones(wb, [(d.get('sociedad'), d.get('cuit'), p) for p in planes])

    cols = ['Presentación', 'Número', 'Cuotas', 'Tipo', 'Consolidado', 'Estado', 'Situación']
    ws = hoja(wb, 'PRESENTACIONES', cols, [12, 11, 8, 70, 15, 11, 15])
    for i, p in enumerate(d.get('presentaciones', []), start=2):
        fila(ws, i, [p.get('presentacion'), p.get('numero'), p.get('cuotas'), p.get('tipo'),
                     p.get('consolidado'), p.get('estado'), p.get('situacion')],
             pesos=(5,), fechas=(1,), wrap=(4,))

    wb.save(salida)
    return {'ok': True, 'archivo': salida, 'planes': len(planes)}


def escribir_obligaciones(wb, planes):
    """Una fila por obligacion incluida en cada plan: solo el concepto, sin importes."""
    cols = ['Sociedad', 'CUIT', 'Nro plan', 'CBU declarado', 'Banco', 'Tipo', 'Impuesto', 'Concepto',
            'Subconcepto', 'Período', 'Vencimiento original']
    ws = hoja(wb, 'OBLIGACIONES', cols, [30, 13, 11, 25, 14, 11, 32, 22, 24, 10, 12])
    i = 2
    for soc, cuit, p in planes:
        obs = p.get('obligaciones') or []
        imp = bool(p.get('cuotasImpagas'))
        if not obs:
            fila(ws, i, [soc, cuit, p.get('nroPlan'), p.get('cbuDeclarado'), banco(p.get('cbuDeclarado')),
                         None, 'Sin dato de obligaciones'], fondo=ROJO_BG if imp else None, rojo=imp)
            i += 1
        for o in obs:
            per = o.get('periodo') or '/'.join(x for x in (o.get('anio'), o.get('mes')) if x)
            fila(ws, i, [soc, cuit, p.get('nroPlan'), p.get('cbuDeclarado'), banco(p.get('cbuDeclarado')),
                         o.get('tipo'), o.get('impuesto'), o.get('concepto'), o.get('subconcepto'), per,
                         o.get('vencimiento')], fechas=(11,), fondo=ROJO_BG if imp else None, rojo=imp)
            i += 1
    ws.auto_filter.ref = f'A1:{get_column_letter(len(cols))}{max(i - 1, 1)}'


# ── Consolidado ──────────────────────────────────────────────────────────────

def escribir_consolidado(lote, salida):
    wb = Workbook()
    datos = []
    for s in lote.get('sociedades', []):
        d = None
        if s.get('ok') and s.get('datos'):
            try:
                with open(s['datos'], encoding='utf-8') as f:
                    d = json.load(f)
            except OSError:
                d = None
        datos.append((s, d))

    cols = ['#', 'Sociedad', 'CUIT', 'Presentaciones', 'Planes vigentes', 'Cuotas IMPAGAS',
            'Situación', 'Detalle de cuotas impagas', 'Saldo de capital (vigentes)',
            'Próximo vencimiento', 'Importe próx.', 'Observaciones']
    ws = hoja(wb, 'INFORME', cols, [5, 32, 13, 9, 9, 9, 22, 70, 16, 12, 14, 50], primera=True)
    def prioridad(sd):
        s, d = sd
        if not s.get('ok'):
            return 0
        return 1 if any(p.get('cuotasImpagas') for p in vigentes(d.get('planes', []))) else 2
    i = 2
    for s, d in sorted(datos, key=prioridad):
        if not s.get('ok'):
            fila(ws, i, [s.get('pos'), s.get('sociedad'), s.get('cuit'), None, None, None,
                         'NO RELEVADA', None, None, None, None, s.get('error')],
                 fondo=GRIS_BG, wrap=(12,))
            i += 1
            continue
        vig = vigentes(d.get('planes', []))
        imp = sum(p.get('cuotasImpagas') or 0 for p in vig)
        detalle = '\n'.join(f"Plan {p['nroPlan']} (CBU {cbu_txt(p)}): {texto_impagas(p)}" for p in vig if p.get('cuotasImpagas'))
        saldo = sum(p.get('saldoCapital') or 0 for p in vig) if vig else None
        prox = sorted([p['proxima'] for p in vig if p.get('proxima')],
                      key=lambda x: a_fecha(x.get('vencimiento')) or datetime.max.date())
        obs = '; '.join(o for p in vig for o in p.get('observaciones', [])
                        if not o.startswith('Cuota ') or 'sin registro' in o)
        fondo = ROJO_BG if imp else (VERDE_BG if vig else None)
        fila(ws, i, [s.get('pos'), s.get('sociedad'), s.get('cuit'), len(d.get('presentaciones', [])),
                     len(vig), imp, s.get('alerta'), detalle or None, saldo,
                     prox[0]['vencimiento'] if prox else None, prox[0]['importe'] if prox else None, obs or None],
             pesos=(9, 11), fechas=(10,), wrap=(8, 12), fondo=fondo, rojo=bool(imp))
        i += 1
    for o in lote.get('omitidos', []):
        if 'repetido' in (o.get('motivo') or ''):
            continue
        fila(ws, i, [o.get('pos'), o.get('titular'), None, None, None, None, 'NO RELEVADA',
                     None, None, None, None, o.get('motivo')], fondo=GRIS_BG, wrap=(12,))
        i += 1
    ws.auto_filter.ref = f'A1:{get_column_letter(len(cols))}{max(i - 1, 1)}'

    cols = ['Sociedad', 'CUIT', 'Nro plan', 'Tipo de plan', 'Cuotas impagas en el plan', 'Cuota N°',
            'Vencimiento', '2° vencimiento', 'Días de atraso', 'Importe cuota', 'Intentos de débito fallidos',
            'Situación', 'CBU declarado', 'Obligaciones incluidas (concepto)']
    ws = hoja(wb, 'CUOTAS IMPAGAS', cols, [30, 13, 11, 40, 10, 8, 12, 12, 9, 15, 50, 22, 34, 50])
    i = 2
    for s, d in datos:
        if not d:
            continue
        for p in vigentes(d.get('planes', [])):
            for c in p.get('impagas', []):
                fila(ws, i, [s.get('sociedad'), s.get('cuit'), p.get('nroPlan'), p.get('tipo'),
                             p.get('cuotasImpagas'), c.get('cuota'), c.get('vencimiento'),
                             c.get('segundoVenc') or None, c.get('diasAtraso'), c.get('importe'),
                             '; '.join(c.get('intentosFallidos') or []), p.get('alerta'),
                             cbu_txt(p), conceptos_txt(p)],
                     pesos=(10,), fechas=(7, 8), wrap=(4, 11, 13, 14), fondo=ROJO_BG, rojo=True)
                i += 1
    if i == 2:
        fila(ws, 2, ['No hay cuotas impagas vencidas en ningún plan vigente.'], fondo=VERDE_BG)

    cols = ['Sociedad', 'CUIT', 'Nro plan', 'Fecha presentación', 'Tipo', 'Cant. cuotas', 'Consolidado',
            'Cuotas pagadas', 'Cuotas IMPAGAS', 'Cuotas a vencer', 'Capital pagado', 'Intereses pagados',
            'Total pagado', 'Saldo de capital', 'Próx. cuota N°', 'Próx. vencimiento', 'Próx. importe',
            'Pagadas fuera de término', 'Situación', 'CBU declarado', 'Obligaciones incluidas (concepto)']
    ws = hoja(wb, 'PLANES VIGENTES', cols, [30, 13, 11, 12, 40, 8, 15, 9, 9, 9, 15, 14, 15, 15, 9, 12, 14, 10, 22, 34, 55])
    i = 2
    for s, d in datos:
        if not d:
            continue
        for p in vigentes(d.get('planes', [])):
            pr = p.get('proxima') or {}
            imp = p.get('cuotasImpagas') or 0
            fila(ws, i, [s.get('sociedad'), s.get('cuit'), p.get('nroPlan'), p.get('presentacion'), p.get('tipo'),
                         p.get('cantCuotas'), p.get('consolidado'), p.get('cuotasPagadas'), imp,
                         p.get('cuotasAVencer'), p.get('capitalPagado'), p.get('interesesPagados'),
                         p.get('totalPagado'), p.get('saldoCapital'), pr.get('cuota'), pr.get('vencimiento'),
                         pr.get('importe'), p.get('fueraDeTermino'), p.get('alerta'),
                         cbu_txt(p), conceptos_txt(p)],
                 pesos=(7, 11, 12, 13, 14, 17), fechas=(4, 16), wrap=(5, 20, 21),
                 fondo=ROJO_BG if imp else None, rojo=bool(imp))
            i += 1
    if i > 2:
        fila(ws, i, ['TOTAL'], negrita=True)
        for col in (7, 11, 12, 13, 14):
            L = get_column_letter(col)
            c = ws.cell(row=i, column=col, value=f'=SUM({L}2:{L}{i - 1})')
            c.number_format = FMT_PESOS
            c.font = Font(name=FUENTE, size=10, bold=True)
        c = ws.cell(row=i, column=9, value=f'=SUM(I2:I{i - 1})')
        c.font = Font(name=FUENTE, size=10, bold=True)

    escribir_obligaciones(wb, [(s.get('sociedad'), s.get('cuit'), p)
                               for s, d in datos if d for p in vigentes(d.get('planes', []))])

    cols = ['Sociedad', 'Nro plan', 'Cuota N°', 'Capital', 'Total 1° vto', 'Fecha venc.', 'Estado',
            'Fecha de pago', 'Total pagado', 'Fuera de término', 'Días de atraso', 'Intentos de débito fallidos']
    ws = hoja(wb, 'DETALLE CUOTAS', cols, [30, 11, 12, 15, 15, 12, 11, 12, 15, 9, 9, 50])
    i = 2
    for s, d in datos:
        if not d:
            continue
        for p in vigentes(d.get('planes', [])):
            for c in p.get('cuotas', []):
                impaga = c.get('estado') == 'IMPAGA'
                fila(ws, i, [s.get('sociedad'), p.get('nroPlan'), c.get('cuota'), c.get('capital'), c.get('total'),
                             c.get('venc1'), c.get('estado'), c.get('fechaPago') or None, c.get('pagadoTotal'),
                             'S' if c.get('fueraDeTermino') else 'N', c.get('diasAtraso'),
                             '; '.join(c.get('intentosFallidos') or [])],
                     pesos=(4, 5, 9), fechas=(6, 8), wrap=(12,),
                     fondo=ROJO_BG if impaga else None, rojo=impaga)
                i += 1
    ws.auto_filter.ref = f'A1:{get_column_letter(len(cols))}{max(i - 1, 1)}'

    ws = hoja(wb, 'CONTROL', ['Sociedad', 'CUIT', 'Tipo', 'Detalle'], [32, 13, 18, 100])
    i = 2
    for s, d in datos:
        if not s.get('ok'):
            fila(ws, i, [s.get('sociedad'), s.get('cuit'), 'ERROR ' + str(s.get('code') or ''), s.get('error')],
                 fondo=AMARILLO_BG, wrap=(4,))
            i += 1
            continue
        for p in vigentes(d.get('planes', [])) if d else []:
            for o in p.get('observaciones', []):
                fila(ws, i, [s.get('sociedad'), s.get('cuit'), f"Plan {p.get('nroPlan')}", o], wrap=(4,))
                i += 1
    for o in lote.get('omitidos', []):
        fila(ws, i, [o.get('titular'), None, 'OMITIDA', o.get('motivo')], fondo=GRIS_BG)
        i += 1
    fila(ws, i + 1, [f"Relevado el {lote.get('fecha')} desde Mis Facilidades (ARCA). "
                     "IMPAGA = cuota sin pago con 1° vencimiento anterior a la fecha del relevamiento."],
         negrita=True)

    wb.save(salida)
    return {'ok': True, 'archivo': salida, 'sociedades': len(datos)}


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == '--consolidado':
        with open(a[1], encoding='utf-8') as f:
            res = escribir_consolidado(json.load(f), a[2])
    else:
        with open(a[0], encoding='utf-8') as f:
            res = escribir_sociedad(json.load(f), a[1])
    print(json.dumps(res, ensure_ascii=False))
