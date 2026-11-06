#!/usr/bin/env python3
"""
dfe_arba_a_excel.py — arma el Excel consolidado del Domicilio Fiscal
Electrónico de ARBA (lo que releva dfe-arba.js).

    python dfe_arba_a_excel.py <lote.json> <salida.xlsx>

Hojas:
  - Control: una fila por sociedad (Maestro de Sociedades + planilla de
    claves), con el estado del relevamiento y los totales de la bandeja.
  - Comunicaciones: todas las intimaciones y notificaciones de la bandeja, de
    la más nueva a la más vieja. En rojo las de los últimos N días.
  - Maestro vs Claves: diferencias entre el Maestro de Sociedades y la
    planilla Claves_Organismos (CUIT duplicados, filas que faltan, CIT que
    el Maestro da por existente y la planilla no tiene, etc.).

Si una hoja no tiene nada, se escribe igual una fila "Sin ...".
"""

import json
import sys
from collections import Counter
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FUENTE = 'Arial'
AZUL = '1F4E79'
ROJO = '9C0006'
GRIS = '808080'


def _encabezado(ws, columnas, anchos):
    for j, h in enumerate(columnas, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = Font(name=FUENTE, size=10, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=AZUL)
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        ws.column_dimensions[get_column_letter(j)].width = anchos[j - 1]
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = 'A2'


def _filas(ws, columnas, anchos, filas, fechas=(), resaltar=None, vacio='Sin datos'):
    _encabezado(ws, columnas, anchos)
    if not filas:
        c = ws.cell(row=2, column=1, value=vacio)
        c.font = Font(name=FUENTE, size=10, italic=True, color=GRIS)
        return
    for i, vals in enumerate(filas, start=2):
        rojo = bool(resaltar and resaltar(vals))
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(name=FUENTE, size=10, bold=rojo, color=ROJO if rojo else None)
            if j in fechas:
                c.number_format = 'DD/MM/YYYY'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(columnas))}{len(filas) + 1}'


def fecha(s):
    try:
        return datetime.strptime(str(s or '')[:10], '%Y-%m-%d').date()
    except ValueError:
        return None


def _base(nombre):
    return str(nombre or '').upper().split('(SUCURSAL')[0].strip()


def cuit_dudoso(ms):
    """Mismo CUIT en varias filas del Maestro con nombres distintos (las
    sucursales de una misma sociedad comparten CUIT y no se marcan)."""
    return len(ms) > 1 and len({_base(x['nombre']) for x in ms}) > 1


def formal(c):
    """Intimaciones y notificaciones; las 'Comunicaciones' son avisos que ARBA no notifica."""
    return not str(c.get('tipo') or '').lower().startswith('comunicaci')


def estado_com(c):
    if not formal(c):
        return 'Leída' if c.get('estaLeida') else 'Aviso sin leer'
    if not c.get('estaNotificada'):
        return 'Sin notificar aún'
    return 'Leída' if c.get('estaLeida') else 'Notificada, sin leer'


def main(entrada, salida):
    lote = json.load(open(entrada, encoding='utf-8'))
    hoy = fecha(lote['fecha'])
    dias = int(lote.get('dias') or 30)
    desde = hoy - timedelta(days=dias)
    claves = {c['cuit']: c for c in lote.get('claves', [])}
    maestro = lote.get('maestro', [])
    m_por_cuit = {}
    for m in maestro:
        m_por_cuit.setdefault(m['cuit'], []).append(m)
    socs = {s['cuit']: s for s in lote.get('sociedades', [])}

    wb = Workbook()

    # ── Control ──────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = 'Control'
    cuits = list(claves.keys()) + [c for c in m_por_cuit if c not in claves]
    filas = []
    for cuit in cuits:
        k = claves.get(cuit)
        ms = m_por_cuit.get(cuit, [])
        m = ms[0] if ms else {}
        s = socs.get(cuit)
        coms = (s or {}).get('comunicaciones') or []
        if s and s.get('ok'):
            est = 'Revisada'
        elif s:
            est = 'CIT rechazada' if s.get('code') == 'arba_login_rechazado' else 'Error'
        elif not k:
            est = 'No está en Claves_Organismos'
        elif not k.get('tieneCit'):
            est = 'Sin CIT de ARBA en la planilla'
        else:
            est = 'No relevada (filtro)'
        obs = []
        if s and not s.get('ok'):
            obs.append(s.get('error') or '')
        if cuit_dudoso(ms):
            obs.append('Mismo CUIT con otro nombre en el Maestro: ' + ' / '.join(x['nombre'] for x in ms))
        if k and not ms and maestro:
            obs.append('No está en el Maestro de Sociedades')
        recientes = [c for c in coms if (fecha(c.get('puestaADisposicion')) or hoy) >= desde]
        ultima = coms[0] if coms else None
        filas.append([
            k['pos'] if k else None, k['titular'] if k else None, m.get('nombre'), cuit,
            m.get('tipo'), m.get('iibb'), m.get('arba'), 'Sí' if k and k.get('tieneCit') else 'No', est,
            len(coms) if s and s.get('ok') else None,
            sum(1 for c in coms if not c.get('estaLeida')) if s and s.get('ok') else None,
            sum(1 for c in coms if formal(c) and not c.get('estaLeida')) if s and s.get('ok') else None,
            sum(1 for c in coms if formal(c) and not c.get('estaNotificada')) if s and s.get('ok') else None,
            len(recientes) if s and s.get('ok') else None,
            fecha(ultima['puestaADisposicion']) if ultima else None,
            (f"{ultima['tipo']}: {ultima['titulo']}" if ultima else ('Bandeja vacía' if s and s.get('ok') else None)),
            ' · '.join(x for x in obs if x),
        ])
    _filas(ws, ['#', 'Sociedad (Claves)', 'Nombre en Maestro', 'CUIT', 'Tipo', 'IIBB (Maestro)', 'CIT ARBA (Maestro)',
                'CIT en Claves', 'Estado', 'Total en bandeja', 'Sin leer (todas)',
                'Intimaciones / notificaciones sin leer', 'Intimaciones / notificaciones sin notificar aún',
                f'Recibidas últimos {dias} días', 'Última recibida', 'Última comunicación', 'Observaciones'],
           [5, 24, 34, 14, 10, 10, 11, 9, 22, 11, 10, 14, 14, 12, 12, 50, 60], filas, fechas=(15,),
           resaltar=lambda v: (v[13] or 0) > 0 or (v[12] or 0) > 0 or v[8] in ('Error', 'CIT rechazada'))

    # ── Comunicaciones ───────────────────────────────────────────────────
    ws = wb.create_sheet('Comunicaciones')
    filas = []
    for s in lote.get('sociedades', []):
        if not s.get('ok'):
            continue
        for c in s.get('comunicaciones') or []:
            filas.append([
                s.get('sociedad'), s.get('cuit'), c.get('tipo'),
                fecha(c.get('puestaADisposicion')), fecha(c.get('notificado')), fecha(c.get('leido')),
                estado_com(c), c.get('impuesto'), c.get('operativo'), c.get('nroNotificacion'), c.get('titulo'),
                c.get('destinatario'), c.get('expediente'), c.get('caracter'),
            ])
    filas.sort(key=lambda v: (v[3] or datetime(1900, 1, 1).date()), reverse=True)
    _filas(ws, ['Sociedad', 'CUIT', 'Tipo', 'Puesta a disposición', 'Notificado', 'Leído', 'Estado',
                'Impuesto / Régimen', 'Operativo', 'Nº notificación', 'Comunicación', 'Destinatario',
                'Expediente', 'Carácter'],
           [24, 14, 14, 13, 12, 12, 18, 18, 10, 12, 60, 40, 16, 12], filas, fechas=(4, 5, 6),
           resaltar=lambda v: v[3] is not None and v[3] >= desde,
           vacio='Sin comunicaciones en el DFE de las sociedades revisadas')

    # ── Maestro vs Claves ────────────────────────────────────────────────
    ws = wb.create_sheet('Maestro vs Claves')
    filas = []
    if lote.get('maestroError'):
        filas.append(['Maestro', None, None, 'No se pudo leer el Maestro de Sociedades', lote['maestroError']])
    for cuit, ms in m_por_cuit.items():
        if cuit_dudoso(ms):
            k = claves.get(cuit)
            filas.append(['Mismo CUIT, nombres distintos (Maestro)', cuit, ' / '.join(x['nombre'] for x in ms),
                          f"#{k['pos']} {k['titular']}" if k else None,
                          'Verificar si es una sucursal o un CUIT mal cargado en una de las filas'])
    for cuit, ms in m_por_cuit.items():
        k = claves.get(cuit)
        m = ms[0]
        if not k:
            filas.append(['En el Maestro, no en Claves', cuit, m['nombre'], None,
                          f"Tipo {m.get('tipo') or '-'}, IIBB {m.get('iibb') or '-'}: no se puede revisar el DFE sin CIT"])
        elif m.get('arba') and not k.get('tieneCit'):
            filas.append(['Maestro dice CIT, Claves no la tiene', cuit, m['nombre'], f"#{k['pos']} {k['titular']}",
                          f"Maestro ARBA = {m['arba']}; cargar la CIT en Claves_Organismos"])
        elif not m.get('arba') and k.get('tieneCit'):
            filas.append(['Claves tiene CIT, Maestro no', cuit, m['nombre'], f"#{k['pos']} {k['titular']}",
                          'Completar la columna ARBA del Maestro'])
    if maestro:
        for cuit, k in claves.items():
            if cuit not in m_por_cuit:
                filas.append(['En Claves, no en el Maestro', cuit, None, f"#{k['pos']} {k['titular']}",
                              'CIT de ARBA cargada' if k.get('tieneCit') else 'Sin CIT de ARBA'])
    _filas(ws, ['Diferencia', 'CUIT', 'Nombre en Maestro', 'Fila en Claves', 'Detalle'],
           [34, 14, 46, 30, 70], filas, vacio='Sin diferencias entre el Maestro y la planilla de claves')

    wb.save(salida)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
