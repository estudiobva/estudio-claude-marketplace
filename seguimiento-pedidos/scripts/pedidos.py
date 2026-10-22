#!/usr/bin/env python3
"""
Planilla de seguimiento de pedidos Estudio <-> Cliente.

Subcomandos (todos devuelven JSON por stdout, salvo `pendientes --formato=texto|slack`):

  crear      [--ruta=<xlsx>]                       Crea la planilla vacia (no pisa una existente).
  estado     --email=<mail> | --persona=<nombre> | --slack=<ID>
                                                   Quien corre, desde cuando leer, configuracion y
                                                   pedidos que ya estan cargados (para no duplicar).
  aplicar    <ops.json> [--solo-revisar]           Agrega pedidos, actualiza estados y suma reclamos.
  pendientes --persona=<nombre> [--formato=json|texto|slack] [--hoy=AAAA-MM-DD]
                                                   Lo abierto de esa persona, vencidos primero.

Ruta de la planilla: --ruta, o PEDIDOS_PLANILLA del .env, o
<BVA_UNIDAD_PATH>/0 - ESTUDIO - 00000000000/Seguimiento de pedidos Estudio-Cliente.xlsx

Formato de <ops.json>:
  {"persona": "Victoria", "desde": "2026-10-08", "hasta": "2026-10-22",
   "ops": [
     {"accion": "agregar", "hoja": "EC", "clave_origen": "gmail:<threadId>#1",
      "datos": {"fecha_pedido": "2026-10-09", "a_quien": "<persona del cliente>", "que": "...", ...}},
     {"accion": "actualizar", "id": "E-003", "datos": {"estado": "Completado", "fecha_recibido": "2026-10-20"}},
     {"accion": "reclamo", "id": "E-003", "fecha": "2026-10-22", "detalle": "Reclamado por mail"}
   ]}
  Hojas: EC = Estudio -> Cliente, CE = Cliente -> Estudio, TI = Tareas internas.
"""
import json
import os
import re
import shutil
import sys
import tempfile
from datetime import date, datetime, timedelta

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    print(json.dumps({'ok': False, 'error': 'Falta openpyxl: pip install openpyxl'}, ensure_ascii=False))
    sys.exit(1)

PLUGIN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAIZ_DEFECTO = r'I:\Unidades compartidas\BVA - Sociedades Farmaceuticas'
PLANILLA_REL = os.path.join('0 - ESTUDIO - 00000000000', 'Seguimiento de pedidos Estudio-Cliente.xlsx')
FILAS = 500  # filas preparadas con formulas y desplegables
FILA_ENC, FILA_1 = 2, 3

ESTADOS = ['Pendiente', 'En curso', 'Respondido parcial', 'Completado', 'Cancelado', 'A confirmar']
CERRADOS = {'Completado', 'Cancelado'}
CANALES = ['Mail', 'WhatsApp', 'Teléfono', 'SharePoint', 'Reunión', 'Slack']
TEMAS = ['EERR', 'Rectificativas', 'Calificaciones bancarias', 'Inspecciones', 'Liquidaciones',
         'Sueldos', 'Bancos', 'Gestión interna', 'Otros']

# (clave, encabezado, ancho, tipo). Tipos: fecha, texto, num, dias (formula), lista:<Lista>
HOJAS = {
    'EC': {
        'titulo': 'Estudio → Cliente', 'prefijo': 'E',
        'nota': 'Lo que el estudio le pide al cliente. Antes de pedir algo, filtrá por tema/sociedad: '
                'si ya está pedido, sumá un reclamo en vez de pedirlo de nuevo.',
        'cols': [
            ('id', 'ID', 7, 'texto'), ('fecha_pedido', 'Fecha pedido', 12, 'fecha'),
            ('pedido_por', 'Pedido por (estudio)', 14, 'lista:Equipo'),
            ('responsable', 'Responsable seguimiento', 14, 'lista:Equipo'),
            ('a_quien', 'A quién (cliente)', 18, 'lista:Cliente'), ('sociedad', 'Sociedad', 14, 'texto'),
            ('tema', 'Tema', 14, 'lista:Tema'), ('que', 'Qué se pidió', 46, 'texto'),
            ('canal', 'Canal', 10, 'lista:Canal'), ('fecha_maxima', 'Fecha máxima', 12, 'fecha'),
            ('estado', 'Estado', 14, 'lista:Estado'), ('fecha_recibido', 'Fecha recibido', 12, 'fecha'),
            ('dias_abiertos', 'Días abiertos', 9, 'dias'), ('reclamos', 'Reclamos (cant.)', 9, 'num'),
            ('ultimo_reclamo', 'Último reclamo', 12, 'fecha'),
            ('historial', 'Historial de reclamos', 30, 'texto'), ('esfuerzo', 'Esfuerzo (días)', 9, 'num'),
            ('link', 'Link / respaldo', 26, 'texto'), ('notas', 'Notas', 40, 'texto'),
            ('clave_origen', 'Clave origen', 22, 'texto'),
        ],
        'cierre': 'fecha_recibido',
    },
    'CE': {
        'titulo': 'Cliente → Estudio', 'prefijo': 'C',
        'nota': 'Lo que el cliente le pide al estudio. Esfuerzo en días: sirve para mostrarle al cliente '
                'todo lo que se hace fuera del abono.',
        'cols': [
            ('id', 'ID', 7, 'texto'), ('fecha_pedido', 'Fecha pedido', 12, 'fecha'),
            ('pedido_por', 'Pedido por (cliente)', 18, 'lista:Cliente'), ('canal', 'Canal', 10, 'lista:Canal'),
            ('responsable', 'Responsable (estudio)', 14, 'lista:Equipo'), ('sociedad', 'Sociedad', 14, 'texto'),
            ('tema', 'Tema', 14, 'lista:Tema'), ('que', 'Qué pidió', 46, 'texto'),
            ('fecha_maxima', 'Fecha máxima', 12, 'fecha'), ('estado', 'Estado', 14, 'lista:Estado'),
            ('fecha_entrega', 'Fecha entrega', 12, 'fecha'), ('dias_abiertos', 'Días abiertos', 9, 'dias'),
            ('reclamos', 'Reclamos recibidos', 9, 'num'), ('ultimo_reclamo', 'Último reclamo', 12, 'fecha'),
            ('esfuerzo', 'Esfuerzo (días)', 9, 'num'), ('link', 'Link / respaldo', 26, 'texto'),
            ('notas', 'Notas', 40, 'texto'), ('clave_origen', 'Clave origen', 22, 'texto'),
        ],
        'cierre': 'fecha_entrega',
    },
    'TI': {
        'titulo': 'Tareas internas', 'prefijo': 'T',
        'nota': 'Tareas que salen de reuniones, Slack o mails internos, con responsable. '
                'Lo que no tiene responsable y fecha, queda en el aire.',
        'cols': [
            ('id', 'ID', 7, 'texto'), ('origen', 'Origen', 22, 'texto'), ('tarea', 'Tarea', 52, 'texto'),
            ('responsable', 'Responsable', 14, 'lista:Equipo'), ('pedido_por', 'Pedido por', 14, 'lista:Equipo'),
            ('fecha_pedido', 'Fecha pedido', 12, 'fecha'), ('fecha_maxima', 'Fecha compromiso', 12, 'fecha'),
            ('estado', 'Estado', 14, 'lista:Estado'), ('fecha_cierre', 'Fecha cierre', 12, 'fecha'),
            ('link', 'Link / respaldo', 26, 'texto'), ('notas', 'Notas', 44, 'texto'),
            ('clave_origen', 'Clave origen', 22, 'texto'),
        ],
        'cierre': 'fecha_cierre',
    },
}
POR_TITULO = {h['titulo']: k for k, h in HOJAS.items()}
OBLIGATORIOS = {'EC': ('que',), 'CE': ('que',), 'TI': ('tarea', 'responsable')}

AZUL = PatternFill('solid', fgColor='1F3864')
GRIS = PatternFill('solid', fgColor='F2F2F2')
ROJO = PatternFill('solid', fgColor='F8CBAD')
VERDE = PatternFill('solid', fgColor='E2EFDA')
AMARILLO = PatternFill('solid', fgColor='FFF2CC')
BLANCO_B = Font(bold=True, color='FFFFFF')
FINO = Side(style='thin', color='BFBFBF')
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)


def fail(msg, **extra):
    print(json.dumps({'ok': False, 'error': msg, **extra}, ensure_ascii=False, indent=1))
    sys.exit(1)


# ── .env (mismo orden que el resto de los plugins) ──────────────────────────
def cargar_env():
    for ruta in (os.environ.get('FISCO_ENV_FILE'), os.path.join(PLUGIN, '.env'),
                 os.path.join(os.path.expanduser('~'), '.fisco-ar', '.env')):
        if not ruta or not os.path.isfile(ruta):
            continue
        with open(ruta, encoding='utf-8') as f:
            for linea in f:
                m = re.match(r'\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)\s*$', linea)
                if not m:
                    continue
                v = m.group(2).strip().strip('"').strip("'")
                if v and not os.environ.get(m.group(1)):
                    os.environ[m.group(1)] = v


def ruta_planilla(args):
    if args.get('ruta'):
        return args['ruta']
    if os.environ.get('PEDIDOS_PLANILLA'):
        return os.environ['PEDIDOS_PLANILLA']
    raiz = (os.environ.get('BVA_UNIDAD_PATH') or '').strip() or RAIZ_DEFECTO
    if not os.path.isdir(raiz):
        fail(f'No encuentro la unidad compartida en "{raiz}". Definí BVA_UNIDAD_PATH en el .env.')
    return os.path.join(raiz, PLANILLA_REL)


def salidas():
    d = os.environ.get('BVA_SALIDAS_PATH') or os.path.join(os.path.expanduser('~'), 'Documents', 'BVA-salidas')
    d = os.path.join(d, 'seguimiento-pedidos')
    os.makedirs(d, exist_ok=True)
    return d


# ── fechas ──────────────────────────────────────────────────────────────────
def a_fecha(v):
    if v in (None, ''):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d/%m/%y', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M:%SZ'):
        try:
            return datetime.strptime(s[:19] if 'T' in s else s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f'Fecha no reconocida: "{v}" (usar AAAA-MM-DD)')


def iso(v):
    d = a_fecha(v) if v not in (None, '') else None
    return d.isoformat() if d else None


# ── crear ───────────────────────────────────────────────────────────────────
def crear_libro():
    wb = Workbook()
    rs = wb.active
    rs.title = 'Resumen'

    def encabezar(ws, fila, textos):
        for j, t in enumerate(textos, start=1):
            c = ws.cell(fila, j, t)
            c.font = BLANCO_B
            c.fill = AZUL
            c.border = BORDE

    # Listas
    ls = wb.create_sheet('Listas')
    listas = {'Estado': ESTADOS, 'Canal': CANALES, 'Tema': TEMAS, 'Cliente': []}
    encabezar(ls, 1, list(listas))
    for j, vals in enumerate(listas.values(), start=1):
        for i, v in enumerate(vals, start=2):
            ls.cell(i, j, v)
        ls.column_dimensions[get_column_letter(j)].width = 24
    ls.cell(1, 6, 'La columna Cliente se completa con las personas del cliente (nombre y apellido).').font = \
        Font(italic=True, color='595959')
    rangos = {'Estado': '=Listas!$A$2:$A$30', 'Canal': '=Listas!$B$2:$B$30', 'Tema': '=Listas!$C$2:$C$40',
              'Cliente': '=Listas!$D$2:$D$80', 'Equipo': '=Equipo!$A$2:$A$40'}

    eq = wb.create_sheet('Equipo')
    encabezar(eq, 1, ['Nombre', 'Email', 'Slack ID', 'Activo'])
    eq.cell(1, 6, 'Una fila por persona del estudio. "Nombre" es el que aparece en las listas de responsables; '
                  'el Email y el Slack ID sirven para saber quién corre la skill.').font = Font(italic=True, color='595959')
    for col, w in zip('ABCD', (16, 30, 14, 8)):
        eq.column_dimensions[col].width = w

    cf = wb.create_sheet('Config')
    encabezar(cf, 1, ['Parámetro', 'Valor', 'Para qué'])
    for i, fila in enumerate([
        ('Dominios del cliente', '', 'Dominios de mail del cliente separados por coma (ej.: cliente.com.ar). '
                                     'Solo esos mails se leen como pedidos del cliente.'),
        ('Días primera corrida', 14, 'Cuántos días hacia atrás se leen la primera vez que alguien corre la skill.'),
        ('Link de la planilla', '', 'Link de esta planilla en Drive, para ponerlo en el mensaje de pendientes.'),
        ('Canales de Slack', '', 'Canales de Slack a leer además de los mensajes que te mencionan y los privados '
                                 '(nombres separados por coma, sin #).'),
    ], start=2):
        for j, v in enumerate(fila, start=1):
            cf.cell(i, j, v)
    for col, w in zip('ABC', (22, 40, 90)):
        cf.column_dimensions[col].width = w

    for k, h in HOJAS.items():
        ws = wb.create_sheet(h['titulo'], index=len(wb.sheetnames) - 3)
        cols = h['cols']
        ws.cell(1, 1, h['nota']).font = Font(italic=True, color='595959')
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
        for j, (_, enc, ancho, _) in enumerate(cols, start=1):
            c = ws.cell(FILA_ENC, j, enc)
            c.font = BLANCO_B
            c.fill = AZUL
            c.border = BORDE
            c.alignment = Alignment(wrap_text=True, vertical='center')
            ws.column_dimensions[get_column_letter(j)].width = ancho
        ws.row_dimensions[FILA_ENC].height = 32
        idx = {c[0]: j for j, c in enumerate(cols, start=1)}
        ultima = FILA_1 + FILAS - 1
        for j, (clave, _, _, tipo) in enumerate(cols, start=1):
            L = get_column_letter(j)
            if tipo == 'fecha':
                for r in range(FILA_1, ultima + 1):
                    ws.cell(r, j).number_format = 'DD/MM/YYYY'
            elif tipo == 'dias':
                F = get_column_letter(idx['fecha_pedido'])
                C = get_column_letter(idx[h['cierre']])
                for r in range(FILA_1, ultima + 1):
                    ws.cell(r, j, f'=IF({F}{r}="","",IF({C}{r}<>"",{C}{r}-{F}{r},TODAY()-{F}{r}))').fill = GRIS
            elif tipo.startswith('lista:'):
                dv = DataValidation(type='list', formula1=rangos[tipo[6:]], allow_blank=True)
                ws.add_data_validation(dv)
                dv.add(f'{L}{FILA_1}:{L}{ultima}')
        E = get_column_letter(idx['estado'])
        M = get_column_letter(idx['fecha_maxima'])
        ref = f'A{FILA_1}:{get_column_letter(len(cols))}{ultima}'
        ws.conditional_formatting.add(ref, FormulaRule(
            formula=[f'OR(${E}{FILA_1}="Completado",${E}{FILA_1}="Cancelado")'], fill=VERDE, stopIfTrue=True))
        ws.conditional_formatting.add(ref, FormulaRule(
            formula=[f'AND(${M}{FILA_1}<>"",${M}{FILA_1}<TODAY())'], fill=ROJO, stopIfTrue=True))
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f'${E}{FILA_1}="A confirmar"'], fill=AMARILLO))
        ws.freeze_panes = f'C{FILA_1}'
        ws.auto_filter.ref = f'A{FILA_ENC}:{get_column_letter(len(cols))}{ultima}'
        ws.column_dimensions[get_column_letter(idx['clave_origen'])].hidden = True

    # Resumen
    rs['A1'] = 'Seguimiento de pedidos Estudio ↔ Cliente — resumen'
    rs['A1'].font = Font(bold=True, size=14)
    rs['A2'] = 'Se actualiza solo. Rojo = pasó la fecha máxima y no está cerrado.'
    rs['A2'].font = Font(italic=True, color='595959')
    ultima = FILA_1 + FILAS - 1

    def rng(k, clave):
        j = [c[0] for c in HOJAS[k]['cols']].index(clave) + 1
        L = get_column_letter(j)
        return f"'{HOJAS[k]['titulo']}'!{L}{FILA_1}:{L}{ultima}"

    orden = ['EC', 'CE', 'TI']
    encabezar(rs, 4, ['Estado'] + [HOJAS[k]['titulo'] for k in orden])
    for i, est in enumerate(ESTADOS, start=5):
        rs.cell(i, 1, est)
        for c, k in enumerate(orden, start=2):
            rs.cell(i, c, f'=COUNTIF({rng(k, "estado")},A{i})')
    f = 5 + len(ESTADOS)
    rs.cell(f, 1, 'Vencidos (abiertos)').font = Font(bold=True, color='C00000')
    for c, k in enumerate(orden, start=2):
        M, E = rng(k, 'fecha_maxima'), rng(k, 'estado')
        rs.cell(f, c, f'=SUMPRODUCT(({M}<>"")*({M}<TODAY())*({E}<>"Completado")*({E}<>"Cancelado"))')
    rs.cell(f + 1, 1, 'Esfuerzo total (días)').font = Font(bold=True)
    rs.cell(f + 1, 2, f'=SUM({rng("EC", "esfuerzo")})')
    rs.cell(f + 1, 3, f'=SUM({rng("CE", "esfuerzo")})')
    f += 3
    encabezar(rs, f, ['Abiertos por responsable'] + [HOJAS[k]['titulo'] for k in orden])
    for i in range(f + 1, f + 21):
        rs.cell(i, 1, f'=IF(Equipo!A{i - f + 1}="","",Equipo!A{i - f + 1})')
        for c, k in enumerate(orden, start=2):
            R, E = rng(k, 'responsable'), rng(k, 'estado')
            rs.cell(i, c, f'=IF($A{i}="","",SUMPRODUCT(({R}=$A{i})*({E}<>"Completado")*({E}<>"Cancelado")*({E}<>"")))')
    for col, w in zip('ABCD', (26, 18, 18, 18)):
        rs.column_dimensions[col].width = w

    ay = wb.create_sheet('Cómo usarla', index=4)
    for i, t in enumerate([
        'Cómo usar esta planilla',
        '',
        '1. Todo pedido al cliente o del cliente se anota acá el mismo día, con quién lo pidió y quién se hace cargo.',
        '2. Antes de pedirle algo al cliente, filtrar por Tema o Sociedad: si ya está pedido, no se pide de nuevo; se suma un reclamo.',
        '3. Cada reclamo: sumar 1 en «Reclamos», poner la fecha en «Último reclamo» y anotarlo en el historial.',
        '4. «Fecha máxima» la ponen las dos partes. Si pasa y no está cerrado, la fila se pone en rojo.',
        '5. «Esfuerzo (días)»: estimar cuánto llevó. Sirve para mostrarle al cliente el trabajo fuera del abono.',
        '6. Al cerrar, poner Estado = Completado y la fecha. La fila se pone en verde.',
        '',
        'La skill «seguimiento-pedidos» lee tus mails con el cliente y tu Slack, carga acá lo nuevo y te manda tus pendientes.',
        'La columna oculta «Clave origen» es la que usa la skill para no cargar dos veces el mismo mail: no la borres.',
        'No convertir este archivo a Hoja de cálculo de Google: la skill escribe sobre el .xlsx.',
    ], start=1):
        ay.cell(i, 1, t)
    ay['A1'].font = Font(bold=True, size=13)
    ay.column_dimensions['A'].width = 110

    co = wb.create_sheet('Corridas')
    encabezar(co, 1, ['Fecha y hora', 'Persona', 'Desde', 'Hasta', 'Agregados', 'Actualizados', 'Reclamos', 'Omitidos'])
    for col, w in zip('ABCDEFGH', (18, 14, 12, 12, 10, 12, 10, 10)):
        co.column_dimensions[col].width = w
    return wb


def cmd_crear(a):
    ruta = ruta_planilla(a)
    if os.path.exists(ruta):
        fail(f'Ya existe {ruta}. No la piso.')
    crear_libro().save(ruta)
    print(json.dumps({'ok': True, 'creada': ruta}, ensure_ascii=False))


# ── lectura ─────────────────────────────────────────────────────────────────
def abrir(ruta):
    if not os.path.isfile(ruta):
        fail(f'No encuentro la planilla en {ruta}. Creala con: pedidos.py crear')
    try:
        return load_workbook(ruta)
    except Exception as e:  # noqa: BLE001
        fail(f'No pude abrir {ruta}: {e}')


def mapa_columnas(ws, k):
    enc = {str(c.value).strip(): c.column for c in ws[FILA_ENC] if c.value}
    m = {}
    for clave, titulo, _, _ in HOJAS[k]['cols']:
        if titulo not in enc:
            fail(f'En la hoja "{HOJAS[k]["titulo"]}" falta la columna "{titulo}". No cambies los encabezados.')
        m[clave] = enc[titulo]
    return m


def filas(wb, k):
    ws = wb[HOJAS[k]['titulo']]
    m = mapa_columnas(ws, k)
    out = []
    for r in range(FILA_1, ws.max_row + 1):
        if ws.cell(r, m['id']).value in (None, ''):
            continue
        d = {c: ws.cell(r, j).value for c, j in m.items() if c != 'dias_abiertos'}
        d['_fila'] = r
        out.append(d)
    return ws, m, out


def equipo(wb):
    if 'Equipo' not in wb.sheetnames:
        return []
    ws = wb['Equipo']
    out = []
    for r in range(2, ws.max_row + 1):
        n = ws.cell(r, 1).value
        if n:
            out.append({'nombre': str(n).strip(), 'email': (ws.cell(r, 2).value or '').strip().lower(),
                        'slack': (ws.cell(r, 3).value or '').strip(),
                        'activo': str(ws.cell(r, 4).value or 'Sí').strip().lower() not in ('no', 'n', '0')})
    return out


def config(wb):
    c = {}
    if 'Config' in wb.sheetnames:
        for r in range(2, wb['Config'].max_row + 1):
            k = wb['Config'].cell(r, 1).value
            if k:
                c[str(k).strip()] = wb['Config'].cell(r, 2).value
    lista = lambda s: [x.strip() for x in str(s or '').split(',') if x.strip()]
    return {'dominios_cliente': [d.lstrip('@').lower() for d in lista(c.get('Dominios del cliente'))],
            'dias_primera_corrida': int(c.get('Días primera corrida') or 14),
            'link': c.get('Link de la planilla') or '',
            'canales_slack': [x.lstrip('#') for x in lista(c.get('Canales de Slack'))]}


def lista_col(wb, titulo):
    ws = wb['Listas']
    for c in ws[1]:
        if c.value == titulo:
            return [ws.cell(r, c.column).value for r in range(2, ws.max_row + 1) if ws.cell(r, c.column).value]
    return []


def resolver_persona(wb, a):
    eq = equipo(wb)
    if a.get('persona'):
        p = a['persona'].strip().lower()
        for e in eq:
            if e['nombre'].lower() == p:
                return e
        fail(f'"{a["persona"]}" no está en la hoja Equipo.', equipo=[e['nombre'] for e in eq])
    for campo, val in (('email', (a.get('email') or '').lower()), ('slack', a.get('slack') or '')):
        if val:
            for e in eq:
                if e[campo] and e[campo] == val:
                    return e
    fail('No sé quién sos: tu mail o Slack ID no está en la hoja Equipo de la planilla. '
         'Agregate ahí (Nombre, Email, Slack ID) o pasá --persona.', equipo=[e['nombre'] for e in eq])


def json_val(v):
    if isinstance(v, (datetime, date)):
        return iso(v)
    return v


def cmd_estado(a):
    ruta = ruta_planilla(a)
    wb = abrir(ruta)
    yo = resolver_persona(wb, a)
    cfg = config(wb)
    ultima = None
    if 'Corridas' in wb.sheetnames:
        ws = wb['Corridas']
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(r, 2).value or '').strip().lower() == yo['nombre'].lower():
                h = a_fecha(ws.cell(r, 4).value)
                if h and (not ultima or h > ultima):
                    ultima = h
    hoy = date.today()
    desde = ultima if ultima else hoy - timedelta(days=cfg['dias_primera_corrida'])
    claves, abiertos = [], []
    for k in HOJAS:
        _, _, fs = filas(wb, k)
        for f in fs:
            if f.get('clave_origen'):
                claves.extend(str(f['clave_origen']).split('|'))
            if f.get('estado') not in CERRADOS:
                abiertos.append({'hoja': k, **{c: json_val(v) for c, v in f.items()
                                               if c != '_fila' and v not in (None, '')}})
    print(json.dumps({
        'ok': True, 'planilla': ruta, 'persona': yo['nombre'], 'email': yo['email'], 'slack': yo['slack'],
        'desde': desde.isoformat(), 'hasta': hoy.isoformat(), 'primera_corrida': ultima is None,
        'config': cfg, 'equipo': [e for e in equipo(wb) if e['activo']],
        'cliente': lista_col(wb, 'Cliente'), 'temas': lista_col(wb, 'Tema'), 'canales': lista_col(wb, 'Canal'),
        'estados': lista_col(wb, 'Estado'), 'claves_cargadas': sorted(set(claves)), 'abiertos': abiertos,
    }, ensure_ascii=False, indent=1))


# ── aplicar ─────────────────────────────────────────────────────────────────
def siguiente_id(fs, prefijo):
    n = 0
    for f in fs:
        m = re.match(rf'{prefijo}-(\d+)$', str(f['id']))
        if m:
            n = max(n, int(m.group(1)))
    return f'{prefijo}-{n + 1:03d}'


def fila_libre(ws, m):
    r = FILA_1
    while ws.cell(r, m['id']).value not in (None, ''):
        r += 1
    return r


def escribir(ws, m, r, k, datos, pisar=False):
    tipos = {c[0]: c[3] for c in HOJAS[k]['cols']}
    hechos = {}
    for c, v in datos.items():
        if c not in m or c in ('id', 'dias_abiertos', '_fila'):
            continue
        if v in (None, ''):
            continue
        celda = ws.cell(r, m[c])
        if not pisar and celda.value not in (None, '') and c not in ('estado',):
            if c in ('notas', 'link'):
                v = f'{celda.value} | {v}' if str(v) not in str(celda.value) else celda.value
            elif tipos[c] == 'fecha' and c != 'fecha_pedido':
                pass  # se actualiza abajo
            else:
                continue
        if tipos[c] == 'fecha':
            v = a_fecha(v)
            celda.number_format = 'DD/MM/YYYY'
        elif tipos[c] == 'num':
            v = float(v) if '.' in str(v) else int(v)
        elif tipos[c] == 'lista:Estado' and v not in ESTADOS:
            raise ValueError(f'Estado inválido "{v}" (válidos: {", ".join(ESTADOS)})')
        celda.value = v
        hechos[c] = json_val(v)
    return hechos


def cmd_aplicar(a, archivo):
    try:
        with open(archivo, encoding='utf-8') as f:
            pedido = json.load(f)
    except Exception as e:  # noqa: BLE001
        fail(f'No pude leer {archivo}: {e}')
    ruta = ruta_planilla(a)
    mtime = os.path.getmtime(ruta) if os.path.exists(ruta) else None
    wb = abrir(ruta)
    persona = resolver_persona(wb, {'persona': pedido.get('persona')})['nombre'] if pedido.get('persona') else None
    hojas = {k: filas(wb, k) for k in HOJAS}
    claves = {}
    for k, (_, _, fs) in hojas.items():
        for f in fs:
            for c in str(f.get('clave_origen') or '').split('|'):
                if c:
                    claves[c] = f['id']
    por_id = {str(f['id']): (k, f) for k, (_, _, fs) in hojas.items() for f in fs}
    res = {'agregados': [], 'actualizados': [], 'reclamos': [], 'omitidos': [], 'errores': []}
    hoy = date.today()

    for i, op in enumerate(pedido.get('ops', []), start=1):
        acc = op.get('accion')
        try:
            if acc == 'agregar':
                k = op.get('hoja')
                if k not in HOJAS:
                    k = POR_TITULO.get(k)
                if not k:
                    raise ValueError(f'hoja inválida "{op.get("hoja")}" (EC, CE o TI)')
                datos = dict(op.get('datos') or {})
                clave = (op.get('clave_origen') or datos.get('clave_origen') or '').strip()
                if not clave:
                    raise ValueError('falta clave_origen (ej.: gmail:<threadId>#1 o slack:<canal>/<ts>)')
                if clave in claves:
                    res['omitidos'].append({'op': i, 'motivo': f'ya cargado como {claves[clave]}', 'clave': clave})
                    continue
                for ob in OBLIGATORIOS[k]:
                    if not datos.get(ob):
                        raise ValueError(f'falta "{ob}"')
                ws, m, fs = hojas[k]
                datos.setdefault('estado', 'Pendiente')
                datos.setdefault('fecha_pedido', hoy.isoformat())
                if k != 'CE':
                    datos.setdefault('pedido_por', persona)
                datos.setdefault('responsable', persona)
                if k in ('EC', 'CE'):
                    datos.setdefault('reclamos', 0)
                datos['clave_origen'] = clave
                nuevo = siguiente_id(fs, HOJAS[k]['prefijo'])
                r = fila_libre(ws, m)
                ws.cell(r, m['id'], nuevo)
                hechos = escribir(ws, m, r, k, datos, pisar=True)
                fs.append({'id': nuevo, '_fila': r, **datos})
                por_id[nuevo] = (k, fs[-1])
                claves[clave] = nuevo
                res['agregados'].append({'id': nuevo, 'hoja': HOJAS[k]['titulo'],
                                         'que': hechos.get('que') or hechos.get('tarea'),
                                         'responsable': hechos.get('responsable'), 'estado': hechos.get('estado')})
            elif acc in ('actualizar', 'reclamo'):
                ident = str(op.get('id') or '')
                if ident not in por_id:
                    raise ValueError(f'no existe el ID "{ident}"')
                k, f = por_id[ident]
                ws, m, _ = hojas[k]
                r = f['_fila']
                if acc == 'actualizar':
                    datos = dict(op.get('datos') or {})
                    if op.get('clave_origen'):
                        prev = str(ws.cell(r, m['clave_origen']).value or '')
                        if op['clave_origen'] in prev.split('|'):
                            res['omitidos'].append({'op': i, 'motivo': 'esa novedad ya estaba aplicada', 'id': ident})
                            continue
                        datos['clave_origen'] = f'{prev}|{op["clave_origen"]}' if prev else op['clave_origen']
                        claves[op['clave_origen']] = ident
                    hechos = escribir(ws, m, r, k, datos, pisar=bool(op.get('pisar')))
                    if 'clave_origen' in datos:
                        ws.cell(r, m['clave_origen'], datos['clave_origen'])
                    res['actualizados'].append({'id': ident, 'cambios': {c: v for c, v in hechos.items()
                                                                           if c != 'clave_origen'}})
                else:
                    if k == 'TI':
                        raise ValueError('las tareas internas no llevan reclamos: usar "actualizar" con notas')
                    fecha = a_fecha(op.get('fecha') or hoy)
                    clave = (op.get('clave_origen') or '').strip()
                    prev = str(ws.cell(r, m['clave_origen']).value or '')
                    if clave and clave in prev.split('|'):
                        res['omitidos'].append({'op': i, 'motivo': 'ese reclamo ya estaba cargado', 'id': ident})
                        continue
                    n = int(ws.cell(r, m['reclamos']).value or 0) + 1
                    ws.cell(r, m['reclamos'], n)
                    ws.cell(r, m['ultimo_reclamo'], fecha).number_format = 'DD/MM/YYYY'
                    linea = f'{fecha:%d/%m}: {op.get("detalle") or "reclamo"}'
                    destino = 'historial' if 'historial' in m else 'notas'
                    actual = ws.cell(r, m[destino]).value
                    ws.cell(r, m[destino], f'{actual} | {linea}' if actual else linea)
                    if clave:
                        ws.cell(r, m['clave_origen'], f'{prev}|{clave}' if prev else clave)
                        claves[clave] = ident
                    res['reclamos'].append({'id': ident, 'reclamos': n, 'fecha': fecha.isoformat()})
            else:
                raise ValueError(f'acción inválida "{acc}" (agregar, actualizar o reclamo)')
        except (ValueError, TypeError) as e:
            res['errores'].append({'op': i, 'error': str(e), 'detalle': op})

    if a.get('solo-revisar'):
        print(json.dumps({'ok': True, 'solo_revisar': True, 'planilla': ruta, **res}, ensure_ascii=False, indent=1))
        return

    if 'Corridas' in wb.sheetnames and persona:
        ws = wb['Corridas']
        r = ws.max_row + 1
        for j, v in enumerate([datetime.now().strftime('%Y-%m-%d %H:%M'), persona,
                               a_fecha(pedido.get('desde')) if pedido.get('desde') else None,
                               a_fecha(pedido.get('hasta') or hoy), len(res['agregados']),
                               len(res['actualizados']), len(res['reclamos']), len(res['omitidos'])], start=1):
            c = ws.cell(r, j, v)
            if isinstance(v, date):
                c.number_format = 'DD/MM/YYYY'

    respaldo = os.path.join(salidas(), 'respaldos',
                            f'{datetime.now():%Y%m%d-%H%M%S} - {os.path.basename(ruta)}')
    os.makedirs(os.path.dirname(respaldo), exist_ok=True)
    shutil.copy2(ruta, respaldo)
    if mtime is not None and os.path.getmtime(ruta) != mtime:
        fail('La planilla cambió mientras la procesaba (alguien la guardó). No escribí nada: volvé a correr.')
    tmp = os.path.join(tempfile.gettempdir(), f'pedidos-{os.getpid()}.xlsx')
    wb.save(tmp)
    try:
        shutil.copyfile(tmp, ruta)  # copia sobre el mismo archivo: Drive conserva el link
    except PermissionError:
        fail('No pude guardar: la planilla está abierta en Excel en esta compu. Cerrala y volvé a correr.',
             respaldo=respaldo)
    finally:
        os.remove(tmp)
    print(json.dumps({'ok': True, 'planilla': ruta, 'respaldo': respaldo, **res}, ensure_ascii=False, indent=1))


# ── pendientes ──────────────────────────────────────────────────────────────
def cmd_pendientes(a):
    ruta = ruta_planilla(a)
    wb = abrir(ruta)
    yo = resolver_persona(wb, a)['nombre']
    hoy = a_fecha(a.get('hoy') or date.today())
    cfg = config(wb)
    items = []
    for k in HOJAS:
        _, _, fs = filas(wb, k)
        for f in fs:
            if f.get('estado') in CERRADOS:
                continue
            resp = str(f.get('responsable') or '').strip()
            if resp.lower() != yo.lower() and not (k == 'TI' and resp.lower() == 'el grupo'):
                continue
            fp, fm = a_fecha(f.get('fecha_pedido')), a_fecha(f.get('fecha_maxima'))
            items.append({
                'id': f['id'], 'hoja': k, 'estado': f.get('estado') or 'Pendiente',
                'texto': f.get('que') or f.get('tarea'),
                'con': f.get('a_quien') or ('' if str(f.get('pedido_por') or '').lower() == yo.lower()
                                            else f.get('pedido_por') or ''),
                'sociedad': f.get('sociedad') or '', 'fecha_pedido': iso(fp), 'fecha_maxima': iso(fm),
                'dias': (hoy - fp).days if fp else None, 'reclamos': int(f.get('reclamos') or 0),
                'vencido': bool(fm and fm < hoy), 'vence_pronto': bool(fm and hoy <= fm <= hoy + timedelta(days=2)),
            })
    items.sort(key=lambda x: (not x['vencido'], x['fecha_maxima'] or '9999', x['fecha_pedido'] or '9999'))
    fmt = a.get('formato', 'json')
    if fmt == 'json':
        print(json.dumps({'ok': True, 'persona': yo, 'hoy': hoy.isoformat(), 'total': len(items),
                          'vencidos': sum(i['vencido'] for i in items), 'items': items}, ensure_ascii=False, indent=1))
        return
    slack = fmt == 'slack'
    b = (lambda s: f'**{s}**') if slack else (lambda s: s.upper())  # el conector de Slack usa markdown estándar

    def renglon(x):
        partes = [f'{x["id"]}', x['texto'] or '']
        if x['con']:
            partes.append(f'({x["con"]}{" · " + x["sociedad"] if x["sociedad"] else ""})')
        extra = []
        if x['vencido']:
            extra.append(f'venció {datetime.fromisoformat(x["fecha_maxima"]):%d/%m}')
        elif x['fecha_maxima']:
            extra.append(f'vence {datetime.fromisoformat(x["fecha_maxima"]):%d/%m}')
        if x['dias']:
            extra.append(f'{x["dias"]} día{"s" if x["dias"] > 1 else ""} abierto')
        if x['reclamos']:
            extra.append(f'{x["reclamos"]} reclamo{"s" if x["reclamos"] > 1 else ""}')
        if x['estado'] not in ('Pendiente',):
            extra.append(x['estado'].lower())
        marca = ':red_circle: ' if (slack and x['vencido']) else (':large_yellow_circle: ' if slack and x['vence_pronto'] else '• ')
        return f'{marca}{" ".join(partes)}' + (f' — {", ".join(extra)}' if extra else '')

    lineas = [b(f'Tus pendientes al {hoy:%d/%m/%Y}') + f' ({len(items)})']
    if not items:
        lineas.append('No tenés nada abierto en la planilla de seguimiento. :tada:' if slack else 'No hay pendientes.')
    for k, titulo in (('EC', 'Le pediste al cliente (hacer seguimiento)'),
                      ('CE', 'Te pidió el cliente (entregar)'), ('TI', 'Tareas internas')):
        grupo = [x for x in items if x['hoja'] == k]
        if grupo:
            lineas += ['', b(titulo)] + [renglon(x) for x in grupo]
    if cfg['link']:
        lineas += ['', f'Planilla: {cfg["link"]}']
    print('\n'.join(lineas))


def parse(argv):
    a, pos = {}, []
    for x in argv:
        if x.startswith('--'):
            k, _, v = x[2:].partition('=')
            a[k] = v if v else True
        else:
            pos.append(x)
    return a, pos


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    cargar_env()
    a, pos = parse(sys.argv[1:])
    if not pos:
        print(__doc__)
        sys.exit(1)
    cmd = pos[0]
    if cmd == 'crear':
        cmd_crear(a)
    elif cmd == 'estado':
        cmd_estado(a)
    elif cmd == 'aplicar':
        if len(pos) < 2:
            fail('Uso: pedidos.py aplicar <ops.json> [--solo-revisar]')
        cmd_aplicar(a, pos[1])
    elif cmd == 'pendientes':
        cmd_pendientes(a)
    else:
        fail(f'Subcomando desconocido "{cmd}" (crear, estado, aplicar, pendientes)')


if __name__ == '__main__':
    main()
