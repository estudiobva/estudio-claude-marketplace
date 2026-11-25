#!/usr/bin/env python3
"""
Aplica clasificaciones a proveedores que quedaron SIN CLASIFICAR.

Uso:
  python agregar_clasificaciones.py --compras <input_portal_COMPRAS.xlsx> --padron <Clasificacion_Proveedores.xlsx> \
      --json <clasificaciones.json> --modo sugerida|confirmada

clasificaciones.json: lista de objetos
  {"cuit": 30000000000, "clasificacion": "Viáticos y movilidad",
   "nombre": "CONCESIONARIA VIAL S.A.",             (opcional, para el padrón)
   "actividad": "Concesionaria vial (peajes)",      (opcional)
   "confianza": "Alta|Media|Baja",                  (opcional)
   "fuente": "Vialidad Nacional"}                   (opcional; "Usuario" si la dio el usuario)

--modo sugerida   : escribe la clasificación en COMPRAS resaltada (amarillo = alta/media, naranja = baja)
                    con comentario de actividad y fuente. NO toca el padrón.
--modo confirmada : escribe la clasificación sin resaltado, la agrega al padrón (si el CUIT no está)
                    y marca Estado = "Agregado al padrón".
Los archivos se modifican en el lugar. Después: recalcular COMPRAS con recalc.py del skill xlsx.
"""
import argparse, json, copy, sys, os
sys.path.insert(0, os.path.dirname(__file__))
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from procesar_portal_iva import canon, YELLOW, ORANGE, hoja_resumen

NOFILL = PatternFill(fill_type=None)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--compras', required=True)
    ap.add_argument('--padron', required=True)
    ap.add_argument('--json', required=True)
    ap.add_argument('--modo', choices=['sugerida', 'confirmada'], required=True)
    a = ap.parse_args()
    items = json.load(open(a.json, encoding='utf-8'))
    by = {int(i['cuit']): dict(i, clasificacion=canon(i['clasificacion'])) for i in items}

    wb = openpyxl.load_workbook(a.compras)
    ws = wb['PORTAL IVA']
    nombres, n = {}, 0
    for r in range(3, ws.max_row + 1):
        cu = ws.cell(r, 5).value
        if cu not in by:
            continue
        nombres.setdefault(cu, ws.cell(r, 6).value)
        i = by[cu]; m = ws.cell(r, 13); m.value = i['clasificacion']; n += 1
        if a.modo == 'sugerida':
            conf = i.get('confianza', 'Media')
            m.fill = ORANGE if conf == 'Baja' else YELLOW
            m.comment = Comment(f"Clasificación SUGERIDA (confianza {conf.lower()}). {i.get('actividad', '')}. Fuente: {i.get('fuente', '-')}.", 'Claude')
        else:
            m.fill = NOFILL; m.comment = None
    if 'Proveedores sin clasificar' in wb.sheetnames:
        s = wb['Proveedores sin clasificar']
        for r in range(2, s.max_row + 1):
            cu = s.cell(r, 1).value
            if cu not in by:
                continue
            i = by[cu]
            if i.get('actividad'): s.cell(r, 5, i['actividad'])
            s.cell(r, 6, i['clasificacion'])
            s.cell(r, 7, i.get('confianza', '-'))
            s.cell(r, 8, i.get('fuente', '-'))
            if a.modo == 'sugerida':
                s.cell(r, 6).fill = ORANGE if i.get('confianza') == 'Baja' else YELLOW
                s.cell(r, 9, 'Sugerida')
            else:
                s.cell(r, 6).fill = NOFILL
                s.cell(r, 9, 'Agregado al padrón')
    hoja_resumen(wb)
    wb.save(a.compras)
    print('comprobantes actualizados:', n, '- resumen reconstruido')

    if a.modo == 'confirmada':
        wb = openpyxl.load_workbook(a.padron); ws = wb.active
        existing = {ws.cell(r, 2).value for r in range(2, ws.max_row + 1)}
        ref = [ws.cell(2, j) for j in range(1, 4)]
        r0 = ws.max_row + 1; k = 0
        for cu, i in by.items():
            if cu in existing:
                continue
            row = [i.get('nombre') or nombres.get(cu) or '', cu, i['clasificacion']]
            for j, v in enumerate(row):
                c = ws.cell(r0 + k, j + 1, v); c.font = copy.copy(ref[j].font); c.number_format = ref[j].number_format
            k += 1
        wb.save(a.padron)
        print(f'padrón: {k} filas agregadas ({r0} a {r0 + k - 1})' if k else 'padrón: sin filas nuevas')


if __name__ == '__main__':
    main()
