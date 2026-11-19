"""Letra de cada solapa del papel de trabajo (pedido 2026-10-28): se mantiene la del original.

    Ventas Zetti / Ventas Zetti (depurado)  -> Aptos Narrow 11 (como sale el reporte de Zetti)
    Ventas portal iva / Ventas MCR          -> Arial 10 (como salen de ARCA)
    cuadro resumen de Ventas portal iva     -> Aptos Narrow 11

aplicar(ws, ZETTI) cambia nombre y tamano de la letra y conserva negrita, cursiva y color.
"""
from copy import copy

from openpyxl.styles import Font

ZETTI = ('Aptos Narrow', 11)
ARCA = ('Arial', 10)


def letra(nombre_tam, **kw):
    """Font nueva con la letra indicada (y bold/italic/color si se pasan)."""
    return Font(name=nombre_tam[0], size=nombre_tam[1], **kw)


def aplicar(ws, nombre_tam, celdas=None):
    """Aplica la letra a las celdas dadas o, sin `celdas`, a todas las que tienen valor."""
    if celdas is None:
        celdas = (c for fila in ws.iter_rows() for c in fila if c.value is not None)
    for c in celdas:
        f = copy(c.font)
        f.name, f.sz = nombre_tam
        # Sin esto la letra queda como "fuente del tema" (scheme minor) y Excel, al guardar, la
        # reemplaza por la del tema (Calibri): paso con Compras Zetti en un papel de 09/2026.
        f.scheme = None
        c.font = f
