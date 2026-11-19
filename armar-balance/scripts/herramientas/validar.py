"""Valida Balances generados recalculando las fórmulas con pycel (sin Excel): errores de fórmulas, ESP comparativo, Anexo 2.
Uso: python validar.py <carpeta o archivo BALANCE ...xlsx>   (requiere: pip install pycel)"""
import sys, glob, os, openpyxl
from pycel import ExcelCompiler
_a = sys.argv[1]
_fs = [_a] if _a.lower().endswith(".xlsx") else glob.glob(os.path.join(_a, "BALANCE *.xlsx"))
for f in sorted(x for x in _fs if not os.path.basename(x).startswith("~$")):
    x = ExcelCompiler(filename=f)
    def ev(r):
        try: return x.evaluate(r)
        except Exception: return "ERR"
    nom = os.path.basename(f).split(" - ")[1]
    wb = openpyxl.load_workbook(f); errs = []
    for sh in ("Situación Patrimonial", "Estado de Resultados", "Flujo de Efectivo", "Anexo 2", "Anexo 3", "Evolución PN", "Nota 2", "Resumen Patrimonial"):
        for row in wb[sh].iter_rows(min_row=6):
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    v = ev(f"'{sh}'!{c.coordinate}")
                    if v == "ERR" or (isinstance(v, str) and v.startswith("#")): errs.append(f"{sh}!{c.coordinate}")
    er = lambda r: ev(f"'Estado de Resultados'!B{r}") or 0
    print(f"{nom:20s} ventas={er(6):>16,.0f} gastos={-(er(9)+er(10)+er(11)):>14,.0f} resultado={er(17):>15,.0f} "
          f"ESP25dif={ev(chr(39)+'Situación Patrimonial'+chr(39)+'!C32')} A2={ev('Controles!C9')} pend={ev('Controles!C11')} prov={ev('Controles!C12')} errores={len(errs)} {errs[:3]}")
