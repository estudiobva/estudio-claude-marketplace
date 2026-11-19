#!/usr/bin/env python3
"""Genera el Balance de las sociedades de config/*.json (o las pedidas) en '0 - ESTUDIO/EERR 2026'.
Uso: generar_lote.py [nombre-config ...] [--destino=<carpeta>] [--eerr-dir=<carpeta con los EERR ya calculados>]"""
import glob, json, os, shutil, subprocess, sys, tempfile
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import fuentes  # resuelve la unidad (BVA_UNIDAD_PATH) y la carpeta de configs
U = fuentes.UNIDAD
EERR_DIR = next((a.split("=", 1)[1] for a in sys.argv[1:] if a.startswith("--eerr-dir=")), os.path.join(U, "0 - ESTUDIO - 00000000000", "EERR 2026", "EECC anterior"))
args = [a for a in sys.argv[1:] if not a.startswith("--")]
dest = next((a.split("=", 1)[1] for a in sys.argv[1:] if a.startswith("--destino=")), os.path.join(U, "0 - ESTUDIO - 00000000000", "EERR 2026"))
CFG_DIR = next((a.split("=", 1)[1] for a in sys.argv[1:] if a.startswith("--config-dir=")), fuentes.CONFIG_DIR)
if not os.path.isdir(CFG_DIR):
    sys.exit(f"No encuentro la carpeta de configs: {CFG_DIR} (BALANCE_CONFIG_DIR o --config-dir)")
cfgs = [os.path.join(CFG_DIR, a + ".json") for a in args] or sorted(glob.glob(os.path.join(CFG_DIR, "*.json")))
tmp = tempfile.mkdtemp()
for cf in cfgs:
    c = json.load(open(cf, encoding="utf-8"))
    from datetime import date
    C = date.fromisoformat(c.get("cierre", "2026-11-30")); ini = (C.year * 12 + C.month - 12)  # mes de inicio (índice absoluto)
    per = f"{ini % 12 + 1:02d}-{ini // 12} a {C.month:02d}-{C.year}"
    eerr = os.path.join(EERR_DIR, f"EERR {per} - {c['nombre']}.xlsx")
    local = os.path.join(tmp, os.path.basename(eerr)); shutil.copy2(eerr, local)
    out = os.path.join(dest, f"BALANCE {C.strftime('%d-%m-%Y')} - {c['nombre']} - PROTOTIPO.xlsx")
    r = subprocess.run([sys.executable, os.path.join(AQUI, "armar_balance.py"), local, out, cf], capture_output=True, text=True)
    print(("OK  " if r.returncode == 0 else "ERR ") + c["nombre"], "" if r.returncode == 0 else r.stderr.strip().splitlines()[-1])
