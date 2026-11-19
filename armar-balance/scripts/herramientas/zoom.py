"""Recorta y amplía páginas de un PDF escaneado para leer cifras: python zoom.py <pdf> <prefijo salida> "<pag>:<rot>:y0:y1:x0:x1[,...]"."""
import sys, pypdfium2 as pdfium
from PIL import Image
pdf_path, out, pages = sys.argv[1], sys.argv[2], sys.argv[3]
pdf = pdfium.PdfDocument(pdf_path)
for spec in pages.split(","):
    p, rot, y0, y1, x0, x1 = spec.split(":")
    im = pdf[int(p) - 1].render(scale=2.2).to_pil().convert("L")
    if rot != "0": im = im.rotate(int(rot), expand=True)
    w, h = im.size
    im.crop((int(w * float(x0)), int(h * float(y0)), int(w * float(x1)), int(h * float(y1)))).save(f"{out}_p{p}.png")
