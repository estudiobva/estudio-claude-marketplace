# Extrae del F.931 (PDF) los datos de la hoja Sueldos: 301, 302, 351, 352, 312, empleados, Remun. con ART, 028.
# Caso normal: sección VIII. Rectificativa con la VIII en cero: secciones I, II y VI (criterio skill f931-sueldos).
import re
from pypdf import PdfReader
def num(s): return float(s.replace('.','').replace(',','.'))
N=r'(-?[\d\.]+,\d{2})'
def leer(ruta):
    t=PdfReader(ruta).pages[0].extract_text()
    g=lambda pat: (lambda m: num(m.group(1)) if m else None)(re.search(pat,t))
    rect=int((re.search(r'Rect\. \(1/9\):\s*(\d)',t) or [0,0])[1])
    per=(re.search(r'\n(\d{2}/\d{4})',t) or [None,None])[1]
    cuit=(re.search(r'C\.U\.I\.T\.\s*([\d-]+)',t) or [None,''])[1].replace('-','')
    d=dict(rect=rect, periodo=per, cuit=cuit,
      empleados=int((re.search(r'Empleados en n[oó]mina:\s*(\d+)',t) or [0,0])[1]),
      rem_art=g(r'Remun\. con ART\s+'+N),
      c301=g(r'301 - Aportes de Seguridad Social\s+'+N), c302=g(r'302 - Aportes de Obra Social\s+'+N),
      c351=g(r'351 - Contribuciones de Seguridad Social\s+'+N), c352=g(r'352 - Contribuciones de Obra Social\s+'+N),
      c312=g(r'312 - L\.R\.T\.\s+'+N), c028=g(r'028 - Seguro Colectivo de Vida Obligatorio\s+'+N),
      costo_emision=g(r'Costo Emisi[oó]n:\s*'+N), fuente='VIII')
    viii=[d[k] for k in ('c301','c302','c351','c352','c312','c028')]
    if rect>0 and all((v or 0)==0 for v in viii):
        d.update(c301=g(r'a3 - Aportes S\.S\. a pagar\s+'+N), c302=g(r'a3 - Aportes O\.S\. a pagar\s+'+N),
                 c351=g(r'Contribuciones S\.S\. a pagar\s+'+N), c352=g(r'Contribuciones O\.S\. a pagar\s+'+N),
                 c312=g(r'L\.R\.T\. total a pagar\s+'+N), c028=g(r'S\.C\.V\.O\. a Pagar:\s*'+N), fuente='I/II/VI (rectificativa con VIII en cero)')
    faltan=[k for k in ('c301','c302','c351','c352','c312','c028','rem_art') if d[k] is None]
    d['faltan']=faltan
    return d
if __name__=='__main__':
    import sys,json
    for f in sys.argv[1:]: print(f.split('/')[-1], json.dumps(leer(f),ensure_ascii=False))
