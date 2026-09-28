// bundle-ncr.js — se ejecuta DENTRO de la pestaña del SharePoint, en el perfil de
// Chrome "Trabajo" (estudiobva), que ya tiene la sesion de invitado abierta.
// No hay otra forma de entrar: el acceso es con la cuenta de Google que
// autorizaron las farmacias y no hay claves para usar desde un script.
//
// Antes de pegarlo, definir el periodo:   window.__NCR_PERIODO = '202608';
//
// Hace solo GET (lectura). Baja SUIZO/ + MONROE/ de NC por NRF/<AAAA>/<MES>/ y
// los entrega como UNA descarga, SP-NCR-<AAAAMM>.bundle, porque Chrome frena las
// descargas multiples automaticas. Formato: 4 bytes (largo del indice, big
// endian) + indice JSON + los archivos concatenados. Lo abre descargar-ncr.js.
(async () => {
  const periodo = String(window.__NCR_PERIODO || '');
  if (!/^\d{6}$/.test(periodo)) throw new Error('Defini window.__NCR_PERIODO = "AAAAMM" antes de correr esto.');
  const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
                 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
  const sitio = '/sites/EstudioContable';
  const h = { headers: { Accept: 'application/json;odata=nometadata' } };
  const q = (p) => p.replace(/'/g, "''");
  const norm = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
  const api = async (ruta) => {
    const r = await fetch(`${sitio}/_api/${ruta}`, h);
    if (!r.ok) throw new Error(`SharePoint ${r.status} en ${ruta}`);
    return (await r.json()).value || [];
  };
  const carpetas = (p) => api(`web/GetFolderByServerRelativePath(decodedurl='${q(p)}')/Folders?$select=Name`);
  const archivos = (p) => api(`web/GetFolderByServerRelativePath(decodedurl='${q(p)}')/Files?$select=Name,Length,TimeLastModified`);

  const anio = periodo.slice(0, 4), mes = MESES[Number(periodo.slice(4)) - 1];
  const baseAnio = `${sitio}/Shared Documents/NC por NRF/${anio}`;
  const meses = (await carpetas(baseAnio)).map((x) => x.Name);
  const carpetaMes = meses.find((m) => norm(m) === mes || (mes === 'SEPTIEMBRE' && norm(m) === 'SETIEMBRE'));
  if (!carpetaMes) return { ok: false, error: `No hay carpeta ${mes} en NC por NRF/${anio}`, hay: meses };
  const baseMes = `${baseAnio}/${carpetaMes}`;

  const indice = { periodo, origen: baseMes, generado: new Date().toISOString(), archivos: [] };
  const partes = [];
  for (const sub of (await carpetas(baseMes)).map((x) => x.Name)) {
    const drog = norm(sub);
    if (drog !== 'SUIZO' && drog !== 'MONROE') continue;
    for (const a of await archivos(`${baseMes}/${sub}`)) {
      const r = await fetch(`${sitio}/_api/web/GetFileByServerRelativePath(decodedurl='${q(`${baseMes}/${sub}/${a.Name}`)}')/$value`);
      if (!r.ok) throw new Error(`No pude bajar ${sub}/${a.Name}: ${r.status}`);
      const buf = await r.arrayBuffer();
      indice.archivos.push({ drogueria: drog, nombre: a.Name, bytes: buf.byteLength, modificado: a.TimeLastModified });
      partes.push(buf);
    }
  }
  const cab = new TextEncoder().encode(JSON.stringify(indice));
  const largo = new Uint8Array(4); new DataView(largo.buffer).setUint32(0, cab.length);
  const blob = new Blob([largo, cab, ...partes]);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = `SP-NCR-${periodo}.bundle`;
  document.body.appendChild(a); a.click(); a.remove();
  const cuenta = (d) => indice.archivos.filter((x) => x.drogueria === d).length;
  return { ok: true, archivo: a.download, bytes: blob.size, suizo: cuenta('SUIZO'), monroe: cuenta('MONROE') };
})();
