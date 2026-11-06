/**
 * Seguimiento de pedidos Estudio <-> Cliente — Apps Script de la Hoja de cálculo de Google.
 *
 * Va pegado en la planilla (Extensiones > Apps Script) y publicado como app web.
 * El script de la skill (scripts/pedidos.py) le habla por POST con un token:
 *
 *   {"token": "...", "accion": "estado",     "email": "...", "slack": "..."}
 *   {"token": "...", "accion": "aplicar",    "pedido": {persona, desde, hasta, ops: [...]}, "solo_revisar": true}
 *   {"token": "...", "accion": "pendientes", "persona": "...", "hoy": "AAAA-MM-DD"}
 *   {"token": "...", "accion": "importar",   "datos": {...}}       (migración, solo con la planilla vacía)
 *   {"token": "...", "accion": "probar"}
 *
 * Instalación (una vez): correr instalar() desde el editor, copiar el token del registro,
 * Implementar > Nueva implementación > App web (ejecutar como: yo; acceso: cualquier persona).
 */

var FILA_ENC = 2;
var FILA_1 = 3;
var ESTADOS = ['Pendiente', 'En curso', 'Respondido parcial', 'Completado', 'Cancelado', 'A confirmar'];
var CERRADOS = ['Completado', 'Cancelado'];
var CANALES = ['Mail', 'WhatsApp', 'Teléfono', 'SharePoint', 'Reunión', 'Slack'];
var TEMAS = ['EERR', 'Rectificativas', 'Calificaciones bancarias', 'Inspecciones', 'Liquidaciones',
  'Sueldos', 'Bancos', 'Gestión interna', 'Otros'];

// [clave, encabezado, ancho, tipo]. Tipos: fecha, texto, num, dias (fórmula), lista:<Lista>
var HOJAS = {
  EC: {
    titulo: 'Estudio → Cliente', prefijo: 'E', cierre: 'fecha_recibido',
    nota: 'Lo que el estudio le pide al cliente. Antes de pedir algo, filtrá por tema/sociedad: si ya está pedido, sumá un reclamo en vez de pedirlo de nuevo.',
    cols: [
      ['id', 'ID', 60, 'texto'], ['fecha_pedido', 'Fecha pedido', 90, 'fecha'],
      ['pedido_por', 'Pedido por (estudio)', 110, 'lista:Equipo'],
      ['responsable', 'Responsable seguimiento', 110, 'lista:Equipo'],
      ['a_quien', 'A quién (cliente)', 130, 'lista:Cliente'], ['sociedad', 'Sociedad', 110, 'texto'],
      ['tema', 'Tema', 110, 'lista:Tema'], ['que', 'Qué se pidió', 320, 'texto'],
      ['canal', 'Canal', 80, 'lista:Canal'], ['fecha_maxima', 'Fecha máxima', 90, 'fecha'],
      ['estado', 'Estado', 110, 'lista:Estado'], ['fecha_recibido', 'Fecha recibido', 90, 'fecha'],
      ['dias_abiertos', 'Días abiertos', 70, 'dias'], ['reclamos', 'Reclamos (cant.)', 70, 'num'],
      ['ultimo_reclamo', 'Último reclamo', 90, 'fecha'],
      ['historial', 'Historial de reclamos', 220, 'texto'], ['esfuerzo', 'Esfuerzo (días)', 70, 'num'],
      ['link', 'Link / respaldo', 180, 'texto'], ['notas', 'Notas', 280, 'texto'],
      ['clave_origen', 'Clave origen', 160, 'texto']
    ]
  },
  CE: {
    titulo: 'Cliente → Estudio', prefijo: 'C', cierre: 'fecha_entrega',
    nota: 'Lo que el cliente le pide al estudio. Esfuerzo en días: sirve para mostrarle al cliente todo lo que se hace fuera del abono.',
    cols: [
      ['id', 'ID', 60, 'texto'], ['fecha_pedido', 'Fecha pedido', 90, 'fecha'],
      ['pedido_por', 'Pedido por (cliente)', 130, 'lista:Cliente'], ['canal', 'Canal', 80, 'lista:Canal'],
      ['responsable', 'Responsable (estudio)', 110, 'lista:Equipo'], ['sociedad', 'Sociedad', 110, 'texto'],
      ['tema', 'Tema', 110, 'lista:Tema'], ['que', 'Qué pidió', 320, 'texto'],
      ['fecha_maxima', 'Fecha máxima', 90, 'fecha'], ['estado', 'Estado', 110, 'lista:Estado'],
      ['fecha_entrega', 'Fecha entrega', 90, 'fecha'], ['dias_abiertos', 'Días abiertos', 70, 'dias'],
      ['reclamos', 'Reclamos recibidos', 70, 'num'], ['ultimo_reclamo', 'Último reclamo', 90, 'fecha'],
      ['esfuerzo', 'Esfuerzo (días)', 70, 'num'], ['link', 'Link / respaldo', 180, 'texto'],
      ['notas', 'Notas', 280, 'texto'], ['clave_origen', 'Clave origen', 160, 'texto']
    ]
  },
  TI: {
    titulo: 'Tareas internas', prefijo: 'T', cierre: 'fecha_cierre',
    nota: 'Tareas que salen de reuniones, Slack o mails internos, con responsable. Lo que no tiene responsable y fecha, queda en el aire.',
    cols: [
      ['id', 'ID', 60, 'texto'], ['origen', 'Origen', 150, 'texto'], ['tarea', 'Tarea', 360, 'texto'],
      ['responsable', 'Responsable', 110, 'lista:Equipo'], ['pedido_por', 'Pedido por', 110, 'lista:Equipo'],
      ['fecha_pedido', 'Fecha pedido', 90, 'fecha'], ['fecha_maxima', 'Fecha compromiso', 90, 'fecha'],
      ['estado', 'Estado', 110, 'lista:Estado'], ['fecha_cierre', 'Fecha cierre', 90, 'fecha'],
      ['link', 'Link / respaldo', 180, 'texto'], ['notas', 'Notas', 300, 'texto'],
      ['clave_origen', 'Clave origen', 160, 'texto']
    ]
  }
};
var ORDEN = ['EC', 'CE', 'TI'];
var OBLIGATORIOS = { EC: ['que'], CE: ['que'], TI: ['tarea', 'responsable'] };
var TZ = 'America/Argentina/Buenos_Aires';

// ───────────────────────────── app web ─────────────────────────────

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return salida_({ ok: false, error: 'El pedido no es JSON válido.' });
  }
  var token = PropertiesService.getScriptProperties().getProperty('PEDIDOS_TOKEN');
  if (!token || req.token !== token) return salida_({ ok: false, error: 'Token inválido.' });
  try {
    return salida_(atender_(SpreadsheetApp.getActiveSpreadsheet(), req));
  } catch (err) {
    return salida_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return salida_({ ok: true, servicio: 'seguimiento-pedidos', mensaje: 'Usar POST con token.' });
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function atender_(ss, req) {
  switch (req.accion) {
    case 'probar': {
      var columnas = {};
      ORDEN.forEach(function (k) {
        columnas[k] = {};
        HOJAS[k].cols.forEach(function (c) { columnas[k][c[0]] = c[1]; });
      });
      return { ok: true, planilla: ss.getName(), url: ss.getUrl(), columnas: columnas };
    }
    case 'estado': return estado_(ss, req);
    case 'pendientes': return pendientes_(ss, req);
    case 'aplicar':
    case 'importar': {
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        return req.accion === 'aplicar' ? aplicar_(ss, req.pedido || {}, !!req.solo_revisar) : importar_(ss, req.datos || {});
      } finally {
        lock.releaseLock();
      }
    }
    default: throw new Error('Acción desconocida "' + req.accion + '" (estado, aplicar, pendientes, importar, probar).');
  }
}

// ───────────────────────────── utilidades ─────────────────────────────

function hoy_() {
  return aFecha_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
}

/** 'AAAA-MM-DD' | 'DD/MM/AAAA' | Date → Date a medianoche local (o null). */
function aFecha_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) return new Date(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
  throw new Error('Fecha no reconocida: "' + v + '" (usar AAAA-MM-DD)');
}

function iso_(v) {
  var d = aFecha_(v);
  if (!d) return null;
  var p = function (n) { return (n < 10 ? '0' : '') + n; };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

function vacio_(v) {
  return v === null || v === undefined || v === '';
}

function jsonVal_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' ? iso_(v) : v;
}

function hoja_(ss, nombre) {
  var sh = ss.getSheetByName(nombre);
  if (!sh) throw new Error('Falta la hoja "' + nombre + '". Correr instalar() desde el editor de Apps Script.');
  return sh;
}

/** Lee una hoja de datos: mapa clave → columna (1-based) y las filas con ID. */
function tabla_(ss, k) {
  var h = HOJAS[k];
  var sh = hoja_(ss, h.titulo);
  var valores = sh.getDataRange().getValues();
  var enc = valores[FILA_ENC - 1] || [];
  var m = {};
  h.cols.forEach(function (c) {
    var j = enc.indexOf(c[1]);
    if (j < 0) throw new Error('En la hoja "' + h.titulo + '" falta la columna "' + c[1] + '". No cambies los encabezados.');
    m[c[0]] = j + 1;
  });
  var filas = [];
  for (var r = FILA_1; r <= valores.length; r++) {
    var fila = valores[r - 1];
    if (vacio_(fila[m.id - 1])) continue;
    var d = { _fila: r };
    h.cols.forEach(function (c) {
      if (c[0] !== 'dias_abiertos') d[c[0]] = fila[m[c[0]] - 1];
    });
    filas.push(d);
  }
  return { k: k, sh: sh, m: m, filas: filas, total: valores.length };
}

function equipo_(ss) {
  var sh = ss.getSheetByName('Equipo');
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(1).filter(function (f) { return !vacio_(f[0]); }).map(function (f) {
    return {
      nombre: String(f[0]).trim(), email: String(f[1] || '').trim().toLowerCase(),
      slack: String(f[2] || '').trim(),
      activo: ['no', 'n', '0'].indexOf(String(f[3] || 'Sí').trim().toLowerCase()) < 0
    };
  });
}

function config_(ss) {
  var c = {};
  var sh = ss.getSheetByName('Config');
  if (sh) sh.getDataRange().getValues().slice(1).forEach(function (f) { if (!vacio_(f[0])) c[String(f[0]).trim()] = f[1]; });
  var lista = function (s) {
    return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
  };
  return {
    dominios_cliente: lista(c['Dominios del cliente']).map(function (d) { return d.replace(/^@/, '').toLowerCase(); }),
    dias_primera_corrida: parseInt(c['Días primera corrida'], 10) || 14,
    link: c['Link de la planilla'] || ss.getUrl(),
    canales_slack: lista(c['Canales de Slack']).map(function (x) { return x.replace(/^#/, ''); })
  };
}

function listaCol_(ss, titulo) {
  var sh = ss.getSheetByName('Listas');
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  var j = v[0].indexOf(titulo);
  if (j < 0) return [];
  return v.slice(1).map(function (f) { return f[j]; }).filter(function (x) { return !vacio_(x); });
}

function resolverPersona_(ss, q) {
  var eq = equipo_(ss);
  var nombres = eq.map(function (e) { return e.nombre; });
  if (q.persona) {
    var p = String(q.persona).trim().toLowerCase();
    for (var i = 0; i < eq.length; i++) if (eq[i].nombre.toLowerCase() === p) return eq[i];
    throw new Error('"' + q.persona + '" no está en la hoja Equipo. Equipo: ' + nombres.join(', '));
  }
  var email = String(q.email || '').toLowerCase();
  var slack = String(q.slack || '');
  for (var a = 0; a < eq.length; a++) if (email && eq[a].email === email) return eq[a];
  for (var b = 0; b < eq.length; b++) if (slack && eq[b].slack === slack) return eq[b];
  throw new Error('No sé quién sos: tu mail o Slack ID no está en la hoja Equipo. Agregate ahí (Nombre, Email, Slack ID) o pasá la persona. Equipo: ' + nombres.join(', '));
}

// ───────────────────────────── estado ─────────────────────────────

function estado_(ss, req) {
  var yo = resolverPersona_(ss, req);
  var cfg = config_(ss);
  var ultima = null;
  var co = ss.getSheetByName('Corridas');
  if (co) co.getDataRange().getValues().slice(1).forEach(function (f) {
    if (String(f[1] || '').trim().toLowerCase() === yo.nombre.toLowerCase()) {
      var h = vacio_(f[3]) ? null : aFecha_(f[3]);
      if (h && (!ultima || h > ultima)) ultima = h;
    }
  });
  var hoy = hoy_();
  var desde = ultima || new Date(hoy.getTime() - cfg.dias_primera_corrida * 86400000);
  var claves = {}, abiertos = [];
  ORDEN.forEach(function (k) {
    tabla_(ss, k).filas.forEach(function (f) {
      String(f.clave_origen || '').split('|').forEach(function (c) { if (c) claves[c] = true; });
      if (CERRADOS.indexOf(f.estado) < 0) {
        var o = { hoja: k };
        Object.keys(f).forEach(function (c) { if (c !== '_fila' && !vacio_(f[c])) o[c] = jsonVal_(f[c]); });
        abiertos.push(o);
      }
    });
  });
  return {
    ok: true, planilla: ss.getName(), url: ss.getUrl(), persona: yo.nombre, email: yo.email, slack: yo.slack,
    desde: iso_(desde), hasta: iso_(hoy), primera_corrida: !ultima, config: cfg,
    equipo: equipo_(ss).filter(function (e) { return e.activo; }),
    cliente: listaCol_(ss, 'Cliente'), temas: listaCol_(ss, 'Tema'), canales: listaCol_(ss, 'Canal'),
    estados: listaCol_(ss, 'Estado'), claves_cargadas: Object.keys(claves).sort(), abiertos: abiertos
  };
}

// ───────────────────────────── aplicar ─────────────────────────────

function siguienteId_(filas, prefijo) {
  var n = 0;
  filas.forEach(function (f) {
    var m = String(f.id).match(new RegExp('^' + prefijo + '-(\\d+)$'));
    if (m) n = Math.max(n, +m[1]);
  });
  var s = String(n + 1);
  while (s.length < 3) s = '0' + s;
  return prefijo + '-' + s;
}

function filaLibre_(t, usadas) {
  var ocupadas = {};
  t.filas.forEach(function (f) { ocupadas[f._fila] = true; });
  Object.keys(usadas).forEach(function (r) { ocupadas[r] = true; });
  var r = FILA_1;
  while (ocupadas[r]) r++;
  return r;
}

/** Escribe datos en una fila. pisar=false: no pisa texto que ya hay (salvo estado y fechas; notas/link se agregan). */
function escribir_(t, r, datos, pisar, actual) {
  var tipos = {};
  HOJAS[t.k].cols.forEach(function (c) { tipos[c[0]] = c[3]; });
  var hechos = {};
  Object.keys(datos).forEach(function (c) {
    if (!(c in t.m) || c === 'id' || c === 'dias_abiertos') return;
    var v = datos[c];
    if (vacio_(v)) return;
    var previo = actual ? actual[c] : null;
    if (!pisar && !vacio_(previo) && c !== 'estado') {
      if (c === 'notas' || c === 'link') {
        if (String(previo).indexOf(String(v)) >= 0) return;
        v = previo + ' | ' + v;
      } else if (!(tipos[c] === 'fecha' && c !== 'fecha_pedido')) {
        return;
      }
    }
    if (tipos[c] === 'fecha') v = aFecha_(v);
    else if (tipos[c] === 'num') v = Number(v);
    else if (tipos[c] === 'lista:Estado' && ESTADOS.indexOf(v) < 0) {
      throw new Error('Estado inválido "' + v + '" (válidos: ' + ESTADOS.join(', ') + ')');
    }
    if (!vacio_(previo) && String(jsonVal_(previo)) === String(jsonVal_(v))) return;  // sin cambios
    t.escrituras.push([r, t.m[c], v]);
    if (actual) actual[c] = v;
    hechos[c] = jsonVal_(v);
  });
  return hechos;
}

function aplicar_(ss, pedido, soloRevisar) {
  var persona = pedido.persona ? resolverPersona_(ss, { persona: pedido.persona }).nombre : null;
  var hoy = hoy_();
  var tablas = {}, claves = {}, porId = {}, usadas = {};
  ORDEN.forEach(function (k) {
    var t = tabla_(ss, k);
    t.escrituras = [];
    tablas[k] = t;
    usadas[k] = {};
    t.filas.forEach(function (f) {
      String(f.clave_origen || '').split('|').forEach(function (c) { if (c) claves[c] = f.id; });
      porId[String(f.id)] = { k: k, f: f };
    });
  });
  var res = { agregados: [], actualizados: [], reclamos: [], omitidos: [], errores: [] };

  (pedido.ops || []).forEach(function (op, i) {
    var n = i + 1;
    try {
      if (op.accion === 'agregar') {
        var k = HOJAS[op.hoja] ? op.hoja : null;
        if (!k) ORDEN.forEach(function (x) { if (HOJAS[x].titulo === op.hoja) k = x; });
        if (!k) throw new Error('hoja inválida "' + op.hoja + '" (EC, CE o TI)');
        var datos = JSON.parse(JSON.stringify(op.datos || {}));
        var clave = String(op.clave_origen || datos.clave_origen || '').trim();
        if (!clave) throw new Error('falta clave_origen (ej.: gmail:<threadId>#1 o slack:<canal>/<ts>)');
        if (claves[clave]) { res.omitidos.push({ op: n, motivo: 'ya cargado como ' + claves[clave], clave: clave }); return; }
        OBLIGATORIOS[k].forEach(function (ob) { if (vacio_(datos[ob])) throw new Error('falta "' + ob + '"'); });
        var t = tablas[k];
        if (vacio_(datos.estado)) datos.estado = 'Pendiente';
        if (vacio_(datos.fecha_pedido)) datos.fecha_pedido = iso_(hoy);
        if (k !== 'CE' && vacio_(datos.pedido_por) && persona) datos.pedido_por = persona;
        if (vacio_(datos.responsable) && persona) datos.responsable = persona;
        if (k !== 'TI' && vacio_(datos.reclamos)) datos.reclamos = 0;
        datos.clave_origen = clave;
        var nuevo = siguienteId_(t.filas, HOJAS[k].prefijo);
        var r = filaLibre_(t, usadas[k]);
        usadas[k][r] = true;
        t.escrituras.push([r, t.m.id, nuevo]);
        var fila = { id: nuevo, _fila: r };
        var hechos = escribir_(t, r, datos, true, fila);
        t.filas.push(fila);
        porId[nuevo] = { k: k, f: fila };
        claves[clave] = nuevo;
        res.agregados.push({ id: nuevo, hoja: HOJAS[k].titulo, que: hechos.que || hechos.tarea,
          responsable: hechos.responsable, estado: hechos.estado });
      } else if (op.accion === 'actualizar' || op.accion === 'reclamo') {
        var ident = String(op.id || '');
        if (!porId[ident]) throw new Error('no existe el ID "' + ident + '"');
        var ref = porId[ident], tt = tablas[ref.k], f = ref.f, rr = f._fila;
        var prev = String(f.clave_origen || '');
        var cl = String(op.clave_origen || '').trim();
        if (cl && prev.split('|').indexOf(cl) >= 0) {
          res.omitidos.push({ op: n, motivo: 'esa novedad ya estaba aplicada', id: ident });
          return;
        }
        if (op.accion === 'actualizar') {
          var h = escribir_(tt, rr, op.datos || {}, !!op.pisar, f);
          if (Object.keys(h).length) res.actualizados.push({ id: ident, cambios: h });
          else res.omitidos.push({ op: n, motivo: 'sin cambios', id: ident });
        } else {
          if (ref.k === 'TI') throw new Error('las tareas internas no llevan reclamos: usar "actualizar" con notas');
          var fecha = aFecha_(op.fecha || iso_(hoy));
          var cant = (Number(f.reclamos) || 0) + 1;
          tt.escrituras.push([rr, tt.m.reclamos, cant]);
          tt.escrituras.push([rr, tt.m.ultimo_reclamo, fecha]);
          f.reclamos = cant;
          f.ultimo_reclamo = fecha;
          var linea = Utilities.formatDate(fecha, TZ, 'dd/MM') + ': ' + (op.detalle || 'reclamo');
          var destino = tt.m.historial ? 'historial' : 'notas';
          var nuevoTexto = vacio_(f[destino]) ? linea : f[destino] + ' | ' + linea;
          tt.escrituras.push([rr, tt.m[destino], nuevoTexto]);
          f[destino] = nuevoTexto;
          res.reclamos.push({ id: ident, reclamos: cant, fecha: iso_(fecha) });
        }
        if (cl) {
          f.clave_origen = prev ? prev + '|' + cl : cl;
          tt.escrituras.push([rr, tt.m.clave_origen, f.clave_origen]);
          claves[cl] = ident;
        }
      } else {
        throw new Error('acción inválida "' + op.accion + '" (agregar, actualizar o reclamo)');
      }
    } catch (err) {
      res.errores.push({ op: n, error: String(err && err.message || err), detalle: op });
    }
  });

  if (soloRevisar) return Object.assign({ ok: true, solo_revisar: true, planilla: ss.getName() }, res);

  ORDEN.forEach(function (k) {
    var t = tablas[k];
    t.escrituras.forEach(function (w) { t.sh.getRange(w[0], w[1]).setValue(w[2]); });
  });
  var co = ss.getSheetByName('Corridas');
  if (co && persona) {
    co.appendRow([new Date(), persona, pedido.desde ? aFecha_(pedido.desde) : '', aFecha_(pedido.hasta || iso_(hoy)),
      res.agregados.length, res.actualizados.length, res.reclamos.length, res.omitidos.length]);
  }
  return Object.assign({ ok: true, planilla: ss.getName(), url: ss.getUrl() }, res);
}

// ───────────────────────────── pendientes ─────────────────────────────

function pendientes_(ss, req) {
  var yo = resolverPersona_(ss, req).nombre;
  var hoy = req.hoy ? aFecha_(req.hoy) : hoy_();
  var en2 = new Date(hoy.getTime() + 2 * 86400000);
  var items = [];
  ORDEN.forEach(function (k) {
    tabla_(ss, k).filas.forEach(function (f) {
      if (CERRADOS.indexOf(f.estado) >= 0) return;
      var resp = String(f.responsable || '').trim().toLowerCase();
      if (resp !== yo.toLowerCase() && !(k === 'TI' && resp === 'el grupo')) return;
      var fp = vacio_(f.fecha_pedido) ? null : aFecha_(f.fecha_pedido);
      var fm = vacio_(f.fecha_maxima) ? null : aFecha_(f.fecha_maxima);
      var porQuien = String(f.pedido_por || '');
      items.push({
        id: f.id, hoja: k, estado: f.estado || 'Pendiente', texto: f.que || f.tarea,
        con: f.a_quien || (porQuien.toLowerCase() === yo.toLowerCase() ? '' : porQuien),
        sociedad: f.sociedad || '', fecha_pedido: iso_(fp), fecha_maxima: iso_(fm),
        dias: fp ? Math.round((hoy - fp) / 86400000) : null, reclamos: Number(f.reclamos) || 0,
        vencido: !!(fm && fm < hoy), vence_pronto: !!(fm && fm >= hoy && fm <= en2)
      });
    });
  });
  items.sort(function (a, b) {
    if (a.vencido !== b.vencido) return a.vencido ? -1 : 1;
    var x = (a.fecha_maxima || '9999') + (a.fecha_pedido || '9999');
    var y = (b.fecha_maxima || '9999') + (b.fecha_pedido || '9999');
    return x < y ? -1 : x > y ? 1 : 0;
  });
  return { ok: true, persona: yo, hoy: iso_(hoy), total: items.length,
    vencidos: items.filter(function (i) { return i.vencido; }).length, link: config_(ss).link, items: items };
}

// ───────────────────────────── importar (migración) ─────────────────────────────

/** Carga equipo, listas, config, filas y corridas de la planilla anterior. Solo con las hojas de datos vacías. */
function importar_(ss, d) {
  var res = { ok: true, equipo: 0, cliente: 0, config: 0, corridas: 0 };
  ORDEN.forEach(function (k) {
    var t = tabla_(ss, k);
    if (t.filas.length && (d.filas || {})[k] && d.filas[k].length) {
      throw new Error('La hoja "' + HOJAS[k].titulo + '" ya tiene datos: la importación es solo para una planilla vacía.');
    }
  });
  if (d.equipo && d.equipo.length) {
    var eq = hoja_(ss, 'Equipo');
    eq.getRange(2, 1, d.equipo.length, 4).setValues(d.equipo.map(function (e) {
      return [e.nombre || '', e.email || '', e.slack || '', e.activo === false ? 'No' : 'Sí'];
    }));
    res.equipo = d.equipo.length;
  }
  if (d.cliente && d.cliente.length) {
    var ls = hoja_(ss, 'Listas');
    var j = ls.getRange(1, 1, 1, ls.getLastColumn()).getValues()[0].indexOf('Cliente') + 1;
    ls.getRange(2, j, d.cliente.length, 1).setValues(d.cliente.map(function (c) { return [c]; }));
    res.cliente = d.cliente.length;
  }
  if (d.config) {
    var cf = hoja_(ss, 'Config');
    var v = cf.getDataRange().getValues();
    for (var r = 1; r < v.length; r++) {
      var p = String(v[r][0] || '').trim();
      if (p && !vacio_(d.config[p])) { cf.getRange(r + 1, 2).setValue(d.config[p]); res.config++; }
    }
  }
  ORDEN.forEach(function (k) {
    var filas = (d.filas || {})[k] || [];
    if (!filas.length) return;
    var t = tabla_(ss, k);
    t.escrituras = [];
    filas.forEach(function (f, i) {
      var r = FILA_1 + i;
      t.escrituras.push([r, t.m.id, f.id]);
      escribir_(t, r, f, true, null);
    });
    t.escrituras.forEach(function (w) { t.sh.getRange(w[0], w[1]).setValue(w[2]); });
    res[k] = filas.length;
  });
  if (d.corridas && d.corridas.length) {
    var co = hoja_(ss, 'Corridas');
    d.corridas.forEach(function (c) {
      co.appendRow([c[0] || '', c[1] || '', vacio_(c[2]) ? '' : aFecha_(c[2]), vacio_(c[3]) ? '' : aFecha_(c[3]),
        c[4] || 0, c[5] || 0, c[6] || 0, c[7] || 0]);
    });
    res.corridas = d.corridas.length;
  }
  return res;
}

// ───────────────────────────── instalación ─────────────────────────────

/** Correr UNA vez desde el editor: arma las hojas que falten y genera el token de la app web. */
function instalar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  configurar_(ss);
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('PEDIDOS_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('PEDIDOS_TOKEN', token);
  }
  Logger.log('Listo. Token para el .env (PEDIDOS_TOKEN): ' + token);
  Logger.log('Ahora: Implementar > Nueva implementación > App web, ejecutar como "Yo", acceso "Cualquier persona", y copiar la URL a PEDIDOS_URL.');
}

/** Genera un token nuevo (el anterior deja de servir). */
function regenerarToken() {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('PEDIDOS_TOKEN', token);
  Logger.log('Token nuevo (PEDIDOS_TOKEN): ' + token);
}

function configurar_(ss) {
  ss.setSpreadsheetTimeZone(TZ);
  var AZUL = '#1F3864';
  var gris = '#595959';
  var obtener = function (nombre) { return ss.getSheetByName(nombre) || ss.insertSheet(nombre); };
  var encabezar = function (sh, fila, textos) {
    sh.getRange(fila, 1, 1, textos.length).setValues([textos])
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(AZUL);
  };

  // Listas, Equipo, Config, Corridas (no se pisan si ya tienen datos)
  var ls = obtener('Listas');
  if (ls.getLastRow() < 2) {
    encabezar(ls, 1, ['Estado', 'Canal', 'Tema', 'Cliente']);
    [ESTADOS, CANALES, TEMAS].forEach(function (vals, j) {
      ls.getRange(2, j + 1, vals.length, 1).setValues(vals.map(function (v) { return [v]; }));
    });
    ls.getRange(1, 6).setValue('La columna Cliente se completa con las personas del cliente (nombre y apellido).')
      .setFontStyle('italic').setFontColor(gris);
  }
  var eq = obtener('Equipo');
  if (eq.getLastRow() < 1) {
    encabezar(eq, 1, ['Nombre', 'Email', 'Slack ID', 'Activo']);
    eq.getRange(1, 6).setValue('Una fila por persona del estudio. El Email y el Slack ID sirven para saber quién corre la skill.')
      .setFontStyle('italic').setFontColor(gris);
  }
  var cf = obtener('Config');
  if (cf.getLastRow() < 1) {
    encabezar(cf, 1, ['Parámetro', 'Valor', 'Para qué']);
    cf.getRange(2, 1, 4, 3).setValues([
      ['Dominios del cliente', '', 'Dominios de mail del cliente separados por coma. Solo esos mails se leen como pedidos del cliente.'],
      ['Días primera corrida', 14, 'Cuántos días hacia atrás se leen la primera vez que alguien corre la skill.'],
      ['Link de la planilla', ss.getUrl(), 'Link de esta planilla, para el mensaje de pendientes.'],
      ['Canales de Slack', '', 'Canales a leer además de privados y menciones (nombres separados por coma, sin #).']
    ]);
  }
  var co = obtener('Corridas');
  if (co.getLastRow() < 1) {
    encabezar(co, 1, ['Fecha y hora', 'Persona', 'Desde', 'Hasta', 'Agregados', 'Actualizados', 'Reclamos', 'Omitidos']);
    co.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
    co.getRange('C:D').setNumberFormat('dd/mm/yyyy');
  }

  var rangos = {
    Estado: ls.getRange('A2:A30'), Canal: ls.getRange('B2:B30'), Tema: ls.getRange('C2:C40'),
    Cliente: ls.getRange('D2:D80'), Equipo: eq.getRange('A2:A40')
  };

  // Hojas de datos
  ORDEN.forEach(function (k) {
    var h = HOJAS[k];
    var sh = obtener(h.titulo);
    var n = h.cols.length;
    if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
    if (sh.getMaxRows() < 1000) sh.insertRowsAfter(sh.getMaxRows(), 1000 - sh.getMaxRows());
    sh.getRange(1, 1).setValue(h.nota).setFontStyle('italic').setFontColor(gris);
    encabezar(sh, FILA_ENC, h.cols.map(function (c) { return c[1]; }));
    sh.getRange(FILA_ENC, 1, 1, n).setWrap(true).setVerticalAlignment('middle');
    var idx = {};
    h.cols.forEach(function (c, j) { idx[c[0]] = j + 1; });
    var ultima = sh.getMaxRows();
    var alto = ultima - FILA_1 + 1;
    var letra = function (j) { return sh.getRange(1, j).getA1Notation().replace(/\d+/g, ''); };
    h.cols.forEach(function (c, j0) {
      var j = j0 + 1;
      sh.setColumnWidth(j, c[2]);
      var col = sh.getRange(FILA_1, j, alto, 1);
      if (c[3] === 'fecha') col.setNumberFormat('dd/mm/yyyy');
      if (c[3] === 'texto') col.setWrap(true);
      if (c[3] === 'dias') {
        var F = letra(idx.fecha_pedido), C = letra(idx[h.cierre]);
        sh.getRange(FILA_1, j).setFormula('=ARRAYFORMULA(IF(' + F + FILA_1 + ':' + F + '="","",IF(' + C + FILA_1 + ':' + C +
          '<>"",' + C + FILA_1 + ':' + C + '-' + F + FILA_1 + ':' + F + ',TODAY()-' + F + FILA_1 + ':' + F + ')))');
        col.setBackground('#F2F2F2').setNumberFormat('0');
      }
      if (c[3].indexOf('lista:') === 0) {
        col.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInRange(rangos[c[3].slice(6)], true).setAllowInvalid(true).build());
      }
    });
    var E = letra(idx.estado), M = letra(idx.fecha_maxima);
    var todo = sh.getRange(FILA_1, 1, alto, n);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR($' + E + FILA_1 + '="Completado",$' + E + FILA_1 + '="Cancelado")')
        .setBackground('#E2EFDA').setRanges([todo]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($' + M + FILA_1 + '<>"",$' + M + FILA_1 + '<TODAY())')
        .setBackground('#F8CBAD').setRanges([todo]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + E + FILA_1 + '="A confirmar"')
        .setBackground('#FFF2CC').setRanges([todo]).build()
    ]);
    sh.setFrozenRows(FILA_ENC);
    sh.setFrozenColumns(2);
    if (!sh.getFilter()) sh.getRange(FILA_ENC, 1, alto + 1, n).createFilter();
    sh.hideColumns(idx.clave_origen);
  });

  // Resumen
  var rs = obtener('Resumen');
  rs.clear();
  rs.getRange('A1').setValue('Seguimiento de pedidos Estudio ↔ Cliente — resumen').setFontWeight('bold').setFontSize(14);
  rs.getRange('A2').setValue('Se actualiza solo. Rojo = pasó la fecha máxima y no está cerrado.').setFontStyle('italic').setFontColor(gris);
  var rng = function (k, clave) {
    var j = HOJAS[k].cols.map(function (c) { return c[0]; }).indexOf(clave) + 1;
    var L = hoja_(ss, HOJAS[k].titulo).getRange(1, j).getA1Notation().replace(/\d+/g, '');
    return "'" + HOJAS[k].titulo + "'!" + L + FILA_1 + ':' + L;
  };
  encabezar(rs, 4, ['Estado'].concat(ORDEN.map(function (k) { return HOJAS[k].titulo; })));
  ESTADOS.forEach(function (est, i) {
    var r = 5 + i;
    rs.getRange(r, 1).setValue(est);
    ORDEN.forEach(function (k, c) { rs.getRange(r, c + 2).setFormula('=COUNTIF(' + rng(k, 'estado') + ',A' + r + ')'); });
  });
  var f = 5 + ESTADOS.length;
  rs.getRange(f, 1).setValue('Vencidos (abiertos)').setFontWeight('bold').setFontColor('#C00000');
  ORDEN.forEach(function (k, c) {
    var M2 = rng(k, 'fecha_maxima'), E2 = rng(k, 'estado');
    rs.getRange(f, c + 2).setFormula('=SUMPRODUCT((' + M2 + '<>"")*(' + M2 + '<TODAY())*(' + E2 + '<>"Completado")*(' + E2 + '<>"Cancelado"))');
  });
  rs.getRange(f + 1, 1).setValue('Esfuerzo total (días)').setFontWeight('bold');
  rs.getRange(f + 1, 2).setFormula('=SUM(' + rng('EC', 'esfuerzo') + ')');
  rs.getRange(f + 1, 3).setFormula('=SUM(' + rng('CE', 'esfuerzo') + ')');
  f += 3;
  encabezar(rs, f, ['Abiertos por responsable'].concat(ORDEN.map(function (k) { return HOJAS[k].titulo; })));
  for (var i = 1; i <= 20; i++) {
    var r2 = f + i;
    rs.getRange(r2, 1).setFormula('=IF(Equipo!A' + (i + 1) + '="","",Equipo!A' + (i + 1) + ')');
    ORDEN.forEach(function (k, c) {
      var R = rng(k, 'responsable'), E3 = rng(k, 'estado');
      rs.getRange(r2, c + 2).setFormula('=IF($A' + r2 + '="","",SUMPRODUCT((' + R + '=$A' + r2 + ')*(' + E3 +
        '<>"Completado")*(' + E3 + '<>"Cancelado")*(' + E3 + '<>"")))');
    });
  }
  rs.setColumnWidth(1, 200);
  [2, 3, 4].forEach(function (j) { rs.setColumnWidth(j, 140); });

  // Cómo usarla
  var ay = obtener('Cómo usarla');
  ay.clear();
  var texto = [
    'Cómo usar esta planilla', '',
    '1. Todo pedido al cliente o del cliente se anota acá el mismo día, con quién lo pidió y quién se hace cargo.',
    '2. Antes de pedirle algo al cliente, filtrar por Tema o Sociedad: si ya está pedido, no se pide de nuevo; se suma un reclamo.',
    '3. Cada reclamo: sumar 1 en «Reclamos», poner la fecha en «Último reclamo» y anotarlo en el historial.',
    '4. «Fecha máxima» la ponen las dos partes. Si pasa y no está cerrado, la fila se pone en rojo.',
    '5. «Esfuerzo (días)»: estimar cuánto llevó. Sirve para mostrarle al cliente el trabajo fuera del abono.',
    '6. Al cerrar, poner Estado = Completado y la fecha. La fila se pone en verde.', '',
    'La skill «seguimiento-pedidos» lee tus mails con el cliente y tu Slack, carga acá lo nuevo y te manda tus pendientes.',
    'La columna oculta «Clave origen» es la que usa la skill para no cargar dos veces el mismo mail: no la borres.',
    'No cambiar los encabezados de la fila 2: la skill busca las columnas por nombre.'
  ];
  ay.getRange(1, 1, texto.length, 1).setValues(texto.map(function (t) { return [t]; }));
  ay.getRange('A1').setFontWeight('bold').setFontSize(13);
  ay.setColumnWidth(1, 800);

  // Orden de las hojas y limpieza de la hoja vacía inicial
  ['Resumen', HOJAS.EC.titulo, HOJAS.CE.titulo, HOJAS.TI.titulo, 'Cómo usarla', 'Listas', 'Equipo', 'Config', 'Corridas']
    .forEach(function (nombre, i) {
      var sh = ss.getSheetByName(nombre);
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    });
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (nombre) {
    var sh = ss.getSheetByName(nombre);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(ss.getSheetByName('Resumen'));
}
