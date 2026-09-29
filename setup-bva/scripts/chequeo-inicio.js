#!/usr/bin/env node
// chequeo-inicio.js — hook SessionStart del plugin setup-bva.
//
// Cada vez que se abre Claude mira, rapido y sin instalar nada, si a algun plugin
// del estudio le faltan dependencias (pasa despues de cada actualizacion: el
// plugin queda en una carpeta nueva sin node_modules), si falta el .env o si el
// marketplace publico un plugin que aca no esta instalado. Si todo esta bien no
// dice nada; si no, le deja a Claude la indicacion de que hacer.
// Nunca falla: un error aca no puede trabar el arranque.

const path = require('path');

try {
  const fs = require('fs');
  const p = require('../lib/plugins');
  const root = path.join(__dirname, '..');
  const inst = p.instalados();
  const avisos = [];

  const sinDeps = inst.filter(x => p.nodeFaltantes(x.dir).length).map(x => x.nombre);
  if (sinDeps.length) {
    avisos.push(`A estos plugins del estudio les faltan dependencias (se acaban de instalar o actualizar): ${sinDeps.join(', ')}. ` +
      `Antes de correr cualquier script de ellos, corre: node "${path.join(root, 'scripts', 'dependencias.js')}" ` +
      '(tarda un par de minutos; avisale al usuario que se estan instalando).');
  }

  if (!fs.existsSync(p.envFile())) {
    avisos.push(`No existe ${p.envFile()}: los scripts no van a encontrar la planilla de claves ni la unidad. ` +
      'Pedile al usuario que corra "Instalar Claude BVA" (unidad compartida > 0 - ESTUDIO - 00000000000 > automatizacion_bva > setup); ' +
      'la contrasena de la planilla la escribe el en esa ventana, nunca en el chat.');
  }

  const pub = p.publicados();
  const nuevos = pub ? pub.filter(n => !inst.some(i => i.nombre === n)) : [];
  if (nuevos.length) {
    avisos.push(`El marketplace ${p.MARKET} tiene plugins que no estan instalados aca: ${nuevos.join(', ')}. ` +
      `Ofrecele al usuario instalarlos (skill setup-bva).`);
  }

  if (avisos.length) {
    process.stdout.write(JSON.stringify({
      hookSpecificOutput: { hookEventName: 'SessionStart', additionalContext: `[setup-bva] ${avisos.join(' ')}` },
    }));
  }
} catch { /* silencio: nunca trabar el arranque */ }
process.exit(0);
