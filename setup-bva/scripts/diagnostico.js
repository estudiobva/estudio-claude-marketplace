#!/usr/bin/env node
/**
 * diagnostico.js — revisa que esta computadora pueda correr los plugins del
 * Estudio BVA. No entra a ARCA ni a ningun banco y nunca muestra claves.
 *
 *   node scripts/diagnostico.js          reporte legible
 *   node scripts/diagnostico.js --json   para parsear
 *
 * Chequea: Node, Git y acceso al repo privado, Python y sus paquetes, que esten
 * instalados todos los plugins del marketplace y con sus dependencias, que
 * Chromium abra de verdad, el .env, la unidad compartida, que
 * Claves_Organismos.xlsx y Claves_Bancos.xlsx abran, y la carpeta de salida.
 * Exit 1 si hay algun ERROR.
 */

const fs = require('fs');
const path = require('path');
const p = require('../lib/plugins');

const JSON_OUT = process.argv.includes('--json');
const checks = [];
const add = (id, estado, detalle, arreglo) => checks.push({ id, estado, detalle, ...(arreglo ? { arreglo } : {}) });
const INSTALADOR = 'Corre de nuevo "Instalar Claude BVA" (unidad compartida > 0 - ESTUDIO - 00000000000 > automatizacion_bva > setup).';

(async () => {
  // Como los plugins: el .env completa lo que no este en el entorno.
  const env = p.leerEnv();
  for (const [k, v] of Object.entries(env)) if (process.env[k] === undefined) process.env[k] = v;

  // Node
  const major = Number(process.versions.node.split('.')[0]);
  add('node', major >= 20 ? 'ok' : 'error', `v${process.versions.node}`, major >= 20 ? null : 'Hace falta Node 20 o superior. ' + INSTALADOR);

  // Git y el repo privado (lo que usa Claude para bajar y actualizar el marketplace)
  const git = p.correr('git', ['--version'], { timeout: 15000 });
  if (!git.ok) add('git', 'error', 'no instalado', INSTALADOR);
  else {
    const r = p.correr('git', ['ls-remote', `https://github.com/${p.REPO}`, 'HEAD'],
      // Sin ventanas de login: si no hay credenciales guardadas, que falle.
      { timeout: 30000, env: { GIT_TERMINAL_PROMPT: '0', GCM_INTERACTIVE: 'never' } });
    add('github', r.ok ? 'ok' : 'error', r.ok ? `acceso a ${p.REPO}` : `sin acceso a ${p.REPO}`,
      r.ok ? null : 'Tu usuario de GitHub tiene que estar invitado a la organizacion estudiobva y logueado (gh auth login + gh auth setup-git). ' + INSTALADOR);
  }

  // Python
  const py = p.python(env);
  const v = p.correr(py, ['-c', 'import sys;print(sys.executable, "%d.%d" % sys.version_info[:2])'], { timeout: 20000 });
  const pyOk = v.ok && !/WindowsApps/i.test(v.out);
  add('python', pyOk ? 'ok' : 'error', pyOk ? v.out : `no arranca: ${py}`, pyOk ? null : INSTALADOR);

  // Plugins instalados vs publicados
  const inst = p.instalados();
  const pub = p.publicados();
  if (!pub) add('marketplace', 'error', `no esta agregado el marketplace ${p.MARKET}`, INSTALADOR);
  else {
    const faltan = pub.filter(n => !inst.some(i => i.nombre === n));
    add('plugins', faltan.length ? 'aviso' : 'ok',
      `${inst.length} instalados${faltan.length ? ` · sin instalar: ${faltan.join(', ')}` : ''}`,
      faltan.length ? `claude plugin install <nombre>@${p.MARKET}, o: ${INSTALADOR}` : null);
  }

  // Dependencias de cada plugin
  let conPlaywright = null;
  for (const pl of inst) {
    const falta = [...p.nodeFaltantes(pl.dir).map(d => `npm:${d}`),
                   ...(pyOk ? p.modulosFaltantes(py, p.modulosRequeridos(pl.dir)).map(m => `python:${m}`) : [])];
    add(`plugin:${pl.nombre}`, falta.length ? 'error' : 'ok', `${pl.version}${falta.length ? ` · falta ${falta.join(', ')}` : ''}`,
      falta.length ? 'node scripts/dependencias.js (en la carpeta de setup-bva)' : null);
    if (!conPlaywright && !p.nodeFaltantes(pl.dir).length && p.dependenciasNode(pl.dir).includes('playwright')) conPlaywright = pl;
  }

  // Chromium: que abra de verdad
  if (conPlaywright) {
    const r = p.correr(process.execPath, ['-e',
      "require('playwright').chromium.launch({headless:true}).then(b=>b.close()).then(()=>process.exit(0),e=>{console.error(e.message.split('\\n')[0]);process.exit(1)})"],
      { cwd: conPlaywright.dir, timeout: 90000 });
    add('chromium', r.ok ? 'ok' : 'error', r.ok ? 'abre' : (r.err || 'no abre'), r.ok ? null : 'node scripts/dependencias.js');
  }

  // .env
  const file = p.envFile();
  add('env', fs.existsSync(file) ? 'ok' : 'error', fs.existsSync(file) ? file : `no existe ${file}`, fs.existsSync(file) ? null : INSTALADOR);

  // Unidad compartida
  const unidad = p.variable('BVA_UNIDAD_PATH', env);
  const unidadOk = unidad && fs.existsSync(path.join(unidad, '0 - ESTUDIO - 00000000000'));
  add('unidad', unidadOk ? 'ok' : 'error', unidad || 'BVA_UNIDAD_PATH sin definir',
    unidadOk ? null : 'Abri Google Drive para escritorio con la cuenta del estudio. ' + INSTALADOR);

  // Planillas de claves (con el mismo codigo de los plugins; solo se cuentan filas)
  const conClaves = inst.find(x => fs.existsSync(path.join(x.dir, 'lib', 'claves.js')) && !p.nodeFaltantes(x.dir).length);
  if (conClaves) {
    try {
      const s = require(path.join(conClaves.dir, 'lib', 'claves.js')).status();
      add('claves-organismos', s.ok ? 'ok' : 'error',
        s.ok ? `abre · ${s.titulares} titulares, ${s.conClaveArca} con clave de ARCA` : (s.error || 'no abre'),
        s.ok ? null : 'Revisa la contrasena de Claves_Organismos.xlsx. ' + INSTALADOR);
      if (s.ok && s.file) {
        const b = Buffer.alloc(4); const fd = fs.openSync(s.file, 'r'); fs.readSync(fd, b, 0, 4, 0); fs.closeSync(fd);
        if (b.readUInt32LE(0) !== 0xe011cfd0) {
          add('claves-cifrado', 'aviso', 'Claves_Organismos.xlsx no tiene contrasena de apertura',
            'Cualquiera con acceso a la unidad puede leer las claves: conviene protegerlo (Archivo > Informacion > Proteger libro > Cifrar con contrasena).');
        }
      }
    } catch (e) { add('claves-organismos', 'error', e.message.split('\n')[0], INSTALADOR); }
  }
  const bancos = inst.find(x => x.nombre === 'bancos-bva');
  if (bancos && !p.nodeFaltantes(bancos.dir).length) {
    try {
      const { bancos: filas } = await require(path.join(bancos.dir, 'lib', 'claves-bancos.js')).load();
      add('claves-bancos', 'ok', `abre · ${filas.length} bancos`);
    } catch (e) {
      add('claves-bancos', 'aviso', e.message.split('\n')[0], 'Solo lo usa bancos-bva. ' + INSTALADOR);
    }
  }

  // Salida
  const salidas = p.variable('BVA_SALIDAS_PATH', env) || path.join(require('os').homedir(), 'Documents', 'BVA-salidas');
  try { fs.mkdirSync(salidas, { recursive: true }); fs.accessSync(salidas, fs.constants.W_OK); add('salidas', 'ok', salidas); }
  catch (e) { add('salidas', 'error', `no puedo escribir en ${salidas}`, 'Defini BVA_SALIDAS_PATH en el .env.'); }

  // Actualizacion automatica
  const s = p.leerJson(path.join(p.claudeDir(), 'settings.json')) || {};
  const au = s.extraKnownMarketplaces && s.extraKnownMarketplaces[p.MARKET] && s.extraKnownMarketplaces[p.MARKET].autoUpdate === true;
  add('auto-update', au ? 'ok' : 'aviso', au ? 'el marketplace se actualiza solo al abrir Claude' : 'actualizacion automatica apagada',
    au ? null : 'node scripts/configurar.js');

  const ok = !checks.some(c => c.estado === 'error');
  if (JSON_OUT) {
    console.log(JSON.stringify({ ok, checks }, null, 2));
  } else {
    const et = { ok: 'OK   ', aviso: 'AVISO', error: 'ERROR' };
    for (const c of checks) {
      console.log(`  ${et[c.estado]} ${c.id.padEnd(30)} ${c.detalle}`);
      if (c.arreglo) console.log(`        ${' '.repeat(30)} -> ${c.arreglo}`);
    }
    console.log(ok ? '\nTodo en orden: los plugins estan listos para usar.' : '\nHay errores: arreglalos y volve a correr el diagnostico.');
  }
  process.exit(ok ? 0 : 1);
})();
