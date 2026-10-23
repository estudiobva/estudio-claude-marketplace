// prestamos/index.js — navegacion hasta los prestamos de cada banco, para
// scripts/cuadro-marcha.js. El login lo sigue haciendo lib/bancos/<banco>.js.
//
// Cada modulo exporta:
//   banco, sesion: 'unica' (un login, se cambia de empresa adentro) |
//                  'por_empresa' (salir y volver a entrar por empresa)
//   empresas(page)          -> nombres tal como los muestra el banco (string, o
//                              { nombre, cuit } si el banco muestra el CUIT)
//   entrar(page, nombre)
//   listar(page)            -> { texto, prestamos: [] | [{ id, descripcion, ... }] | null }
//                              null = hay algo que todavia no se sabe leer (relevar)
//   cuadro(page, prestamo, { guardar, aviso })  opcional: baja el cuadro de marcha;
//                              guardar(buffer, nombreSugerido, tipo) lo archiva (tipo:
//                              'cuadro' por defecto, o 'detalle' | 'cuotas' | 'contrato');
//                              aviso(texto) anota algo que no se pudo bajar; puede
//                              completar prestamo.id y prestamo.datos con lo leido
//   salir(page)

const MODULOS = {
  NACION: require('./nacion'),
  MACRO:  require('./macro'),
  PATAGONIA: require('./patagonia'),
  PROVINCIA: require('./provincia'),
};

module.exports = { MODULOS, bancos: Object.keys(MODULOS) };
