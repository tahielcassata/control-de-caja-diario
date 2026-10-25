/**
 * Pruebas del backend. Ejecutar con:  node test/test.js
 */
const assert = require('assert');
const { crearEntorno } = require('./simulador');

let pasadas = 0;
function prueba(nombre, fn) {
  try { fn(); pasadas++; console.log('  ✓ ' + nombre); }
  catch (e) { console.error('  ✗ ' + nombre + '\n    ' + e.message); process.exitCode = 1; }
}
function lanza(fn, texto) {
  assert.throws(fn, e => e.message.includes(texto), 'Se esperaba el error: ' + texto);
}
const cobro = (o = {}) => Object.assign({ puesto: 1, servicio: 'Corte', monto: 8000, medio: 'Efectivo', propina: 0, medioPropina: '' }, o);

/** Instalación típica: setup, PIN asignado y servicios cargados desde Administración. */
function preparar(gas) {
  gas.setup();
  gas.asignarPin_('1234');
  [['Corte', 8000], ['Barba', 5000], ['Corte y barba', 12000], ['Color', 15000]].forEach(([nombre, precio]) =>
    gas.adminGuardarServicio('1234', { original: '', nombre, precio, activo: true }));
}

console.log('\nSistema de cobros — pruebas\n');

prueba('setup crea hojas, puestos, servicios y activador', () => {
  const { gas, hojas, triggers } = crearEntorno();
  gas.setup();
  ['Cobros', 'Cierres', 'Puestos', 'Servicios'].forEach(h => assert(hojas.has(h), 'falta ' + h));
  const e = gas.getEstado();
  assert.deepStrictEqual(e.puestos, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
  assert.strictEqual(e.servicios.length, 0, 'arranca sin servicios');
  assert.strictEqual(triggers.length, 1);
  assert.strictEqual(e.hoy, '2026-10-21');
  gas.setup(); // repetir no duplica
  assert.strictEqual(gas.getEstado().puestos.length, 10);
  assert.strictEqual(triggers.length, 1);
});

prueba('guardar cobros y calcular totales con propinas en distinto medio', () => {
  const { gas } = crearEntorno(); preparar(gas);
  gas.guardarCobro(cobro());
  gas.guardarCobro(cobro({ puesto: 2, servicio: 'Corte y barba', monto: 12000, medio: 'Transferencia', propina: 1000, medioPropina: 'Efectivo' }));
  const e = gas.guardarCobro(cobro({ puesto: 1, servicio: 'Barba', monto: 5000, medio: 'Transferencia', propina: 2000, medioPropina: 'Transferencia' }));
  const r = e.resumen;
  assert.strictEqual(r.atenciones, 3);
  assert.strictEqual(r.servicios, 25000);
  assert.strictEqual(r.propinas, 3000);
  assert.strictEqual(r.efectivo, 9000);         // 8000 + propina 1000
  assert.strictEqual(r.transferencias, 19000);  // 12000 + 5000 + 2000
  assert.strictEqual(r.total, 28000);
  const p1 = r.porPuesto.find(p => p.puesto === 1);
  assert.strictEqual(p1.atenciones, 2); assert.strictEqual(p1.servicios, 13000); assert.strictEqual(p1.propinas, 2000);
  assert.strictEqual(e.ultimos[0].servicio, 'Barba');
});

prueba('validaciones del cobro', () => {
  const { gas } = crearEntorno(); preparar(gas);
  lanza(() => gas.guardarCobro(cobro({ puesto: 99 })), 'puesto');
  lanza(() => gas.guardarCobro(cobro({ servicio: 'Inventado' })), 'servicio');
  lanza(() => gas.guardarCobro(cobro({ monto: 0 })), 'monto');
  lanza(() => gas.guardarCobro(cobro({ medio: 'Cheque' })), 'medio de pago');
  lanza(() => gas.guardarCobro(cobro({ propina: 500 })), 'medio de la propina');
  assert.strictEqual(gas.getEstado().resumen.atenciones, 0);
});

prueba('deshacer borra solo el último cobro de hoy', () => {
  const { gas } = crearEntorno(); preparar(gas);
  lanza(() => gas.deshacerUltimo(), 'No hay cobros');
  gas.guardarCobro(cobro());
  gas.guardarCobro(cobro({ puesto: 3, monto: 9000 }));
  const e = gas.deshacerUltimo();
  assert.strictEqual(e.borrado.puesto, 3);
  assert.strictEqual(e.resumen.atenciones, 1);
  assert.strictEqual(e.resumen.total, 8000);
});

prueba('cerrar el día guarda el resumen y bloquea cargas y deshacer', () => {
  const { gas, hojas } = crearEntorno(); preparar(gas);
  gas.guardarCobro(cobro());
  gas.guardarCobro(cobro({ puesto: 2, propina: 500, medioPropina: 'Efectivo' }));
  const c = gas.cerrarDia();
  assert.strictEqual(c.cerrado, true);
  assert.strictEqual(c.resumen.total, 16500);
  const filas = hojas.get('Cierres').datos.slice(1);
  assert.strictEqual(filas.length, 3); // puesto 1, puesto 2 y TOTAL
  assert.strictEqual(filas[2][1], 'TOTAL');
  lanza(() => gas.guardarCobro(cobro()), 'cerrado');
  lanza(() => gas.deshacerUltimo(), 'cerrado');
  assert.strictEqual(gas.getEstado().cerrado, true);
  // el cierre guardado se lee igual
  assert.strictEqual(gas.getCierre('2026-10-21').resumen.propinas, 500);
});

prueba('reabrir con PIN permite cargar un cobro tarde y recalcula al cerrar', () => {
  const { gas, hojas } = crearEntorno(); preparar(gas);
  gas.guardarCobro(cobro());
  gas.cerrarDia();
  lanza(() => gas.reabrirDia('0000', '2026-10-21'), 'PIN incorrecto');
  gas.reabrirDia('1234', '2026-10-21');
  gas.guardarCobro(cobro({ puesto: 4, monto: 15000, servicio: 'Color' }));
  const c = gas.cerrarDia();
  assert.strictEqual(c.resumen.total, 23000);
  assert.strictEqual(hojas.get('Cierres').datos.slice(1).length, 3); // no duplica filas viejas
});

prueba('al día siguiente arranca en cero y el día anterior se cierra solo', () => {
  const env = crearEntorno(); const { gas } = env; preparar(gas);
  gas.guardarCobro(cobro());
  env.setAhora(new Date('2026-10-22T10:00:00-03:00'));
  const e = gas.getEstado();
  assert.strictEqual(e.hoy, '2026-10-22');
  assert.strictEqual(e.resumen.total, 0);
  assert.strictEqual(e.cerrado, false);
  const ayer = gas.getCierre('2026-10-21');
  assert.strictEqual(ayer.cerrado, true);
  assert.strictEqual(ayer.resumen.total, 8000);
});

prueba('cierre automático por activador (también de madrugada)', () => {
  const env = crearEntorno(); const { gas, props } = env; preparar(gas);
  gas.guardarCobro(cobro());
  env.setAhora(new Date('2026-10-21T23:30:00-03:00'));
  gas.cierreAutomatico();
  assert.strictEqual(props.get('CERRADO_2026-10-21'), 'Automático');
  // madrugada: cierra el día anterior
  env.setAhora(new Date('2026-10-22T12:00:00-03:00'));
  gas.guardarCobro(cobro());
  env.setAhora(new Date('2026-10-23T01:30:00-03:00'));
  gas.cierreAutomatico();
  assert.strictEqual(props.get('CERRADO_2026-10-22'), 'Automático');
});

prueba('sin PIN asignado no se puede entrar a administración', () => {
  const { gas } = crearEntorno(); gas.setup();
  lanza(() => gas.adminEntrar('1234'), 'Todavía no hay un PIN');
  lanza(() => gas.reabrirDia('1234', '2026-10-21'), 'Todavía no hay un PIN');
});

prueba('el instalador asigna y resetea el PIN del cliente', () => {
  const { gas } = crearEntorno(); preparar(gas);
  lanza(() => gas.asignarPin_('12'), 'entre 4 y 8');
  lanza(() => gas.asignarPin_('12ab'), 'entre 4 y 8');
  gas.asignarPin_('5739');
  gas.adminEntrar('5739');
  for (let i = 0; i < 5; i++) lanza(() => gas.adminEntrar('0000'), 'PIN incorrecto');
  lanza(() => gas.adminEntrar('5739'), 'Demasiados intentos');
  gas.asignarPin_('8642'); // reset: también destraba el bloqueo
  lanza(() => gas.adminEntrar('5739'), 'PIN incorrecto');
  gas.adminEntrar('8642');
  assert.strictEqual(typeof gas.adminCambiarPin, 'undefined', 'el cliente no debe poder cambiar el PIN');
});

prueba('administración: puestos y servicios', () => {
  const { gas } = crearEntorno(); preparar(gas);
  lanza(() => gas.adminEntrar('9999'), 'PIN incorrecto');
  let a = gas.adminEntrar('1234');

  a = gas.adminAgregarPuesto('1234');
  assert.strictEqual(a.puestos.length, 11);
  assert.strictEqual(a.puestos[10].numero, 11);
  gas.adminCambiarPuesto('1234', 3, false);
  assert(!gas.getEstado().puestos.includes(3), 'el puesto 3 debería estar oculto');
  lanza(() => gas.guardarCobro(cobro({ puesto: 3 })), 'puesto');

  gas.adminGuardarServicio('1234', { original: 'Corte', nombre: 'Corte', precio: 9000, activo: true });
  gas.adminGuardarServicio('1234', { original: '', nombre: 'Cejas', precio: 3000, activo: true });
  lanza(() => gas.adminGuardarServicio('1234', { original: '', nombre: 'cejas', precio: 1, activo: true }), 'Ya existe');
  const sv = gas.getEstado().servicios;
  assert.strictEqual(sv.find(s => s.nombre === 'Corte').precio, 9000);
  assert(sv.some(s => s.nombre === 'Cejas'));
});

prueba('polirubro: renombrar servicios y nombre del negocio', () => {
  const { gas } = crearEntorno(); preparar(gas);
  gas.guardarCobro(cobro()); // un Corte de antes
  gas.adminGuardarServicio('1234', { original: 'Corte', nombre: 'Uñas', precio: 7000, activo: true });
  gas.adminGuardarServicio('1234', { original: 'Barba', nombre: 'Depilación', precio: 9000, activo: true });
  let e = gas.getEstado();
  assert(e.servicios.some(s => s.nombre === 'Uñas' && s.precio === 7000));
  assert(!e.servicios.some(s => s.nombre === 'Corte'));
  assert.strictEqual(e.ultimos[0].servicio, 'Corte', 'el historial conserva el nombre viejo');
  gas.guardarCobro(cobro({ servicio: 'Uñas', monto: 7000 }));
  assert.strictEqual(gas.getEstado().resumen.atenciones, 2);

  assert.strictEqual(e.nombre, 'Control de caja diario');
  const a = gas.adminNombreNegocio('1234', 'Nails Studio');
  assert.strictEqual(a.nombreNegocio, 'Nails Studio');
  assert.strictEqual(gas.getEstado().nombre, 'Nails Studio');
  lanza(() => gas.adminNombreNegocio('1234', 'x'.repeat(41)), 'muy largo');
  gas.adminNombreNegocio('1234', '');
  assert.strictEqual(gas.getEstado().nombre, 'Control de caja diario');
});

prueba('cambiar un precio no modifica cobros anteriores', () => {
  const { gas } = crearEntorno(); preparar(gas);
  gas.guardarCobro(cobro());
  gas.adminGuardarServicio('1234', { original: 'Corte', nombre: 'Corte', precio: 10000, activo: true });
  assert.strictEqual(gas.getEstado().resumen.total, 8000);
});

prueba('bloqueo después de 5 PIN incorrectos', () => {
  const { gas } = crearEntorno(); preparar(gas);
  for (let i = 0; i < 5; i++) lanza(() => gas.adminEntrar('0000'), 'PIN incorrecto');
  lanza(() => gas.adminEntrar('1234'), 'Demasiados intentos');
});

prueba('rinde con un año de historial (110.000 cobros)', () => {
  const env = crearEntorno(); const { gas, hojas } = env; preparar(gas);
  const hoja = hojas.get('Cobros');
  const fila = (fecha, i) => ['id' + i, fecha, '10:00', 1 + (i % 12), 'Corte', 8000, 'Efectivo', '', '', new Date()];
  let i = 0;
  for (let d = 0; d < 300; d++) {
    const f = new Date(Date.UTC(2025, 11, 1 + d)).toISOString().slice(0, 10);
    for (let k = 0; k < 360; k++) hoja.datos.push(fila(f, i++));
  }
  const t = Date.now();
  gas.guardarCobro(cobro());
  const e = gas.getEstado();
  assert.strictEqual(e.resumen.atenciones, 1);
  assert(Date.now() - t < 2000, 'demasiado lento');
});

console.log('\n' + pasadas + ' pruebas pasadas' + (process.exitCode ? ', con errores' : '') + '\n');
