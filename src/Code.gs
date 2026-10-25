/**
 * Control de caja diario
 * Sistema de cobros para barberías.
 * Backend en Google Apps Script + Google Sheets.
 *
 * Hojas que usa (las crea `setup()`):
 *   Cobros     → una fila por cobro (la escribe solo el sistema)
 *   Cierres    → resumen de cada día, por puesto y total
 *   Puestos    → puestos de trabajo (número y si está activo)
 *   Servicios  → servicios, precio y si está activo
 *
 * Configuración guardada en Script Properties:
 *   PIN_HASH, PIN_SALT, HORA_CIERRE, CERRADO_<yyyy-MM-dd>
 */

// ───────────────────────── Constantes ─────────────────────────

var HOJA_COBROS = 'Cobros';
var HOJA_CIERRES = 'Cierres';
var HOJA_PUESTOS = 'Puestos';
var HOJA_SERVICIOS = 'Servicios';

var COLS_COBROS = ['ID', 'Fecha', 'Hora', 'Puesto', 'Servicio', 'Monto', 'Medio', 'Propina', 'Medio propina', 'Registrado'];
var COLS_CIERRES = ['Fecha', 'Puesto', 'Cortes', 'Servicios', 'Propinas', 'Efectivo', 'Transferencias', 'Total', 'Tipo', 'Cerrado el'];
var COLS_PUESTOS = ['Puesto', 'Activo'];
var COLS_SERVICIOS = ['Servicio', 'Precio', 'Activo'];

var MEDIOS = ['Efectivo', 'Transferencia'];
var HORA_CIERRE_INICIAL = 23;
var MONTO_MAXIMO = 10000000;
var BLOQUE_LECTURA = 2000;
var MAX_INTENTOS_PIN = 5;

// ───────────────────────── Web app ─────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Control de caja diario')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Control de caja')
    .addItem('Configurar sistema (primera vez)', 'setup')
    .addItem('Asignar PIN del cliente', 'asignarPinCliente')
    .addToUi();
}

// ───────────────────────── Instalación ─────────────────────────

/**
 * Ejecutar una vez desde el editor. Es seguro volver a ejecutarlo:
 * no borra datos, solo completa lo que falte.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();

  var cobros = crearHoja_(ss, HOJA_COBROS, COLS_COBROS);
  cobros.getRange('B:C').setNumberFormat('@');
  cobros.getRange('F:F').setNumberFormat('$#,##0');
  cobros.getRange('H:H').setNumberFormat('$#,##0');
  protegerHoja_(cobros);

  var cierres = crearHoja_(ss, HOJA_CIERRES, COLS_CIERRES);
  cierres.getRange('A:A').setNumberFormat('@');
  cierres.getRange('D:H').setNumberFormat('$#,##0');
  protegerHoja_(cierres);

  var puestos = crearHoja_(ss, HOJA_PUESTOS, COLS_PUESTOS);
  if (puestos.getLastRow() < 2) {
    var filas = [];
    for (var i = 1; i <= 10; i++) filas.push([i, true]);
    puestos.getRange(2, 1, filas.length, 2).setValues(filas);
  }

  var servicios = crearHoja_(ss, HOJA_SERVICIOS, COLS_SERVICIOS);
  servicios.getRange('B:B').setNumberFormat('$#,##0');
  if (servicios.getLastRow() < 2) {
    servicios.getRange(2, 1, 4, 3).setValues([
      ['Corte', 8000, true],
      ['Barba', 5000, true],
      ['Corte y barba', 12000, true],
      ['Color', 15000, true]
    ]);
  }

  if (!props.getProperty('HORA_CIERRE')) props.setProperty('HORA_CIERRE', String(HORA_CIERRE_INICIAL));
  instalarCierreAutomatico_(Number(props.getProperty('HORA_CIERRE')));

  var hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);

  return 'Listo. Ahora asigná el PIN del cliente desde el menú Control de caja → Asignar PIN del cliente.';
}

/**
 * Lo usa el instalador desde la planilla: menú Control de caja → Asignar PIN del cliente.
 * Sirve para el PIN inicial y para resetearlo si el cliente se lo olvida.
 * No se puede llamar desde la web app (necesita la interfaz de la planilla).
 */
function asignarPinCliente() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Asignar PIN del cliente', 'Entre 4 y 8 números. Es el PIN para entrar a Administración.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    asignarPin_(r.getResponseText().trim());
    ui.alert('PIN asignado. Pasáselo al dueño del local.');
  } catch (e) {
    ui.alert(e.message);
  }
}

function asignarPin_(pin) {
  pin = String(pin || '');
  if (!/^\d{4,8}$/.test(pin)) throw new Error('El PIN tiene que tener entre 4 y 8 números.');
  guardarPin_(pin);
  CacheService.getScriptCache().remove('PIN_INTENTOS');
}

function crearHoja_(ss, nombre, columnas) {
  var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, columnas.length).setValues([columnas]).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function protegerHoja_(hoja) {
  var existentes = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (existentes.length) return;
  hoja.protect()
    .setDescription('La escribe el sistema. No editar a mano.')
    .setWarningOnly(true);
}

function instalarCierreAutomatico_(hora) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cierreAutomatico') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cierreAutomatico')
    .timeBased()
    .everyDays(1)
    .atHour(hora)
    .nearMinute(30)
    .inTimezone(zona_())
    .create();
}

// ───────────────────────── API pública (pantalla de cobro) ─────────────────────────

/** Todo lo que necesita la pantalla principal en una sola llamada. */
function getEstado() {
  cerrarAyerSiQuedoAbierto_();
  var hoy = fechaHoy_();
  var cobros = cobrosDeFecha_(hoy);
  var puestos = leerPuestos_();
  var resumen = calcularResumen(cobros, puestos.filter(function (p) { return p.activo; }).map(function (p) { return p.numero; }));
  return {
    hoy: hoy,
    cerrado: estaCerrado_(hoy),
    puestos: puestos.filter(function (p) { return p.activo; }).map(function (p) { return p.numero; }),
    servicios: leerServicios_().filter(function (s) { return s.activo; }),
    resumen: resumen,
    ultimos: cobros.slice(-5).reverse(),
    urlPlanilla: getSS_().getUrl()
  };
}

function guardarCobro(datos) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var hoy = fechaHoy_();
    if (estaCerrado_(hoy)) throw new Error('El día ya está cerrado. Para cargar un cobro tarde, reabrilo desde Cierre del día.');

    var cobro = validarCobro(datos, leerPuestos_(), leerServicios_());
    var ahora = new Date();
    getHoja_(HOJA_COBROS).appendRow([
      Utilities.getUuid(),
      hoy,
      Utilities.formatDate(ahora, zona_(), 'HH:mm'),
      cobro.puesto,
      cobro.servicio,
      cobro.monto,
      cobro.medio,
      cobro.propina || '',
      cobro.propina ? cobro.medioPropina : '',
      ahora
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getEstado();
}

function deshacerUltimo() {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  var borrado;
  try {
    var hoy = fechaHoy_();
    if (estaCerrado_(hoy)) throw new Error('El día está cerrado. No se puede deshacer.');
    var hoja = getHoja_(HOJA_COBROS);
    var ultima = hoja.getLastRow();
    if (ultima < 2) throw new Error('No hay cobros para deshacer.');
    var fila = hoja.getRange(ultima, 1, 1, COLS_COBROS.length).getValues()[0];
    if (String(fila[1]) !== hoy) throw new Error('No hay cobros de hoy para deshacer.');
    borrado = filaACobro_(fila);
    hoja.deleteRow(ultima);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  var estado = getEstado();
  estado.borrado = borrado;
  return estado;
}

// ───────────────────────── API pública (cierre del día) ─────────────────────────

/** Resumen de una fecha (yyyy-MM-dd). Si ya se cerró, sale de la hoja Cierres. */
function getCierre(fecha) {
  fecha = fecha || fechaHoy_();
  validarFecha_(fecha);
  var guardado = leerCierreGuardado_(fecha);
  var resumen = guardado || calcularResumen(
    cobrosDeFecha_(fecha),
    leerPuestos_().filter(function (p) { return p.activo; }).map(function (p) { return p.numero; })
  );
  return {
    fecha: fecha,
    hoy: fechaHoy_(),
    cerrado: estaCerrado_(fecha),
    resumen: resumen
  };
}

function cerrarDia() {
  var hoy = fechaHoy_();
  cerrarFecha_(hoy, 'Manual');
  return getCierre(hoy);
}

/** Lo ejecuta el activador diario. */
function cierreAutomatico() {
  var ahora = new Date();
  var hora = Number(Utilities.formatDate(ahora, zona_(), 'H'));
  // Si el cierre está programado de madrugada, corresponde al día anterior.
  var fecha = hora < 6 ? sumarDias_(fechaHoy_(), -1) : fechaHoy_();
  if (!estaCerrado_(fecha)) cerrarFecha_(fecha, 'Automático');
}

// ───────────────────────── API pública (administración, con PIN) ─────────────────────────

function adminEntrar(pin) {
  verificarPin_(pin);
  return adminConfig_();
}

function adminAgregarPuesto(pin) {
  verificarPin_(pin);
  var hoja = getHoja_(HOJA_PUESTOS);
  var puestos = leerPuestos_();
  var siguiente = puestos.reduce(function (max, p) { return Math.max(max, p.numero); }, 0) + 1;
  hoja.appendRow([siguiente, true]);
  return adminConfig_();
}

function adminCambiarPuesto(pin, numero, activo) {
  verificarPin_(pin);
  var hoja = getHoja_(HOJA_PUESTOS);
  var datos = leerTabla_(hoja, COLS_PUESTOS.length);
  for (var i = 0; i < datos.length; i++) {
    if (Number(datos[i][0]) === Number(numero)) {
      hoja.getRange(i + 2, 2).setValue(!!activo);
      return adminConfig_();
    }
  }
  throw new Error('No existe el puesto ' + numero + '.');
}

/**
 * Crea o modifica un servicio.
 * datos = { original: 'Corte' | '', nombre, precio, activo }
 * El cambio de precio no toca los cobros ni los cierres anteriores.
 */
function adminGuardarServicio(pin, datos) {
  verificarPin_(pin);
  var nombre = String(datos && datos.nombre || '').trim();
  var precio = Math.round(Number(datos && datos.precio));
  if (!nombre) throw new Error('Poné un nombre para el servicio.');
  if (nombre.length > 40) throw new Error('El nombre es muy largo (máximo 40 letras).');
  if (!isFinite(precio) || precio < 0 || precio > MONTO_MAXIMO) throw new Error('Revisá el precio.');

  var hoja = getHoja_(HOJA_SERVICIOS);
  var filas = leerTabla_(hoja, COLS_SERVICIOS.length);
  var original = String(datos.original || '').trim().toLowerCase();
  var filaOriginal = -1;
  for (var i = 0; i < filas.length; i++) {
    var actual = String(filas[i][0]).trim().toLowerCase();
    if (original && actual === original) filaOriginal = i;
    else if (actual === nombre.toLowerCase()) throw new Error('Ya existe un servicio llamado "' + nombre + '".');
  }
  var activo = datos.activo !== false;
  if (original && filaOriginal === -1) throw new Error('No se encontró el servicio "' + datos.original + '".');
  if (filaOriginal >= 0) hoja.getRange(filaOriginal + 2, 1, 1, 3).setValues([[nombre, precio, activo]]);
  else hoja.appendRow([nombre, precio, activo]);
  return adminConfig_();
}

function adminHoraCierre(pin, hora) {
  verificarPin_(pin);
  hora = Number(hora);
  if (!(hora >= 0 && hora <= 23) || Math.floor(hora) !== hora) throw new Error('Hora inválida.');
  PropertiesService.getScriptProperties().setProperty('HORA_CIERRE', String(hora));
  instalarCierreAutomatico_(hora);
  return adminConfig_();
}

/** Reabre un día cerrado para cargar un cobro tarde. Al volver a cerrar se recalcula. */
function reabrirDia(pin, fecha) {
  verificarPin_(pin);
  validarFecha_(fecha);
  if (fecha !== fechaHoy_()) throw new Error('Solo se puede reabrir el día de hoy.');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    borrarCierreGuardado_(fecha);
    PropertiesService.getScriptProperties().deleteProperty('CERRADO_' + fecha);
  } finally {
    lock.releaseLock();
  }
  return getCierre(fecha);
}

// ───────────────────────── Lógica pura (testeable) ─────────────────────────

/**
 * Suma un día de cobros.
 * cobros: [{puesto, servicio, monto, medio, propina, medioPropina}]
 * puestos: números de puesto a mostrar aunque no tengan cobros.
 */
function calcularResumen(cobros, puestos) {
  var porPuesto = {};
  (puestos || []).forEach(function (n) { porPuesto[n] = filaVacia_(n); });
  var r = { cortes: 0, servicios: 0, propinas: 0, efectivo: 0, transferencias: 0, total: 0 };

  cobros.forEach(function (c) {
    var p = porPuesto[c.puesto] || (porPuesto[c.puesto] = filaVacia_(c.puesto));
    var monto = Number(c.monto) || 0;
    var propina = Number(c.propina) || 0;
    p.cortes += 1;
    p.servicios += monto;
    p.propinas += propina;
    r.cortes += 1;
    r.servicios += monto;
    r.propinas += propina;
    if (c.medio === 'Efectivo') { r.efectivo += monto; p.efectivo += monto; }
    else { r.transferencias += monto; p.transferencias += monto; }
    if (propina) {
      if (c.medioPropina === 'Efectivo') { r.efectivo += propina; p.efectivo += propina; }
      else { r.transferencias += propina; p.transferencias += propina; }
    }
  });

  r.total = r.efectivo + r.transferencias;
  r.porPuesto = Object.keys(porPuesto)
    .map(function (k) { return porPuesto[k]; })
    .sort(function (a, b) { return a.puesto - b.puesto; });
  return r;
}

function filaVacia_(n) {
  return { puesto: Number(n), cortes: 0, servicios: 0, propinas: 0, efectivo: 0, transferencias: 0 };
}

/** Valida y normaliza un cobro. Lanza un Error con un mensaje para el usuario. */
function validarCobro(d, puestos, servicios) {
  d = d || {};
  var puesto = Number(d.puesto);
  var existePuesto = puestos.some(function (p) { return p.numero === puesto && p.activo; });
  if (!existePuesto) throw new Error('Elegí un puesto válido.');

  var servicio = String(d.servicio || '');
  var existeServicio = servicios.some(function (s) { return s.nombre === servicio && s.activo; });
  if (!existeServicio) throw new Error('Elegí un servicio válido.');

  var monto = Math.round(Number(d.monto));
  if (!isFinite(monto) || monto <= 0 || monto > MONTO_MAXIMO) throw new Error('Revisá el monto.');
  if (MEDIOS.indexOf(d.medio) === -1) throw new Error('Elegí el medio de pago.');

  var propina = d.propina === '' || d.propina == null ? 0 : Math.round(Number(d.propina));
  if (!isFinite(propina) || propina < 0 || propina > MONTO_MAXIMO) throw new Error('Revisá la propina.');
  if (propina > 0 && MEDIOS.indexOf(d.medioPropina) === -1) throw new Error('Elegí el medio de la propina.');

  return {
    puesto: puesto,
    servicio: servicio,
    monto: monto,
    medio: d.medio,
    propina: propina,
    medioPropina: propina > 0 ? d.medioPropina : ''
  };
}

// ───────────────────────── Cierres ─────────────────────────

function cerrarFecha_(fecha, tipo) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (estaCerrado_(fecha)) return;
    var resumen = calcularResumen(cobrosDeFecha_(fecha), []);
    if (resumen.cortes === 0 && tipo === 'Automático') {
      // Día sin cobros (franco): se marca cerrado sin escribir filas.
      PropertiesService.getScriptProperties().setProperty('CERRADO_' + fecha, tipo);
      return;
    }
    borrarCierreGuardado_(fecha);
    var ahora = new Date();
    var filas = resumen.porPuesto.map(function (p) {
      return [fecha, p.puesto, p.cortes, p.servicios, p.propinas, p.efectivo, p.transferencias, p.efectivo + p.transferencias, tipo, ahora];
    });
    filas.push([fecha, 'TOTAL', resumen.cortes, resumen.servicios, resumen.propinas, resumen.efectivo, resumen.transferencias, resumen.total, tipo, ahora]);
    var hoja = getHoja_(HOJA_CIERRES);
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, COLS_CIERRES.length).setValues(filas);
    PropertiesService.getScriptProperties().setProperty('CERRADO_' + fecha, tipo);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

/** Red de seguridad por si el activador no corrió: al abrir la app se cierra el día anterior. */
function cerrarAyerSiQuedoAbierto_() {
  var ayer = sumarDias_(fechaHoy_(), -1);
  if (!estaCerrado_(ayer)) cerrarFecha_(ayer, 'Automático');
}

function estaCerrado_(fecha) {
  return !!PropertiesService.getScriptProperties().getProperty('CERRADO_' + fecha);
}

function leerCierreGuardado_(fecha) {
  var filas = filasCierre_(fecha);
  if (!filas.length) return null;
  var r = { porPuesto: [] };
  filas.forEach(function (f) {
    var v = f.valores;
    if (v[1] === 'TOTAL') {
      r.cortes = Number(v[2]); r.servicios = Number(v[3]); r.propinas = Number(v[4]);
      r.efectivo = Number(v[5]); r.transferencias = Number(v[6]); r.total = Number(v[7]);
    } else {
      r.porPuesto.push({
        puesto: Number(v[1]), cortes: Number(v[2]), servicios: Number(v[3]),
        propinas: Number(v[4]), efectivo: Number(v[5]), transferencias: Number(v[6])
      });
    }
  });
  r.porPuesto.sort(function (a, b) { return a.puesto - b.puesto; });
  return r;
}

function borrarCierreGuardado_(fecha) {
  var hoja = getHoja_(HOJA_CIERRES);
  filasCierre_(fecha)
    .map(function (f) { return f.fila; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (fila) { hoja.deleteRow(fila); });
}

/** Busca desde abajo las filas de Cierres de una fecha. */
function filasCierre_(fecha) {
  var hoja = getHoja_(HOJA_CIERRES);
  var encontradas = [];
  recorrerDesdeAbajo_(hoja, COLS_CIERRES.length, function (valores, fila) {
    var f = String(valores[0]);
    if (f === fecha) encontradas.push({ fila: fila, valores: valores });
    return f >= fecha; // seguir mientras no pasemos a fechas anteriores
  });
  return encontradas;
}

// ───────────────────────── Lectura de datos ─────────────────────────

/**
 * Cobros de una fecha, en orden de carga. Lee la hoja de abajo hacia arriba
 * en bloques, así no se pone lenta aunque tenga años de historial.
 */
function cobrosDeFecha_(fecha) {
  var hoja = getHoja_(HOJA_COBROS);
  var lista = [];
  recorrerDesdeAbajo_(hoja, COLS_COBROS.length, function (valores) {
    var f = String(valores[1]);
    if (f === fecha) lista.push(filaACobro_(valores));
    return f >= fecha;
  });
  return lista.reverse();
}

/** Recorre filas de abajo hacia arriba. `fn` devuelve false para cortar. */
function recorrerDesdeAbajo_(hoja, columnas, fn) {
  var ultima = hoja.getLastRow();
  while (ultima >= 2) {
    var desde = Math.max(2, ultima - BLOQUE_LECTURA + 1);
    var valores = hoja.getRange(desde, 1, ultima - desde + 1, columnas).getValues();
    for (var i = valores.length - 1; i >= 0; i--) {
      if (fn(valores[i], desde + i) === false) return;
    }
    ultima = desde - 1;
  }
}

function filaACobro_(v) {
  return {
    id: String(v[0]),
    fecha: String(v[1]),
    hora: String(v[2]),
    puesto: Number(v[3]),
    servicio: String(v[4]),
    monto: Number(v[5]) || 0,
    medio: String(v[6]),
    propina: Number(v[7]) || 0,
    medioPropina: String(v[8] || '')
  };
}

function leerPuestos_() {
  return leerTabla_(getHoja_(HOJA_PUESTOS), COLS_PUESTOS.length)
    .filter(function (f) { return f[0] !== '' && !isNaN(Number(f[0])); })
    .map(function (f) { return { numero: Number(f[0]), activo: esVerdadero_(f[1]) }; })
    .sort(function (a, b) { return a.numero - b.numero; });
}

function leerServicios_() {
  return leerTabla_(getHoja_(HOJA_SERVICIOS), COLS_SERVICIOS.length)
    .filter(function (f) { return String(f[0]).trim() !== ''; })
    .map(function (f) { return { nombre: String(f[0]).trim(), precio: Number(f[1]) || 0, activo: esVerdadero_(f[2]) }; });
}

function leerTabla_(hoja, columnas) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  return hoja.getRange(2, 1, ultima - 1, columnas).getValues();
}

function adminConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    puestos: leerPuestos_(),
    servicios: leerServicios_(),
    horaCierre: Number(props.getProperty('HORA_CIERRE') || HORA_CIERRE_INICIAL)
  };
}

// ───────────────────────── PIN ─────────────────────────

function verificarPin_(pin) {
  var cache = CacheService.getScriptCache();
  var intentos = Number(cache.get('PIN_INTENTOS') || 0);
  if (intentos >= MAX_INTENTOS_PIN) throw new Error('Demasiados intentos. Esperá 5 minutos.');
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN_HASH')) throw new Error('Todavía no hay un PIN asignado. Pedíselo a quien instaló el sistema.');
  var ok = hashPin_(String(pin || ''), props.getProperty('PIN_SALT')) === props.getProperty('PIN_HASH');
  if (!ok) {
    cache.put('PIN_INTENTOS', String(intentos + 1), 300);
    throw new Error('PIN incorrecto.');
  }
  cache.remove('PIN_INTENTOS');
}

function guardarPin_(pin) {
  var salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({
    PIN_SALT: salt,
    PIN_HASH: hashPin_(pin, salt)
  });
}

function hashPin_(pin, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// ───────────────────────── Utilidades ─────────────────────────

/** El script va pegado a la planilla, así que siempre usa la planilla que lo contiene. */
function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getHoja_(nombre) {
  var hoja = getSS_().getSheetByName(nombre);
  if (!hoja) throw new Error('Falta la hoja "' + nombre + '". Ejecutá setup() desde el editor.');
  return hoja;
}

function zona_() {
  return Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
}

function fechaHoy_() {
  return Utilities.formatDate(new Date(), zona_(), 'yyyy-MM-dd');
}

function sumarDias_(fecha, dias) {
  var p = fecha.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + dias, 12));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function validarFecha_(fecha) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fecha))) throw new Error('Fecha inválida.');
}

function esVerdadero_(v) {
  return v === true || String(v).toLowerCase() === 'true' || String(v).toLowerCase() === 'verdadero';
}
