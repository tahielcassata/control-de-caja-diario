/**
 * Control de caja diario
 * Sistema de cobros para negocios de servicios (barberías, peluquerías, uñas, estética…).
 * Backend en Google Apps Script + Google Sheets.
 *
 * Hojas que usa (las crea `setup()`):
 *   Cobros     → una fila por cobro (la escribe solo el sistema)
 *   Cierres    → resumen de cada día, por puesto y total
 *   Puestos    → puestos de trabajo (número y si está activo)
 *   Servicios  → servicios, precio y si está activo
 *
 * Configuración guardada en Script Properties:
 *   PIN_HASH, PIN_SALT, HORA_CIERRE, NOMBRE_NEGOCIO, CERRADO_<yyyy-MM-dd>
 */

// ───────────────────────── Constantes ─────────────────────────

var HOJA_COBROS = 'Cobros';
var HOJA_CIERRES = 'Cierres';
var HOJA_PUESTOS = 'Puestos';
var HOJA_SERVICIOS = 'Servicios';

var COLS_COBROS = ['ID', 'Fecha', 'Hora', 'Puesto', 'Servicio', 'Monto', 'Medio', 'Propina', 'Medio propina', 'Registrado'];
var COLS_CIERRES = ['Fecha', 'Puesto', 'Atenciones', 'Cobrado', 'Propinas', 'Efectivo', 'Transferencias', 'Total', 'Tipo', 'Cerrado el'];
var COLS_PUESTOS = ['Puesto', 'Activo'];
var COLS_SERVICIOS = ['Servicio', 'Precio', 'Activo'];

var MEDIOS = ['Efectivo', 'Transferencia'];
var HORA_CIERRE_INICIAL = 23;
var MONTO_MAXIMO = 10000000;
var BLOQUE_LECTURA = 2000;
var INTENTOS_POR_BLOQUEO = 5;
// Cada 5 intentos fallidos seguidos el bloqueo es más largo: 5 minutos, 1 hora, 24 horas.
var BLOQUEOS_PIN_MIN = [5, 60, 24 * 60];
var NOMBRE_INICIAL = 'Control de caja diario';

var LICENCIA_CACHE_SEG = 60 * 60; // consulta el panel de licencias como mucho una vez por hora

/** Servicios típicos por rubro, para cargar de una vez desde Administración. Precios de ejemplo. */
var RUBROS = {
  'Barbería': [['Corte', 8000], ['Barba', 5000], ['Corte y barba', 12000], ['Perfilado de cejas', 3000]],
  'Peluquería': [['Corte', 10000], ['Brushing', 8000], ['Color', 20000], ['Mechas', 30000], ['Alisado', 35000]],
  'Uñas': [['Esmaltado semipermanente', 9000], ['Uñas esculpidas', 14000], ['Manicura', 6000], ['Pedicura', 8000], ['Retiro', 4000]],
  'Estética': [['Limpieza facial', 15000], ['Depilación', 9000], ['Perfilado de cejas', 5000], ['Lifting de pestañas', 12000]],
  'Masajes y spa': [['Masaje descontracturante', 18000], ['Masaje relajante', 16000], ['Drenaje linfático', 17000], ['Piedras calientes', 20000]]
};

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
    .addItem('Configurar licencia', 'configurarLicencia')
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

  // Arranca sin servicios: cada negocio carga los de su rubro desde Administración.
  var servicios = crearHoja_(ss, HOJA_SERVICIOS, COLS_SERVICIOS);
  servicios.getRange('B:B').setNumberFormat('$#,##0');

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
  var r = ui.prompt('Asignar PIN del cliente', 'Entre 6 y 8 números, al azar (no uses 123456 ni fechas). Es el PIN para entrar a Administración.', ui.ButtonSet.OK_CANCEL);
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
  if (!/^\d{6,8}$/.test(pin)) throw new Error('El PIN tiene que tener entre 6 y 8 números.');
  if (esPinDebil(pin)) throw new Error('Ese PIN es muy fácil de adivinar. Usá números al azar.');
  guardarPin_(pin);
  PropertiesService.getScriptProperties().deleteProperty('PIN_FALLOS');
  PropertiesService.getScriptProperties().deleteProperty('PIN_BLOQUEO_HASTA');
}

/** Rechaza PIN obvios: todos iguales (111111) o en escalera (123456, 987654). */
function esPinDebil(pin) {
  if (/^(\d)\1+$/.test(pin)) return true;
  var sube = true, baja = true;
  for (var i = 1; i < pin.length; i++) {
    var d = Number(pin[i]) - Number(pin[i - 1]);
    if (d !== 1 && d !== -9) sube = false;
    if (d !== -1 && d !== 9) baja = false;
  }
  return sube || baja;
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
    nombre: nombreNegocio_(),
    licencia: licencia_(),
    hoy: hoy,
    cerrado: estaCerrado_(hoy),
    puestos: puestos.filter(function (p) { return p.activo; }).map(function (p) { return p.numero; }),
    servicios: leerServicios_().filter(function (s) { return s.activo; }),
    resumen: resumen,
    ultimos: cobros.slice(-5).reverse(),
    urlPlanilla: getSS_().getUrl()
  };
}

/**
 * Guarda un cobro. La pantalla manda:
 *   id      → código único generado en la tablet (evita duplicados si se reintenta)
 *   cuando  → momento real del cobro (sirve cuando se cargó sin internet y se sube después)
 * Es idempotente: si llega dos veces el mismo id, se guarda una sola vez.
 */
function guardarCobro(datos) {
  datos = datos || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  var fecha, recalcular = false, duplicado = false;
  try {
    var ahora = new Date();
    var momento = momentoDelCobro_(datos.cuando, ahora);
    fecha = Utilities.formatDate(momento, zona_(), 'yyyy-MM-dd');
    var id = String(datos.id || '');
    if (id && !/^[A-Za-z0-9_-]{8,64}$/.test(id)) throw new Error('Código de cobro inválido.');

    if (id && cobrosDeFecha_(fecha).some(function (c) { return c.id === id; })) {
      duplicado = true;
    } else {
      if (estaCerrado_(fecha)) {
        var cierre = momentoCierre_(fecha);
        // Se acepta solo si el cobro se hizo antes de cerrar (por ejemplo, quedó sin subir por falta de internet).
        if (cierre && momento.getTime() >= cierre.getTime()) {
          throw new Error('El día ya está cerrado. Para cargar un cobro tarde, reabrilo desde Cierre del día.');
        }
        recalcular = true;
      }
      var diferido = ahora.getTime() - momento.getTime() > 2 * 60 * 1000;
      // Con la licencia suspendida no se cargan cobros nuevos. Los que quedaron
      // pendientes por falta de internet sí se guardan: ya ocurrieron.
      if (!diferido && licencia_().estado === 'suspendido') {
        throw new Error('El sistema está suspendido por falta de pago. Tus datos siguen guardados.');
      }
      var cobro = validarCobro(datos, leerPuestos_(), leerServicios_(), diferido);
      getHoja_(HOJA_COBROS).appendRow([
        id || Utilities.getUuid(),
        fecha,
        Utilities.formatDate(momento, zona_(), 'HH:mm'),
        cobro.puesto,
        cobro.servicio,
        cobro.monto,
        cobro.medio,
        cobro.propina || '',
        cobro.propina ? cobro.medioPropina : '',
        ahora
      ]);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
  if (recalcular) recalcularCierre_(fecha);
  var estado = getEstado();
  estado.duplicado = duplicado;
  return estado;
}

/** Usa la hora real del cobro si es razonable (hasta 3 días atrás, no futura); si no, la actual. */
function momentoDelCobro_(cuando, ahora) {
  if (!cuando) return ahora;
  var m = new Date(cuando);
  if (isNaN(m.getTime())) return ahora;
  var diff = ahora.getTime() - m.getTime();
  if (diff < -2 * 60 * 1000 || diff > 3 * 24 * 60 * 60 * 1000) return ahora;
  return diff < 0 ? ahora : m;
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

/**
 * Carga los servicios típicos de un rubro. No toca los que ya existen
 * (mismo nombre), así que se puede usar más de una vez o combinar rubros.
 */
function adminCargarRubro(pin, rubro) {
  verificarPin_(pin);
  var lista = RUBROS[rubro];
  if (!lista) throw new Error('Rubro desconocido.');
  var existentes = leerServicios_().map(function (s) { return s.nombre.toLowerCase(); });
  var nuevos = lista.filter(function (s) { return existentes.indexOf(s[0].toLowerCase()) === -1; })
    .map(function (s) { return [s[0], s[1], true]; });
  if (nuevos.length) {
    var hoja = getHoja_(HOJA_SERVICIOS);
    hoja.getRange(hoja.getLastRow() + 1, 1, nuevos.length, 3).setValues(nuevos);
  }
  var config = adminConfig_();
  config.agregados = nuevos.length;
  return config;
}

/** Nombre que aparece arriba de la pantalla (por ejemplo, el nombre del local). */
function adminNombreNegocio(pin, nombre) {
  verificarPin_(pin);
  nombre = String(nombre || '').trim();
  if (nombre.length > 40) throw new Error('El nombre es muy largo (máximo 40 letras).');
  var props = PropertiesService.getScriptProperties();
  if (nombre) props.setProperty('NOMBRE_NEGOCIO', nombre);
  else props.deleteProperty('NOMBRE_NEGOCIO');
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

// ───────────────────────── Licencia ─────────────────────────

/**
 * Lo usa el instalador desde la planilla: menú Control de caja → Configurar licencia.
 * Se pega la "línea de instalación" que da el panel de licencias: URL|CÓDIGO|CLAVE.
 * Vacío = sin control de licencia.
 */
function configurarLicencia() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Configurar licencia', 'Pegá la línea de instalación del panel de licencias (o dejalo vacío para quitar el control).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    var lic = configurarLicencia_(r.getResponseText());
    ui.alert(lic ? 'Licencia configurada. Estado actual: ' + lic.estado + '.' : 'Control de licencia quitado.');
  } catch (e) {
    ui.alert(e.message);
  }
}

function configurarLicencia_(linea) {
  var props = PropertiesService.getScriptProperties();
  CacheService.getScriptCache().remove('LICENCIA');
  linea = String(linea || '').trim();
  if (!linea) {
    ['LIC_URL', 'LIC_CODIGO', 'LIC_CLAVE', 'LIC_ULTIMA'].forEach(function (k) { props.deleteProperty(k); });
    return null;
  }
  var partes = linea.split('|');
  if (partes.length !== 3 || !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(partes[0]) || !/^[A-Z0-9]{4,12}$/.test(partes[1]) || partes[2].length < 16) {
    throw new Error('La línea no es válida. Copiala completa desde el panel de licencias.');
  }
  props.setProperties({ LIC_URL: partes[0], LIC_CODIGO: partes[1], LIC_CLAVE: partes[2] });
  props.deleteProperty('LIC_ULTIMA');
  return licencia_();
}

/**
 * Estado de la licencia: { estado: 'activo'|'aviso'|'suspendido', suspendeEl, contacto }.
 * Consulta el panel como mucho una vez por hora. Si el panel no responde o no hay internet,
 * usa el último estado conocido; si nunca respondió, el sistema sigue activo.
 * Un problema del panel nunca le corta la caja al cliente.
 */
function licencia_() {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty('LIC_URL');
  if (!url) return { estado: 'activo', suspendeEl: null, contacto: '' };
  var cache = CacheService.getScriptCache();
  var enCache = cache.get('LICENCIA');
  if (enCache) return JSON.parse(enCache);

  var ultima = null;
  try { ultima = JSON.parse(props.getProperty('LIC_ULTIMA') || 'null'); } catch (e) { ultima = null; }
  var resultado = ultima || { estado: 'activo', suspendeEl: null, contacto: '' };
  try {
    var resp = UrlFetchApp.fetch(url + '?c=' + encodeURIComponent(props.getProperty('LIC_CODIGO')) + '&k=' + encodeURIComponent(props.getProperty('LIC_CLAVE')), {
      muteHttpExceptions: true, followRedirects: true
    });
    if (resp.getResponseCode() === 200) {
      var r = JSON.parse(resp.getContentText());
      if (r && ['activo', 'aviso', 'suspendido'].indexOf(r.estado) !== -1) {
        resultado = {
          estado: r.estado,
          suspendeEl: /^\d{4}-\d{2}-\d{2}$/.test(r.suspendeEl || '') ? r.suspendeEl : null,
          contacto: String(r.contacto || '').slice(0, 120)
        };
        props.setProperty('LIC_ULTIMA', JSON.stringify(resultado));
      }
    }
  } catch (e) {
    // Sin respuesta del panel: queda el último estado conocido.
  }
  cache.put('LICENCIA', JSON.stringify(resultado), LICENCIA_CACHE_SEG);
  return resultado;
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
  var r = { atenciones: 0, servicios: 0, propinas: 0, efectivo: 0, transferencias: 0, total: 0 };

  cobros.forEach(function (c) {
    var p = porPuesto[c.puesto] || (porPuesto[c.puesto] = filaVacia_(c.puesto));
    var monto = Number(c.monto) || 0;
    var propina = Number(c.propina) || 0;
    p.atenciones += 1;
    p.servicios += monto;
    p.propinas += propina;
    r.atenciones += 1;
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
  return { puesto: Number(n), atenciones: 0, servicios: 0, propinas: 0, efectivo: 0, transferencias: 0 };
}

/** Valida y normaliza un cobro. Lanza un Error con un mensaje para el usuario. */
function validarCobro(d, puestos, servicios, diferido) {
  d = d || {};
  var puesto = Number(d.puesto);
  var servicio = String(d.servicio || '').trim();
  if (diferido) {
    // Cobro cargado sin internet: el puesto o el servicio pudieron cambiar en el medio.
    // Se guarda igual, porque el cobro ya ocurrió.
    if (!(puesto >= 1) || Math.floor(puesto) !== puesto) throw new Error('Elegí un puesto válido.');
    if (!servicio || servicio.length > 60) throw new Error('Elegí un servicio válido.');
  } else {
    var existePuesto = puestos.some(function (p) { return p.numero === puesto && p.activo; });
    if (!existePuesto) throw new Error('Elegí un puesto válido.');
    var existeServicio = servicios.some(function (s) { return s.nombre === servicio && s.activo; });
    if (!existeServicio) throw new Error('Elegí un servicio válido.');
  }

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
    if (resumen.atenciones === 0 && tipo === 'Automático') {
      // Día sin cobros (franco): se marca cerrado sin escribir filas.
      PropertiesService.getScriptProperties().setProperty('CERRADO_' + fecha, tipo + '|' + new Date().toISOString());
      return;
    }
    borrarCierreGuardado_(fecha);
    var ahora = new Date();
    var filas = resumen.porPuesto.map(function (p) {
      return [fecha, p.puesto, p.atenciones, p.servicios, p.propinas, p.efectivo, p.transferencias, p.efectivo + p.transferencias, tipo, ahora];
    });
    filas.push([fecha, 'TOTAL', resumen.atenciones, resumen.servicios, resumen.propinas, resumen.efectivo, resumen.transferencias, resumen.total, tipo, ahora]);
    var hoja = getHoja_(HOJA_CIERRES);
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, COLS_CIERRES.length).setValues(filas);
    PropertiesService.getScriptProperties().setProperty('CERRADO_' + fecha, tipo + '|' + ahora.toISOString());
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

/** Momento en que se cerró el día (null si no se sabe). */
function momentoCierre_(fecha) {
  var v = PropertiesService.getScriptProperties().getProperty('CERRADO_' + fecha) || '';
  var partes = v.split('|');
  if (partes.length < 2) return null;
  var d = new Date(partes[1]);
  return isNaN(d.getTime()) ? null : d;
}

/** Vuelve a calcular el cierre de un día ya cerrado (llegó un cobro que estaba pendiente). */
function recalcularCierre_(fecha) {
  var props = PropertiesService.getScriptProperties();
  var tipo = (props.getProperty('CERRADO_' + fecha) || 'Automático').split('|')[0];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    borrarCierreGuardado_(fecha);
    props.deleteProperty('CERRADO_' + fecha);
  } finally {
    lock.releaseLock();
  }
  cerrarFecha_(fecha, tipo);
}

function leerCierreGuardado_(fecha) {
  var filas = filasCierre_(fecha);
  if (!filas.length) return null;
  var r = { porPuesto: [] };
  filas.forEach(function (f) {
    var v = f.valores;
    if (v[1] === 'TOTAL') {
      r.atenciones = Number(v[2]); r.servicios = Number(v[3]); r.propinas = Number(v[4]);
      r.efectivo = Number(v[5]); r.transferencias = Number(v[6]); r.total = Number(v[7]);
    } else {
      r.porPuesto.push({
        puesto: Number(v[1]), atenciones: Number(v[2]), servicios: Number(v[3]),
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
    horaCierre: Number(props.getProperty('HORA_CIERRE') || HORA_CIERRE_INICIAL),
    nombreNegocio: nombreNegocio_(),
    rubros: Object.keys(RUBROS)
  };
}

// ───────────────────────── PIN ─────────────────────────

function verificarPin_(pin) {
  var props = PropertiesService.getScriptProperties();
  var hasta = Number(props.getProperty('PIN_BLOQUEO_HASTA') || 0);
  var ahora = Date.now();
  if (hasta > ahora) throw new Error('Demasiados intentos. Probá de nuevo en ' + duracionTexto_(hasta - ahora) + '.');
  if (!props.getProperty('PIN_HASH')) throw new Error('Todavía no hay un PIN asignado. Pedíselo a quien instaló el sistema.');
  var ok = hashPin_(String(pin || ''), props.getProperty('PIN_SALT')) === props.getProperty('PIN_HASH');
  if (ok) {
    props.deleteProperty('PIN_FALLOS');
    props.deleteProperty('PIN_BLOQUEO_HASTA');
    return;
  }
  var fallos = Number(props.getProperty('PIN_FALLOS') || 0) + 1;
  props.setProperty('PIN_FALLOS', String(fallos));
  if (fallos % INTENTOS_POR_BLOQUEO === 0) {
    var nivel = Math.min(fallos / INTENTOS_POR_BLOQUEO, BLOQUEOS_PIN_MIN.length) - 1;
    var minutos = BLOQUEOS_PIN_MIN[nivel];
    props.setProperty('PIN_BLOQUEO_HASTA', String(ahora + minutos * 60 * 1000));
    throw new Error('PIN incorrecto. Por seguridad, Administración queda bloqueada ' + duracionTexto_(minutos * 60 * 1000) + '.');
  }
  throw new Error('PIN incorrecto.');
}

function duracionTexto_(ms) {
  var min = Math.ceil(ms / 60000);
  if (min < 60) return min + (min === 1 ? ' minuto' : ' minutos');
  var h = Math.ceil(min / 60);
  return h + (h === 1 ? ' hora' : ' horas');
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

function nombreNegocio_() {
  return PropertiesService.getScriptProperties().getProperty('NOMBRE_NEGOCIO') || NOMBRE_INICIAL;
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
