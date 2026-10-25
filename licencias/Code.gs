/**
 * Panel de licencias — Control de caja diario
 *
 * Proyecto aparte, instalado en la cuenta de Google del desarrollador (no del cliente).
 * Una planilla con una fila por cliente y una web app que responde a cada sistema
 * si está activo, con pago pendiente (aviso) o suspendido.
 *
 * Hoja Clientes:
 *   Cliente | Código | Clave | Modo | Pagado hasta | Días de gracia | Último contacto | Notas
 *   Modo: Automático (por fecha de pago) · Activo (forzar) · Suspendido (forzar)
 *
 * Hoja Config:
 *   Contacto → texto que ve el cliente para comunicarse (por ejemplo "Tahiel · WhatsApp 11 5555-5555")
 */

var HOJA_CLIENTES = 'Clientes';
var HOJA_CONFIG = 'Config';
var COLS = ['Cliente', 'Código', 'Clave', 'Modo', 'Pagado hasta', 'Días de gracia', 'Último contacto', 'Notas'];
var MODOS = ['Automático', 'Activo', 'Suspendido'];
var GRACIA_INICIAL = 5;

// ───────────────────────── Web app: la consultan los sistemas de los clientes ─────────────────────────

function doGet(e) {
  var p = (e && e.parameter) || {};
  var r;
  try {
    r = consultar_(String(p.c || ''), String(p.k || ''));
  } catch (err) {
    r = { estado: 'error' };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function consultar_(codigo, clave) {
  if (!/^[A-Z0-9]{4,12}$/.test(codigo) || !clave) return { estado: 'desconocido' };
  var hoja = hoja_(HOJA_CLIENTES);
  var ultima = hoja.getLastRow();
  if (ultima < 2) return { estado: 'desconocido' };
  var filas = hoja.getRange(2, 1, ultima - 1, COLS.length).getValues();
  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    if (String(f[1]) !== codigo) continue;
    if (!igualSeguro_(String(f[2]), clave)) return { estado: 'desconocido' };
    var lic = calcularLicencia(String(f[3] || 'Automático'), fechaTexto_(f[4]), Number(f[5]), fechaHoy_());
    hoja.getRange(i + 2, 7).setValue(new Date());
    lic.contacto = contacto_();
    return lic;
  }
  return { estado: 'desconocido' };
}

/**
 * Estado de la licencia.
 *   modo         Automático | Activo | Suspendido
 *   pagadoHasta  'yyyy-MM-dd' o '' (sin fecha = activo)
 *   gracia       días de aviso después del vencimiento antes de suspender
 *   hoy          'yyyy-MM-dd'
 */
function calcularLicencia(modo, pagadoHasta, gracia, hoy) {
  if (modo === 'Suspendido') return { estado: 'suspendido', suspendeEl: null };
  if (modo === 'Activo' || !pagadoHasta) return { estado: 'activo', suspendeEl: null };
  gracia = gracia >= 0 ? Math.floor(gracia) : GRACIA_INICIAL;
  var suspendeEl = sumarDias_(pagadoHasta, gracia + 1);
  if (hoy <= pagadoHasta) return { estado: 'activo', suspendeEl: null };
  if (hoy < suspendeEl) return { estado: 'aviso', suspendeEl: suspendeEl };
  return { estado: 'suspendido', suspendeEl: suspendeEl };
}

// ───────────────────────── Menú de la planilla ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Licencias')
    .addItem('Configurar (primera vez)', 'setup')
    .addItem('Nuevo cliente', 'nuevoCliente')
    .addItem('Registrar pago (+1 mes) al cliente seleccionado', 'registrarPago')
    .addItem('Ver línea de instalación del cliente seleccionado', 'verLineaInstalacion')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_CLIENTES) || ss.insertSheet(HOJA_CLIENTES);
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  hoja.getRange('D2:D').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MODOS, true).setAllowInvalid(false).build()
  );
  hoja.getRange('E2:E').setNumberFormat('dd/mm/yyyy');
  hoja.getRange('G2:G').setNumberFormat('dd/mm/yyyy hh:mm');

  var config = ss.getSheetByName(HOJA_CONFIG) || ss.insertSheet(HOJA_CONFIG);
  if (config.getLastRow() === 0) {
    config.getRange(1, 1, 2, 2).setValues([['Dato', 'Valor'], ['Contacto', 'Escribí acá cómo te contactan (WhatsApp, mail)']]);
  }
  var h1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (h1 && h1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(h1);
  return 'Listo. Completá tu contacto en la hoja Config y publicá la web app.';
}

function nuevoCliente() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Nuevo cliente', 'Nombre del negocio:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || !r.getResponseText().trim()) return;
  var fila = altaCliente_(r.getResponseText().trim());
  ui.alert('Cliente agregado', 'Línea de instalación (pegala en el sistema del cliente, menú Control de caja → Configurar licencia):\n\n' + lineaInstalacion_(fila.codigo, fila.clave), ui.ButtonSet.OK);
}

function altaCliente_(nombre) {
  var hoja = hoja_(HOJA_CLIENTES);
  var existentes = hoja.getLastRow() < 2 ? [] : hoja.getRange(2, 2, hoja.getLastRow() - 1, 1).getValues().map(function (f) { return String(f[0]); });
  var codigo;
  do { codigo = aleatorio_('ABCDEFGHJKLMNPQRSTUVWXYZ23456789', 6); } while (existentes.indexOf(codigo) !== -1);
  var clave = aleatorio_('abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', 24);
  hoja.appendRow([nombre, codigo, clave, 'Automático', '', GRACIA_INICIAL, '', '']);
  return { codigo: codigo, clave: clave };
}

/** Suma un mes a "Pagado hasta" (desde hoy si ya estaba vencido). */
function registrarPago() {
  var ui = SpreadsheetApp.getUi();
  var hoja = SpreadsheetApp.getActiveSheet();
  var fila = hoja.getActiveRange() ? hoja.getActiveRange().getRow() : 0;
  if (hoja.getName() !== HOJA_CLIENTES || fila < 2) return ui.alert('Seleccioná una celda en la fila del cliente.');
  var nuevo = sumarMes(fechaTexto_(hoja.getRange(fila, 5).getValue()), fechaHoy_());
  var p = nuevo.split('-').map(Number);
  hoja.getRange(fila, 5).setValue(new Date(p[0], p[1] - 1, p[2]));
  ui.alert('Pago registrado. ' + hoja.getRange(fila, 1).getValue() + ' queda pago hasta el ' + p[2] + '/' + p[1] + '/' + p[0] + '.');
}

/** Un mes más desde la fecha pagada, o desde hoy si ya venció o no tenía fecha. */
function sumarMes(pagadoHasta, hoy) {
  var base = pagadoHasta && pagadoHasta >= hoy ? pagadoHasta : hoy;
  var p = base.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1], 1, 12));
  var ultimoDia = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0, 12)).getUTCDate();
  d.setUTCDate(Math.min(p[2], ultimoDia));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function verLineaInstalacion() {
  var ui = SpreadsheetApp.getUi();
  var hoja = SpreadsheetApp.getActiveSheet();
  var fila = hoja.getActiveRange() ? hoja.getActiveRange().getRow() : 0;
  if (hoja.getName() !== HOJA_CLIENTES || fila < 2) return ui.alert('Seleccioná una celda en la fila del cliente.');
  var v = hoja.getRange(fila, 1, 1, 3).getValues()[0];
  ui.alert(String(v[0]), lineaInstalacion_(String(v[1]), String(v[2])), ui.ButtonSet.OK);
}

function lineaInstalacion_(codigo, clave) {
  var url = ScriptApp.getService().getUrl();
  if (!url) return 'Primero publicá la web app (Implementar → Nueva implementación → Aplicación web, acceso: Cualquier usuario).';
  return url + '|' + codigo + '|' + clave;
}

// ───────────────────────── Utilidades ─────────────────────────

function hoja_(nombre) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h) throw new Error('Falta la hoja ' + nombre + '. Ejecutá setup().');
  return h;
}

function contacto_() {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  if (!h || h.getLastRow() < 2) return '';
  var filas = h.getRange(2, 1, h.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < filas.length; i++) if (String(filas[i][0]).trim() === 'Contacto') return String(filas[i][1]).slice(0, 120);
  return '';
}

function igualSeguro_(a, b) {
  if (a.length !== b.length) return false;
  var r = 0;
  for (var i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

function aleatorio_(alfabeto, largo) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid());
  var s = '';
  for (var i = 0; i < largo; i++) s += alfabeto.charAt(((bytes[i % bytes.length] & 0xff) + i * 7) % alfabeto.length);
  return s;
}

function zona_() { return Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires'; }
function fechaHoy_() { return Utilities.formatDate(new Date(), zona_(), 'yyyy-MM-dd'); }

function fechaTexto_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return Utilities.formatDate(v, zona_(), 'yyyy-MM-dd');
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function sumarDias_(fecha, dias) {
  var p = fecha.split('-').map(Number);
  return Utilities.formatDate(new Date(Date.UTC(p[0], p[1] - 1, p[2] + dias, 12)), 'UTC', 'yyyy-MM-dd');
}
