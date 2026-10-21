/**
 * Simulador mínimo de Google Apps Script para probar Code.gs en Node.
 * Imita solo lo que usa el sistema: planillas, propiedades, caché, lock y activadores.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

function crearEntorno(opciones = {}) {
  let ahora = opciones.ahora || new Date('2026-10-21T15:00:00-03:00');
  const hojas = new Map();
  const props = new Map();
  const cache = new Map();
  const triggers = [];

  class Rango {
    constructor(hoja, fila, col, filas, cols) { Object.assign(this, { hoja, fila, col, filas, cols }); }
    getValues() {
      const out = [];
      for (let r = 0; r < this.filas; r++) {
        const fila = this.hoja.datos[this.fila - 1 + r] || [];
        const v = [];
        for (let c = 0; c < this.cols; c++) {
          const x = fila[this.col - 1 + c];
          v.push(x === undefined || x === null ? '' : x);
        }
        out.push(v);
      }
      return out;
    }
    setValues(valores) {
      valores.forEach((fila, r) => fila.forEach((x, c) => this.hoja.poner(this.fila + r, this.col + c, x)));
      return this;
    }
    setValue(x) { this.hoja.poner(this.fila, this.col, x); return this; }
    setNumberFormat() { return this; }
    setFontWeight() { return this; }
  }

  class Hoja {
    constructor(nombre) { this.nombre = nombre; this.datos = []; this.protecciones = []; }
    getName() { return this.nombre; }
    poner(fila, col, x) {
      while (this.datos.length < fila) this.datos.push([]);
      this.datos[fila - 1][col - 1] = x;
    }
    getLastRow() {
      for (let i = this.datos.length - 1; i >= 0; i--) {
        if ((this.datos[i] || []).some(x => x !== undefined && x !== '')) return i + 1;
      }
      return 0;
    }
    getRange(a, b, c, d) {
      if (typeof a === 'string') return new Rango(this, 1, 1, 1, 1); // rangos A1 solo para formato
      return new Rango(this, a, b, c || 1, d || 1);
    }
    appendRow(fila) {
      const n = this.getLastRow() + 1;
      fila.forEach((x, i) => this.poner(n, i + 1, x));
    }
    deleteRow(n) { this.datos.splice(n - 1, 1); }
    setFrozenRows() {}
    protect() {
      const p = { setDescription() { return p; }, setWarningOnly() { return p; } };
      this.protecciones.push(p);
      return p;
    }
    getProtections() { return this.protecciones; }
  }

  const ss = {
    getSheetByName: n => hojas.get(n) || null,
    insertSheet: n => { const h = new Hoja(n); hojas.set(n, h); return h; },
    getSheets: () => [...hojas.values()],
    deleteSheet: h => hojas.delete(h.getName()),
    getUrl: () => 'https://docs.google.com/spreadsheets/d/simulada',
    getId: () => 'simulada'
  };

  const zona = 'America/Argentina/Buenos_Aires';
  function formatear(fecha, tz, patron) {
    const partes = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
      timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
    }).formatToParts(fecha).map(p => [p.type, p.value]));
    return patron
      .replace('yyyy', partes.year).replace('MM', partes.month).replace('dd', partes.day)
      .replace('HH', partes.hour).replace('mm', partes.minute)
      .replace(/^H$/, String(Number(partes.hour)));
  }

  const OrigDate = Date;
  class FakeDate extends OrigDate {
    constructor(...a) { if (a.length === 0) super(ahora.getTime()); else super(...a); }
    static now() { return ahora.getTime(); }
  }
  FakeDate.UTC = OrigDate.UTC;

  const contexto = {
    Date: FakeDate,
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      flush() {},
      getUi: () => ({ createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }),
      ProtectionType: { SHEET: 'SHEET' }
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => (props.has(k) ? props.get(k) : null),
        setProperty: (k, v) => props.set(k, String(v)),
        setProperties: o => Object.entries(o).forEach(([k, v]) => props.set(k, String(v))),
        deleteProperty: k => props.delete(k)
      })
    },
    CacheService: {
      getScriptCache: () => ({
        get: k => (cache.has(k) ? cache.get(k) : null),
        put: (k, v) => cache.set(k, v),
        remove: k => cache.delete(k)
      })
    },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: fn => {
        const t = { fn, getHandlerFunction: () => fn };
        const b = {
          timeBased: () => b, everyDays: () => b, nearMinute: () => b, inTimezone: () => b,
          atHour: h => { t.hora = h; return b; },
          create: () => { triggers.push(t); return t; }
        };
        return b;
      }
    },
    Session: { getScriptTimeZone: () => zona },
    Utilities: {
      formatDate: formatear,
      getUuid: () => crypto.randomUUID(),
      DigestAlgorithm: { SHA_256: 'sha256' },
      Charset: { UTF_8: 'utf8' },
      computeDigest: (alg, texto) => [...crypto.createHash(alg).update(texto, 'utf8').digest()].map(b => (b > 127 ? b - 256 : b))
    },
    HtmlService: {}
  };
  vm.createContext(contexto);
  const codigo = fs.readFileSync(path.join(__dirname, '..', 'src', 'Code.gs'), 'utf8');
  vm.runInContext(codigo, contexto, { filename: 'Code.gs' });

  return {
    gas: contexto,
    hojas,
    props,
    triggers,
    setAhora: d => { ahora = d; }
  };
}

module.exports = { crearEntorno };
