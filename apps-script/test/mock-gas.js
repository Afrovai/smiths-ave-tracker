// Simulador mínimo de las APIs de Google Apps Script que usa Code.gs, para
// poder probar la lógica en Node ANTES de pedirte que la despliegues.
// No toca Drive ni Sheets reales — todo vive en memoria.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet; this.row = row; this.col = col;
    this.numRows = numRows; this.numCols = numCols;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        rowArr.push(this.sheet._get(this.row + r, this.col + c));
      }
      out.push(rowArr);
    }
    return out;
  }
  setValues(values) {
    values.forEach((rowArr, r) => rowArr.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); }
}

class MockSheet {
  constructor(name, grid) {
    this.name = name;
    this.grid = grid || []; // grid[rowIdx][colIdx], 0-indexed
  }
  getName() { return this.name; }
  _get(row, col) { // 1-indexed
    const r = this.grid[row - 1];
    if (!r) return '';
    const v = r[col - 1];
    return v === undefined ? '' : v;
  }
  _set(row, col, val) {
    while (this.grid.length < row) this.grid.push([]);
    this.grid[row - 1][col - 1] = val;
  }
  getLastRow() {
    for (let r = this.grid.length - 1; r >= 0; r--) {
      if ((this.grid[r] || []).some((v) => v !== '' && v !== undefined && v !== null)) return r + 1;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this.grid.forEach((row) => {
      for (let c = row.length - 1; c >= 0; c--) {
        if (row[c] !== '' && row[c] !== undefined && row[c] !== null) { max = Math.max(max, c + 1); break; }
      }
    });
    return max;
  }
  getRange(row, col, numRows, numCols) {
    numRows = numRows || 1; numCols = numCols || 1;
    return new MockRange(this, row, col, numRows, numCols);
  }
  appendRow(values) {
    const nextRow = this.getLastRow() + 1;
    values.forEach((v, i) => this._set(nextRow, i + 1, v));
  }
  deleteRow(row) {
    this.grid.splice(row - 1, 1);
  }
  setFrozenRows() {}
}

class MockSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  insertSheet(name) {
    const s = new MockSheet(name, []);
    this.sheets[name] = s;
    return s;
  }
  _seed(name, grid) { this.sheets[name] = new MockSheet(name, grid); }
}

function buildSandbox(ss, opts) {
  opts = opts || {};
  const driveLog = [];
  const folders = {};
  const DriveApp = {
    getFoldersByName(name) {
      const folder = folders[name];
      let i = 0;
      return { hasNext: () => i < (folder ? 1 : 0), next: () => { i++; return folder; } };
    },
    createFolder(name) {
      const folder = { name, files: [] };
      folders[name] = folder;
      return folder;
    }
  };
  // añade createFile a cada folder mock
  const origCreateFolder = DriveApp.createFolder;
  DriveApp.createFolder = function (name) {
    const folder = origCreateFolder(name);
    folder.createFile = function (blob) {
      const file = { name: blob.name, bytesLength: blob.bytes.length, mimeType: blob.mimeType, getUrl: () => 'https://drive.google.com/mock/' + encodeURIComponent(blob.name) };
      driveLog.push({ action: 'createFile', folder: name, file: blob.name, mimeType: blob.mimeType, size: blob.bytes.length });
      folder.files.push(file);
      return file;
    };
    return folder;
  };
  // por si getFoldersByName encuentra una ya creada
  Object.defineProperty(DriveApp, '_driveLog', { value: driveLog });

  const Utilities = {
    formatDate(date, tz, fmt) {
      const dd = String(date.getDate()).padStart(2, '0');
      const mm = String(date.getMonth() + 1).padStart(2, '0');
      const yyyy = String(date.getFullYear());
      const yy = yyyy.slice(2);
      return fmt.replace('dd', dd).replace('MM', mm).replace('yyyy', yyyy).replace('yy', yy);
    },
    base64Decode(b64) { return Buffer.from(b64, 'base64'); },
    newBlob(bytes, mimeType, name) { return { bytes, mimeType, name }; },
    getUuid() { return 'test-uuid-' + Math.random().toString(36).slice(2, 10); }
  };

  const ContentService = {
    MimeType: { JSON: 'JSON' },
    createTextOutput(text) {
      return { _text: text, setMimeType() { return this; }, getContent() { return this._text; } };
    }
  };

  const SpreadsheetApp = { openById: () => ss };
  const Session = { getScriptTimeZone: () => 'Australia/Perth' };

  // Mock de LockService: en Node no hay concurrencia real dentro de un mismo
  // proceso síncrono, así que basta con un no-op — lo que importa para los
  // tests es que Code.gs pueda llamar getScriptLock()/waitLock()/releaseLock()
  // sin explotar por "LockService is not defined".
  const LockService = {
    getScriptLock() {
      return { waitLock() {}, releaseLock() {} };
    }
  };

  const sandbox = {
    SpreadsheetApp, Utilities, ContentService, Session, DriveApp, LockService,
    console, Object, Number, String, Math, Date, JSON, Array,
    __driveLog: driveLog
  };
  vm.createContext(sandbox);
  return sandbox;
}

function loadCode(sandbox) {
  const src = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  vm.runInContext(src, sandbox, { filename: 'Code.gs' });
  return sandbox;
}

module.exports = { MockSpreadsheet, MockSheet, buildSandbox, loadCode };
