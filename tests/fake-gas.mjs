// A small in-memory fake of the Apps Script services Code.gs uses, so the real
// backend can run under Node in tests.
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

function fakeSheet(name){
  let rows = [];                                   // rows[r][c], 0-based
  let maxCols = 26;
  const cell = (r, c) => (rows[r] && rows[r][c] !== undefined) ? rows[r][c] : '';
  const put = (r, c, v) => { while (rows.length <= r) rows.push([]); rows[r][c] = v; };
  const formats = {};                              // "r,c" → number format, so tests can tell text from numbers
  let filter = null;
  const range = (r, c, nr = 1, nc = 1) => ({
    getValue: () => cell(r - 1, c - 1),
    setNumberFormat: f => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) formats[(r + i) + ',' + (c + j)] = f; return range(r, c, nr, nc); },
    setNumberFormats: fs => { fs.forEach((row, i) => row.forEach((f, j) => { formats[(r + i) + ',' + (c + j)] = f; })); return range(r, c, nr, nc); },
    setBackground: () => range(r, c, nr, nc), setFontColor: () => range(r, c, nr, nc), setFontSize: () => range(r, c, nr, nc),
    setDataValidation: () => range(r, c, nr, nc),
    createFilter: () => { filter = { ref: [r, c, nr, nc], remove: () => { filter = null; } }; return filter; },
    getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => cell(r - 1 + i, c - 1 + j))),
    setValues: v => { v.forEach((row, i) => row.forEach((x, j) => put(r - 1 + i, c - 1 + j, x))); return range(r, c, nr, nc); },
    setValue: v => { put(r - 1, c - 1, v); },
    clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) put(r - 1 + i, c - 1 + j, ''); },
    setFontWeight: () => range(r, c, nr, nc),
  });
  const lastRow = () => { for (let i = rows.length - 1; i >= 0; i--) if ((rows[i] || []).some(x => x !== '' && x !== undefined)) return i + 1; return 0; };
  return {
    name, rows: () => rows,
    getRange: range, getLastRow: lastRow, getMaxColumns: () => maxCols,
    insertColumnsAfter: (_, n) => { maxCols += n; },
    setFrozenRows: () => {},
    setColumnWidth: () => {},
    getFilter: () => filter,
    formats: () => formats,
    clear: () => { rows = []; Object.keys(formats).forEach(k => delete formats[k]); },
    appendRow: v => { const r = lastRow(); v.forEach((x, j) => put(r, j, x)); },
    deleteRow: r => { rows.splice(r - 1, 1); },
  };
}
function fakeBook(){
  const sheets = {};
  return {
    sheets,
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => (sheets[n] = fakeSheet(n)),
    getSheets: () => Object.values(sheets),
    deleteSheet: s => { delete sheets[s.name]; },
    getName: () => 'Test book',
    getSpreadsheetTimeZone: () => 'UTC',
  };
}

export function loadBackend(){
  const book = fakeBook();
  const props = {};
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => book, getUi: () => { throw new Error('no ui'); }, openById: () => book,
      newDataValidation: () => { const b = { requireValueInList: () => b, requireDate: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; } }) },
    Utilities: {
      // UTC, and just the tokens Code.gs uses.
      formatDate: (d, tz, f) => { const iso = new Date(d).toISOString();
        return f.replace("'T'", 'T').replace('yyyy', iso.slice(0, 4)).replace('MM', iso.slice(5, 7)).replace('dd', iso.slice(8, 10))
                .replace('HH', iso.slice(11, 13)).replace('mm', iso.slice(14, 16)).replace('ss', iso.slice(17, 19)); },
      getUuid: () => crypto.randomUUID(), sleep: () => {},
    },
    Logger: { log: () => {} },
    UrlFetchApp: { fetch: () => { throw new Error('offline in tests'); } },
    Session: { getScriptTimeZone: () => 'UTC' },
    console,
  };
  vm.createContext(ctx);
  vm.runInContext(readFileSync(new URL('../Code.gs', import.meta.url), 'utf8'), ctx);
  ctx.setup();
  return { ctx, book, owner: props.API_TOKEN };
}
