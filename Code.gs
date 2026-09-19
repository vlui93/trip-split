/**
 * Greater China Trip Splitter — Apps Script backend
 * Bound to a Google Sheet. Deployed as a Web App (doGet / doPost).
 *
 * SETUP (once):
 *   1. Run setup()  — creates/repairs all tabs + headers, seeds people & rates,
 *                     and generates an API token.
 *   2. Deploy > New deployment > Web app
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      (see README for why "Anyone" + token, not "Only myself")
 *   3. Run showToken() and paste the URL + token into the app's Setup screen.
 */

var SHEET_PEOPLE      = 'People';
var SHEET_EXPENSES    = 'Expenses';
var SHEET_RATES       = 'ExchangeRates';
var SHEET_SETTLEMENTS = 'Settlements';
var SHEET_PROJECTS    = 'Projects';

// Only the starting set. The ExchangeRates tab is the real list from then on,
// so currencies can be added later without touching this file.
var DEFAULT_CURRENCIES = ['AUD', 'HKD', 'MOP', 'CNY'];
var CITY_NAMES = ['Hong Kong', 'Macau', 'China', 'Sydney'];

/** Every currency the sheet knows about. AUD is always first and always present. */
function CURRENCIES() {
  var list = readRows(SHEET_RATES)
    .map(function (r) { return String(r.currency || '').trim().toUpperCase(); })
    .filter(function (c) { return /^[A-Z]{3}$/.test(c); });
  if (list.indexOf('AUD') < 0) list.unshift('AUD');
  return list.filter(function (c, i) { return list.indexOf(c) === i; });
}

var HEADERS = {};
HEADERS[SHEET_PEOPLE]      = ['name'];
// project_id is appended, not inserted, so an existing sheet keeps its data.
HEADERS[SHEET_EXPENSES]    = ['id', 'date', 'city', 'description', 'paid_by', 'currency',
                              'amount_local', 'amount_aud', 'split_type',
                              'split_detail_json', 'item_breakdown_json', 'project_id'];
HEADERS[SHEET_RATES]       = ['currency', 'rate_to_aud', 'last_updated', 'is_manual_override'];
HEADERS[SHEET_SETTLEMENTS] = ['id', 'from', 'to', 'amount_aud', 'date', 'note', 'project_id'];
HEADERS[SHEET_PROJECTS]    = ['id', 'name', 'start_date', 'end_date',
                              'currencies', 'cities', 'members', 'archived'];

/* ------------------------------------------------------------------ *
 * One-time setup
 * ------------------------------------------------------------------ */

function setup() {
  var ss = book();

  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    var hdr = HEADERS[name];
    // A sheet trimmed to exactly its old column count has nowhere to put
    // project_id, and every later getRange would throw. Widen it first.
    var need = hdr.length - sh.getMaxColumns();
    if (need > 0) sh.insertColumnsAfter(sh.getMaxColumns(), need);
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });

  // Remove the default empty "Sheet1" if it is still pristine.
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  // Seed people only if empty.
  var people = ss.getSheetByName(SHEET_PEOPLE);
  if (people.getLastRow() < 2) {
    people.getRange(2, 1, 3, 1).setValues([['Victor'], ['Person 2'], ['Person 3']]);
  }

  // Seed rate rows only if empty.
  var rates = ss.getSheetByName(SHEET_RATES);
  if (rates.getLastRow() < 2) {
    rates.getRange(2, 1, DEFAULT_CURRENCIES.length, 4).setValues(DEFAULT_CURRENCIES.map(function (c) {
      return [c, c === 'AUD' ? 1 : '', '', false];
    }));
  }

  // Seed the two starting projects only if there are none.
  var projects = ss.getSheetByName(SHEET_PROJECTS);
  if (projects.getLastRow() < 2) {
    projects.getRange(2, 1, 2, 8).setValues([
      ['p_trip', 'Greater China 2026', '2026-10-16', '2026-11-01',
       'AUD,HKD,MOP,CNY', 'Hong Kong,Macau,China,Sydney', '', false],
      ['p_general', 'General', '', '', 'AUD', '', '', false]
    ]);
  }

  migrateProjectIds();
  mergeChinaCities();

  // Generate an API token if there isn't one.
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('API_TOKEN')) {
    props.setProperty('API_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  }

  var token = props.getProperty('API_TOKEN');

  // The execution log is the one place output always lands, whether or not the
  // Sheet's UI is available to show a dialog.
  Logger.log('setup() finished on "' + ss.getName() + '"');
  Logger.log('Projects: ' + getProjects().map(function (p) { return p.name; }).join(', '));
  Logger.log('Currencies: ' + CURRENCIES().join(', '));
  Logger.log('Tabs ready: ' + Object.keys(HEADERS).join(', '));
  Logger.log('People: ' + getPeople().join(', '));
  Logger.log('');
  Logger.log('API TOKEN: ' + token);
  Logger.log('');
  Logger.log('Next: Deploy > New deployment > Web app, Execute as Me, Who has access Anyone.');

  showToken();
  return 'setup complete — token: ' + token;
}

function showToken() {
  var t = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  Logger.log('API TOKEN: ' + t);
  try {
    SpreadsheetApp.getUi().alert('API token\n\n' + t +
      '\n\nPaste this into the app Setup screen together with your /exec URL.' +
      '\n\nIt is also in the execution log at the bottom of the editor.');
  } catch (err) { /* no UI when run headless — the log still has it */ }
  return t;
}

function resetToken() {
  var t = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_TOKEN', t);
  return showToken();
}

/* ------------------------------------------------------------------ *
 * Finding the Sheet
 *
 * Normally this script is bound to the Sheet (created via Extensions >
 * Apps Script), and getActiveSpreadsheet() just works. If it was created
 * standalone at script.google.com instead, there is no active spreadsheet,
 * so fall back to an id stored by useSheet().
 * ------------------------------------------------------------------ */

function book() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      throw new Error('SHEET_ID is set to "' + id + '" but that Sheet could not be opened. ' +
        'Check the id and re-run useSheet("<id>").');
    }
  }

  throw new Error(
    'This script is not attached to a Sheet.\n\n' +
    'Easiest fix: open your Google Sheet, choose Extensions > Apps Script, and paste ' +
    'this code there instead.\n\n' +
    'Or, to keep this standalone script, run useSheet("<sheet id>") once. The id is the ' +
    'long string in the Sheet URL between /d/ and /edit.');
}

/**
 * Point a standalone script at a Sheet. Only needed if this script was NOT
 * created from inside the Sheet via Extensions > Apps Script.
 */
function useSheet(id) {
  id = String(id || '').trim();
  // Accept a full URL as well as a bare id.
  var m = id.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) id = m[1];
  if (!id) throw new Error('Pass the Sheet id, e.g. useSheet("1AbC...xyz")');

  var name;
  try {
    name = SpreadsheetApp.openById(id).getName();
  } catch (err) {
    throw new Error('Could not open a Sheet with id "' + id + '". ' +
      'Make sure it is the id from the Sheet URL between /d/ and /edit, and that this ' +
      'account can open it.');
  }
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', id);
  Logger.log('SHEET_ID set to ' + id + ' ("' + name + '"). Now run setup().');
  return name;
}

/* ------------------------------------------------------------------ *
 * HTTP entry points
 * ------------------------------------------------------------------ */

function doGet(e) {
  return handle(e, (e && e.parameter) ? e.parameter : {});
}

function doPost(e) {
  var body = {};
  try {
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'Bad JSON body' });
  }
  return handle(e, body);
}

function handle(e, req) {
  var expected = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (!expected) return json({ ok: false, error: 'Backend not set up — run setup() in the script editor.' });
  if (String(req.token || '') !== expected) return json({ ok: false, error: 'Bad token' });

  var action  = String(req.action || 'bootstrap');
  var payload = req.payload || {};

  // The Claude calls take tens of seconds and touch no rows, so they must not
  // hold the script lock — that would block every other write for the duration.
  if (action === 'aiStatus' || action === 'parseReceipt' || action === 'parseText') {
    try {
      var aiOut;
      if (action === 'aiStatus')          aiOut = { ai: aiEnabled(), provider: aiProvider() };
      else if (action === 'parseReceipt') aiOut = { draft: parseReceipt(payload) };
      else                                aiOut = { draft: parseText(payload) };
      aiOut.ok = true;
      return json(aiOut);
    } catch (err) {
      return json({ ok: false, error: String(err && err.message ? err.message : err) });
    }
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) {
    return json({ ok: false, error: 'Busy, try again' });
  }
  try {
    var out;

    switch (action) {
      case 'bootstrap':    out = bootstrap(!!req.refreshRates); break;
      case 'addExpense':   out = { expense: addExpense(payload) }; break;
      case 'updateExpense':out = { expense: updateExpense(payload) }; break;
      case 'deleteExpense':out = { id: deleteExpense(payload.id) }; break;
      case 'setPeople':    out = { people: setPeople(payload.people) }; break;
      case 'setRate':      out = { rates: setRate(payload.currency, payload.rate_to_aud) }; break;
      case 'addCurrency':  out = { rates: addCurrency(payload.currency) }; break;
      case 'removeCurrency': out = { rates: removeCurrency(payload.currency) }; break;
      case 'clearOverride':out = { rates: clearOverride(payload.currency) }; break;
      case 'refreshRates': out = { rates: refreshRates(true) }; break;
      case 'addProject':   out = { project: addProject(payload) }; break;
      case 'updateProject':out = { project: updateProject(payload) }; break;
      case 'deleteProject':out = { id: deleteProject(payload.id) }; break;
      case 'addSettlement':out = { settlement: addSettlement(payload) }; break;
      case 'deleteSettlement': out = { id: deleteSettlement(payload.id) }; break;
      default: return json({ ok: false, error: 'Unknown action: ' + action });
    }
    out.ok = true;
    return json(out);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ *
 * Sheet helpers
 * ------------------------------------------------------------------ */

function sheet(name) {
  var sh = book().getSheetByName(name);
  if (!sh) throw new Error('Missing tab "' + name + '" — run setup() in the script editor.');
  return sh;
}

function readRows(name) {
  var sh = sheet(name);
  var last = sh.getLastRow();
  var hdr = HEADERS[name];
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, hdr.length).getValues();
  return values
    .filter(function (r) { return r.join('') !== ''; })
    .map(function (r) {
      var o = {};
      hdr.forEach(function (h, i) { o[h] = r[i]; });
      return o;
    });
}

function findRowById(name, id) {
  var sh = sheet(name);
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function ymd(d) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d),
    Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ------------------------------------------------------------------ *
 * Reads
 * ------------------------------------------------------------------ */

function bootstrap(forceRates) {
  return {
    projects:    getProjects(),
    people:      getPeople(),
    expenses:    getExpenses(),
    rates:       refreshRates(!!forceRates),
    settlements: getSettlements(),
    serverDate:  ymd(new Date())
  };
}

function getPeople() {
  return readRows(SHEET_PEOPLE)
    .map(function (r) { return String(r.name).trim(); })
    .filter(function (n) { return n; });
}

function getExpenses() {
  return readRows(SHEET_EXPENSES).map(function (r) {
    return {
      id:           String(r.id),
      date:         r.date instanceof Date ? ymd(r.date) : String(r.date),
      city:         String(r.city),
      description:  String(r.description),
      paid_by:      String(r.paid_by),
      currency:     String(r.currency),
      amount_local: Number(r.amount_local) || 0,
      amount_aud:   Number(r.amount_aud) || 0,
      split_type:   String(r.split_type || 'equal'),
      split_detail: parseJson(r.split_detail_json, {}),
      items:        parseJson(r.item_breakdown_json, null),
      project_id:   String(r.project_id || '')
    };
  });
}

function getSettlements() {
  return readRows(SHEET_SETTLEMENTS).map(function (r) {
    return {
      id:         String(r.id),
      from:       String(r.from),
      to:         String(r.to),
      amount_aud: Number(r.amount_aud) || 0,
      date:       r.date instanceof Date ? ymd(r.date) : String(r.date),
      note:       String(r.note || ''),
      project_id: String(r.project_id || '')
    };
  });
}

function parseJson(s, fallback) {
  if (s === '' || s === null || s === undefined) return fallback;
  try { return JSON.parse(String(s)); } catch (e) { return fallback; }
}

/* ------------------------------------------------------------------ *
 * Projects
 *
 * A project is a group of expenses kept apart from the others: a trip with a
 * date range, or an open-ended everyday one. Balances and settle-up are always
 * computed within a project, never across them.
 * ------------------------------------------------------------------ */

function csv(v) {
  return String(v == null ? '' : v).split(',')
    .map(function (x) { return x.trim(); })
    .filter(function (x) { return x; });
}

function getProjects() {
  return readRows(SHEET_PROJECTS).map(function (r) {
    return {
      id:         String(r.id),
      name:       String(r.name),
      start_date: r.start_date instanceof Date ? ymd(r.start_date) : String(r.start_date || ''),
      end_date:   r.end_date   instanceof Date ? ymd(r.end_date)   : String(r.end_date   || ''),
      currencies: csv(r.currencies).length ? csv(r.currencies) : ['AUD'],
      cities:     csv(r.cities),
      members:    csv(r.members),
      archived:   r.archived === true || String(r.archived).toLowerCase() === 'true'
    };
  }).filter(function (p) { return p.id; });
}

function projectRow(p) {
  return [
    String(p.id || ('pr' + Date.now() + Math.floor(Math.random() * 1000))),
    String(p.name || 'Untitled'),
    String(p.start_date || ''),
    String(p.end_date || ''),
    (p.currencies && p.currencies.length ? p.currencies : ['AUD']).join(','),
    (p.cities || []).join(','),
    (p.members || []).join(','),
    !!p.archived
  ];
}

function addProject(p) {
  var row = projectRow(p);
  sheet(SHEET_PROJECTS).appendRow(row);
  return getProjects().filter(function (x) { return x.id === row[0]; })[0];
}

function updateProject(p) {
  if (!p.id) throw new Error('updateProject needs an id');
  var r = findRowById(SHEET_PROJECTS, p.id);
  if (r < 0) throw new Error('Project not found: ' + p.id);
  var row = projectRow(p);
  sheet(SHEET_PROJECTS).getRange(r, 1, 1, row.length).setValues([row]);
  return getProjects().filter(function (x) { return x.id === p.id; })[0];
}

/** Deleting a project takes its expenses and settlements with it. */
function deleteProject(id) {
  var all = getProjects();
  if (all.length < 2) throw new Error('Keep at least one project.');
  var r = findRowById(SHEET_PROJECTS, id);
  if (r < 0) throw new Error('Project not found: ' + id);

  [SHEET_EXPENSES, SHEET_SETTLEMENTS].forEach(function (name) {
    var sh = sheet(name);
    var last = sh.getLastRow();
    if (last < 2) return;
    var col = HEADERS[name].indexOf('project_id') + 1;
    var vals = sh.getRange(2, col, last - 1, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {         // bottom-up so indices hold
      if (String(vals[i][0]) === String(id)) sh.deleteRow(i + 2);
    }
  });

  sheet(SHEET_PROJECTS).deleteRow(r);
  return id;
}

/**
 * The three mainland cities all spend CNY, so they are one place now. Renaming
 * them is cosmetic — the currency and every amount are untouched.
 */
function mergeChinaCities() {
  var sh = sheet(SHEET_EXPENSES);
  var last = sh.getLastRow();
  if (last < 2) return 0;

  var col = HEADERS[SHEET_EXPENSES].indexOf('city') + 1;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  var merge = { 'Guangzhou': 1, 'Chongqing': 1, 'Chengdu': 1 };
  var n = 0;

  for (var i = 0; i < vals.length; i++) {
    if (merge[String(vals[i][0]).trim()]) { vals[i][0] = 'China'; n++; }
  }
  if (n) {
    sh.getRange(2, col, vals.length, 1).setValues(vals);
    Logger.log('Merged ' + n + ' expense(s) from Guangzhou/Chongqing/Chengdu into "China".');
  }

  // Bring any project still listing the old cities into line.
  getProjects().forEach(function (p) {
    if (!p.cities.length) return;
    var next = [], changed = false;
    p.cities.forEach(function (c) {
      var name = merge[c] ? 'China' : c;
      if (merge[c]) changed = true;
      if (next.indexOf(name) < 0) next.push(name);
    });
    if (changed) { p.cities = next; updateProject(p); }
  });
  return n;
}

/**
 * Give every row without a project_id one, choosing by date where a project
 * has a range and falling back to the first project otherwise. Safe to re-run.
 */
function migrateProjectIds() {
  var projects = getProjects();
  if (!projects.length) return 0;

  var dated = projects.filter(function (p) { return p.start_date && p.end_date; });
  var fallback = projects.filter(function (p) { return !p.start_date && !p.end_date; })[0] || projects[0];

  function pick(dateStr) {
    for (var i = 0; i < dated.length; i++) {
      if (dateStr >= dated[i].start_date && dateStr <= dated[i].end_date) return dated[i].id;
    }
    return fallback.id;
  }

  var moved = 0;
  [SHEET_EXPENSES, SHEET_SETTLEMENTS].forEach(function (name) {
    var sh = sheet(name);
    var last = sh.getLastRow();
    if (last < 2) return;
    var hdr = HEADERS[name];
    var pCol = hdr.indexOf('project_id') + 1;
    var dCol = hdr.indexOf('date') + 1;

    var pids  = sh.getRange(2, pCol, last - 1, 1).getValues();
    var dates = sh.getRange(2, dCol, last - 1, 1).getValues();
    var out = [], changed = false;

    for (var i = 0; i < pids.length; i++) {
      if (String(pids[i][0] || '').trim()) { out.push([pids[i][0]]); continue; }
      var d = dates[i][0];
      d = (d instanceof Date) ? ymd(d) : String(d || '');
      out.push([pick(d)]);
      changed = true; moved++;
    }
    if (changed) sh.getRange(2, pCol, out.length, 1).setValues(out);
  });

  if (moved) Logger.log('Assigned ' + moved + ' existing row(s) to a project.');
  return moved;
}

/* ------------------------------------------------------------------ *
 * Writes — people
 * ------------------------------------------------------------------ */

function setPeople(names) {
  names = (names || []).map(function (n) { return String(n).trim(); })
                       .filter(function (n) { return n; });
  if (!names.length) throw new Error('Need at least one person');
  var sh = sheet(SHEET_PEOPLE);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
  sh.getRange(2, 1, names.length, 1).setValues(names.map(function (n) { return [n]; }));
  return names;
}

/* ------------------------------------------------------------------ *
 * Writes — expenses
 * ------------------------------------------------------------------ */

function expenseRow(p) {
  var id = String(p.id || ('e' + Date.now() + Math.floor(Math.random() * 1000)));
  return [
    id,
    String(p.date || ymd(new Date())),
    String(p.city || ''),
    String(p.description || ''),
    String(p.paid_by || ''),
    String(p.currency || 'AUD'),
    Number(p.amount_local) || 0,
    Number(p.amount_aud) || 0,
    String(p.split_type || 'equal'),
    JSON.stringify(p.split_detail || {}),
    p.items ? JSON.stringify(p.items) : '',
    String(p.project_id || '')
  ];
}

function rowToExpense(row) {
  var o = {};
  HEADERS[SHEET_EXPENSES].forEach(function (h, i) { o[h] = row[i]; });
  return {
    id: o.id, date: o.date, city: o.city, description: o.description,
    paid_by: o.paid_by, currency: o.currency,
    amount_local: Number(o.amount_local), amount_aud: Number(o.amount_aud),
    split_type: o.split_type,
    split_detail: parseJson(o.split_detail_json, {}),
    items: parseJson(o.item_breakdown_json, null),
    project_id: String(o.project_id || '')
  };
}

function addExpense(p) {
  var row = expenseRow(p);
  sheet(SHEET_EXPENSES).appendRow(row);
  return rowToExpense(row);
}

function updateExpense(p) {
  if (!p.id) throw new Error('updateExpense needs an id');
  var r = findRowById(SHEET_EXPENSES, p.id);
  if (r < 0) throw new Error('Expense not found: ' + p.id);
  var row = expenseRow(p);
  sheet(SHEET_EXPENSES).getRange(r, 1, 1, row.length).setValues([row]);
  return rowToExpense(row);
}

function deleteExpense(id) {
  var r = findRowById(SHEET_EXPENSES, id);
  if (r < 0) throw new Error('Expense not found: ' + id);
  sheet(SHEET_EXPENSES).deleteRow(r);
  return id;
}

/* ------------------------------------------------------------------ *
 * Writes — settlements
 * ------------------------------------------------------------------ */

function addSettlement(p) {
  var row = [
    String(p.id || ('s' + Date.now() + Math.floor(Math.random() * 1000))),
    String(p.from || ''),
    String(p.to || ''),
    Number(p.amount_aud) || 0,
    String(p.date || ymd(new Date())),
    String(p.note || ''),
    String(p.project_id || '')
  ];
  sheet(SHEET_SETTLEMENTS).appendRow(row);
  return { id: row[0], from: row[1], to: row[2], amount_aud: row[3],
           date: row[4], note: row[5], project_id: row[6] };
}

function deleteSettlement(id) {
  var r = findRowById(SHEET_SETTLEMENTS, id);
  if (r < 0) throw new Error('Settlement not found: ' + id);
  sheet(SHEET_SETTLEMENTS).deleteRow(r);
  return id;
}

/* ------------------------------------------------------------------ *
 * Exchange rates
 *
 * rate_to_aud = how many AUD one unit of that currency is worth.
 * Manual overrides are never touched by the daily refresh.
 * ------------------------------------------------------------------ */

function readRates() {
  var rows = readRows(SHEET_RATES);
  var map = {};
  rows.forEach(function (r) {
    var cur = String(r.currency).trim().toUpperCase();
    if (!cur) return;
    map[cur] = {
      currency: cur,
      rate_to_aud: Number(r.rate_to_aud) || 0,
      last_updated: r.last_updated instanceof Date
        ? Utilities.formatDate(r.last_updated, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss")
        : String(r.last_updated || ''),
      is_manual_override: r.is_manual_override === true ||
                          String(r.is_manual_override).toLowerCase() === 'true'
    };
  });
  return map;
}

function writeRates(map) {
  // Read the currency list BEFORE clearing — CURRENCIES() reads this same tab,
  // and clearing first would leave nothing to rebuild from.
  var list = CURRENCIES();
  var sh = sheet(SHEET_RATES);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  var rows = list.map(function (c) {
    var r = map[c] || { rate_to_aud: c === 'AUD' ? 1 : 0, last_updated: '', is_manual_override: false };
    return [c, r.rate_to_aud, r.last_updated, !!r.is_manual_override];
  });
  sh.getRange(2, 1, rows.length, 4).setValues(rows);
  return ratesOut(rows);
}

function ratesOut(rows) {
  return rows.map(function (r) {
    return {
      currency: r[0],
      rate_to_aud: Number(r[1]) || 0,
      last_updated: String(r[2] || ''),
      is_manual_override: !!r[3]
    };
  });
}

function asList(map) {
  return CURRENCIES().map(function (c) {
    var r = map[c] || {};
    return {
      currency: c,
      rate_to_aud: Number(r.rate_to_aud) || (c === 'AUD' ? 1 : 0),
      last_updated: String(r.last_updated || ''),
      is_manual_override: !!r.is_manual_override
    };
  });
}

/**
 * Returns the current rate list, fetching fresh rates at most once a day
 * unless force is true. Manual overrides are preserved.
 */
function refreshRates(force) {
  var map = readRates();
  var today = ymd(new Date());

  map.AUD = { currency: 'AUD', rate_to_aud: 1, last_updated: today, is_manual_override: false };

  var stale = CURRENCIES().some(function (c) {
    if (c === 'AUD') return false;
    var r = map[c];
    if (!r || !r.rate_to_aud) return true;
    if (r.is_manual_override) return false;
    return String(r.last_updated || '').slice(0, 10) !== today;
  });

  if (!force && !stale) return asList(map);

  var live = fetchLiveRates(CURRENCIES());
  if (live) {
    CURRENCIES().forEach(function (c) {
      if (c === 'AUD') return;
      if (map[c] && map[c].is_manual_override) return;   // overrides persist until cleared
      if (!live[c]) return;
      map[c] = {
        currency: c,
        rate_to_aud: live[c],
        last_updated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
        is_manual_override: false
      };
    });
  }

  // Last-resort seed so the app is never unusable with a zero rate.
  // Only the original four have a sane offline guess; a currency added later
  // stays at zero and the app shows "no rate" rather than inventing one.
  var FALLBACK = { HKD: 0.19, MOP: 0.185, CNY: 0.21 };
  CURRENCIES().forEach(function (c) {
    if (c === 'AUD') return;
    if ((!map[c] || !map[c].rate_to_aud) && FALLBACK[c]) {
      map[c] = { currency: c, rate_to_aud: FALLBACK[c], last_updated: '', is_manual_override: false };
    }
  });

  var rows = CURRENCIES().map(function (c) {
    return [c, map[c].rate_to_aud, map[c].last_updated, !!map[c].is_manual_override];
  });
  var sh = sheet(SHEET_RATES);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  sh.getRange(2, 1, rows.length, 4).setValues(rows);

  return ratesOut(rows);
}

/**
 * open.er-api.com is the primary source (it covers MOP, which the ECB feed
 * behind Frankfurter does not). Frankfurter is the fallback for HKD/CNY, and
 * MOP is then derived from its de-facto 1.03 MOP : 1 HKD peg.
 * Returns { HKD: audPerUnit, MOP: ..., CNY: ... } or null.
 */
function fetchLiveRates(wanted) {
  var want = (wanted && wanted.length) ? wanted : CURRENCIES();
  want = want.filter(function (c) { return c !== 'AUD'; });
  var out = null;

  try {
    var res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/AUD',
      { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() === 200) {
      var d = JSON.parse(res.getContentText());
      if (d && d.rates) {
        out = {};
        want.forEach(function (c) {
          if (d.rates[c]) out[c] = 1 / Number(d.rates[c]);   // AUD per 1 unit
        });
        var missing = want.filter(function (c) { return !out[c]; });
        if (!missing.length) return out;
      }
    }
  } catch (err) { /* fall through */ }

  try {
    var res2 = UrlFetchApp.fetch(
      'https://api.frankfurter.app/latest?from=AUD&to=' + encodeURIComponent(want.join(',')),
      { muteHttpExceptions: true, followRedirects: true });
    if (res2.getResponseCode() === 200) {
      var d2 = JSON.parse(res2.getContentText());
      if (d2 && d2.rates) {
        out = out || {};
        want.forEach(function (c) {
          if (!out[c] && d2.rates[c]) out[c] = 1 / Number(d2.rates[c]);
        });
      }
    }
  } catch (err2) { /* fall through */ }

  // Frankfurter's ECB feed has no MOP, which is pegged to the HKD.
  if (out && !out.MOP && out.HKD && want.indexOf('MOP') >= 0) out.MOP = out.HKD / 1.03;

  return out;
}

/**
 * Add a currency. The live-rate API is asked for it first, so a typo or a code
 * nobody quotes is rejected here rather than silently producing a dead row.
 */
function addCurrency(code) {
  var cur = String(code || '').trim().toUpperCase();
  if (!/^[A-Z]{3}$/.test(cur)) throw new Error('A currency code is three letters, like SGD or JPY.');
  if (CURRENCIES().indexOf(cur) >= 0) throw new Error(cur + ' is already there.');

  var live = fetchLiveRates([cur]);
  if (!live || !live[cur]) {
    throw new Error('No live rate available for ' + cur + '. Check the code, or add it ' +
                    'manually by typing a row into the ExchangeRates tab.');
  }
  sheet(SHEET_RATES).appendRow([
    cur, live[cur],
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
    false
  ]);
  return refreshRates(false);
}

/** Remove a currency. Expenses already recorded in it keep their AUD amounts. */
function removeCurrency(code) {
  var cur = String(code || '').trim().toUpperCase();
  if (cur === 'AUD') throw new Error('AUD is the base currency and cannot be removed.');
  var r = findRowById(SHEET_RATES, cur);
  if (r < 0) throw new Error(cur + ' is not in the list.');
  sheet(SHEET_RATES).deleteRow(r);
  return refreshRates(false);
}

function setRate(currency, rate) {
  var cur = String(currency).trim().toUpperCase();
  if (CURRENCIES().indexOf(cur) < 0) throw new Error('Unsupported currency: ' + cur);
  if (cur === 'AUD') throw new Error('AUD is the base currency and is always 1');
  var val = Number(rate);
  if (!(val > 0)) throw new Error('Rate must be greater than 0');

  var map = readRates();
  map[cur] = {
    currency: cur,
    rate_to_aud: val,
    last_updated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
    is_manual_override: true
  };
  return writeRates(map);
}

function clearOverride(currency) {
  var cur = String(currency).trim().toUpperCase();
  var map = readRates();
  if (map[cur]) {
    map[cur].is_manual_override = false;
    map[cur].last_updated = '';          // force the next refresh to re-fetch
    writeRates(map);
  }
  return refreshRates(false);
}

/* ------------------------------------------------------------------ *
 * Claude — receipt photos and plain-English descriptions
 *
 * The API key lives in this script's properties, never on the phone and
 * never in the repo. Set it once with setApiKey('sk-ant-...').
 * ------------------------------------------------------------------ */

var CLAUDE_MODEL = 'claude-opus-5';
var CLAUDE_URL   = 'https://api.anthropic.com/v1/messages';

// Gemini's free tier needs no card. Model names drift, so this is only the first
// guess — findGeminiModel() asks Google what actually exists if it 404s.
var GEMINI_MODELS = ['gemini-3.6-flash', 'gemini-flash-latest', 'gemini-2.5-flash', 'gemini-2.0-flash'];
var GEMINI_BASE   = 'https://generativelanguage.googleapis.com/v1beta';

/** Anthropic — paid, best receipt reading. */
function setApiKey(key) {
  key = String(key || '').trim();
  if (!key) throw new Error('Pass your Anthropic API key, e.g. setApiKey("sk-ant-...")');
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', key);
  Logger.log('Anthropic key saved. Receipt scanning and describe-it are on, using Claude.');
  return 'ok';
}

function clearApiKey() {
  PropertiesService.getScriptProperties().deleteProperty('ANTHROPIC_API_KEY');
  Logger.log('Anthropic key removed.');
  return 'ok';
}

/** Google AI Studio — free tier, no card. Get one at aistudio.google.com/apikey */
function setGeminiKey(key) {
  key = String(key || '').trim();
  if (!key) throw new Error('Pass your Google AI Studio key, e.g. setGeminiKey("AIza...")');
  var props = PropertiesService.getScriptProperties();
  props.setProperty('GEMINI_API_KEY', key);
  props.deleteProperty('GEMINI_MODEL');   // re-discover on next use
  Logger.log('Gemini key saved. Receipt scanning and describe-it are on, using Gemini (free tier).');
  return 'ok';
}

function clearGeminiKey() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('GEMINI_API_KEY');
  props.deleteProperty('GEMINI_MODEL');
  Logger.log('Gemini key removed.');
  return 'ok';
}

/**
 * Which provider to use. Claude wins if both keys are set, because it reads
 * receipts better; delete the Anthropic key to fall back to the free tier.
 */
function aiProvider() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('ANTHROPIC_API_KEY')) return 'claude';
  if (props.getProperty('GEMINI_API_KEY'))    return 'gemini';
  return '';
}

function aiEnabled() { return !!aiProvider(); }

/** Run this in the editor to check a key works, without photographing anything. */
function testAi() {
  var who = aiProvider();
  if (!who) {
    Logger.log('No AI key set. Run setGeminiKey("AIza...") for the free option, ' +
               'or setApiKey("sk-ant-...") for Claude.');
    return 'no key';
  }
  Logger.log('Provider: ' + who);
  var draft = parseText({
    text: 'Dinner in Chengdu, 396 yuan, Victor paid, split evenly',
    city: 'Chengdu', currency: 'CNY'
  });
  Logger.log('Parsed OK: ' + JSON.stringify(draft, null, 2));
  return draft;
}

/** The shape every parse returns. All fields required so the schema stays strict. */
function draftSchema(people, cities, currencies) {
  cities     = (cities && cities.length) ? cities : CITY_NAMES;
  currencies = (currencies && currencies.length) ? currencies : CURRENCIES();
  return {
    type: 'object',
    additionalProperties: false,
    required: ['description', 'date', 'city', 'currency', 'amount_local', 'paid_by',
               'split_type', 'participants', 'shares', 'items', 'extra', 'note'],
    properties: {
      description:  { type: 'string', description: 'Short label, e.g. "Hotpot dinner". Never blank.' },
      date:         { type: 'string', description: 'YYYY-MM-DD. Use the date on the receipt if legible, otherwise today.' },
      city:         { type: 'string', enum: cities, description: 'Which trip city this was in. Keep the one given in the prompt unless the receipt or description clearly says otherwise.' },
      currency:     { type: 'string', enum: currencies },
      amount_local: { type: 'number', description: 'Grand total in the local currency. Never converted to AUD.' },
      paid_by:      { type: 'string', enum: people, description: 'Who paid. If unstated, the default payer given in the prompt.' },
      split_type:   { type: 'string', enum: ['equal', 'exact', 'percent', 'itemized'] },
      participants: {
        type: 'array', items: { type: 'string', enum: people },
        description: 'For split_type "equal": who shares it. Empty array otherwise.'
      },
      shares: {
        type: 'array',
        description: 'For "exact": each person’s amount in local currency. For "percent": each person’s percentage. Empty array otherwise.',
        items: {
          type: 'object', additionalProperties: false,
          required: ['person', 'value'],
          properties: {
            person: { type: 'string', enum: people },
            value:  { type: 'number' }
          }
        }
      },
      items: {
        type: 'array',
        description: 'For "itemized": one entry per line item. Empty array otherwise.',
        items: {
          type: 'object', additionalProperties: false,
          required: ['name', 'amount', 'people'],
          properties: {
            name:   { type: 'string' },
            amount: { type: 'number', description: 'Line total in local currency, including quantity.' },
            people: { type: 'array', items: { type: 'string', enum: people } }
          }
        }
      },
      extra: { type: 'number', description: 'Tax, service charge and tip combined, in local currency. 0 if none.' },
      note:  { type: 'string', description: 'One short sentence on anything uncertain, for the human to check. Empty if nothing.' }
    }
  };
}

function draftSystemPrompt(people, city, currency, defaultPayer, todayStr, cities, currencies) {
  return [
    'You turn receipts and plain-English descriptions into a single expense for a',
    'three-person trip splitter. The trip runs 16 Oct - 1 Nov 2026 across Hong Kong,',
    'Macau, Guangzhou, Chongqing and Chengdu, with Sydney as home base.',
    '',
    'The three people are: ' + people.join(', ') + '. Use these names exactly.',
    'Currencies allowed here: ' + currencies.join(', ') + '.',
    (cities.length ? 'Places allowed here: ' + cities.join(', ') + '.' : ''),
    'City currencies: Hong Kong=HKD, Macau=MOP, Guangzhou/Chongqing/Chengdu=CNY, Sydney=AUD.',
    '',
    'Context for this expense:',
    '- City: ' + (city || 'unknown'),
    '- Expected currency: ' + (currency || 'unknown'),
    '- Default payer if not stated: ' + defaultPayer,
    "- Today's date: " + todayStr,
    '',
    'Rules:',
    '1. Report amounts in the local currency exactly as written. Never convert to AUD;',
    '   the app does that itself.',
    '2. amount_local is the grand total actually paid, including tax and service.',
    '3. Put tax, service charge and tip in "extra" — never as a line item.',
    '4. Use "itemized" only when you can read the line items AND you know who had what.',
    '   If the split is not stated, prefer "equal" across all three people.',
    '5. Never invent an amount you cannot read. If the total is unclear, put your best',
    '   reading in amount_local and say so in "note".',
    '6. For "exact", the shares must add up to amount_local. For "percent", to 100.',
    '7. Keep "description" short and concrete — what was bought, not a sentence.',
    '8. Use "note" only for genuine uncertainty. Leave it empty when the reading is clean.'
  ].join('\n');
}

/**
 * One Claude call returning a draft expense. Raw HTTP because Apps Script has
 * no Anthropic SDK. Retries once without the fallback beta if the API rejects it,
 * so an unfamiliar beta flag can never take the feature down.
 */
/** Provider-agnostic entry point. Both paths return the same normalised draft. */
function aiDraft(parts, people, city, currency, defaultPayer, cities, currencies) {
  cities     = (cities && cities.length) ? cities : CITY_NAMES;
  currencies = (currencies && currencies.length) ? currencies : CURRENCIES();
  var who = aiProvider();
  if (!who) {
    throw new Error('AI features are off. Run setGeminiKey("AIza...") for the free option, ' +
                    'or setApiKey("sk-ant-...") for Claude.');
  }
  var system = draftSystemPrompt(people, city, currency, defaultPayer, ymd(new Date()), cities, currencies);
  var raw = (who === 'gemini')
    ? geminiCall(system, parts, people, cities, currencies)
    : claudeCall(system, parts, people, cities, currencies);

  var draft;
  try {
    draft = JSON.parse(raw);
  } catch (err) {
    throw new Error('Could not read the model\u2019s reply as JSON. Try again, or enter it by hand.');
  }
  return normaliseDraft(draft, people, currency, city, cities, currencies);
}

/**
 * `parts` is a neutral list: {text} and/or {image: base64, mediaType}.
 * Each provider renders it into its own content shape.
 */
function claudeContent(parts) {
  return parts.map(function (p) {
    return p.image
      ? { type: 'image', source: { type: 'base64', media_type: p.mediaType, data: p.image } }
      : { type: 'text', text: p.text };
  });
}

function claudeCall(system, parts, people, cities, currencies) {
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  var userContent = claudeContent(parts);
  var body = {
    model: CLAUDE_MODEL,
    max_tokens: 4000,
    system: draftSystemPrompt(people, city, currency, defaultPayer, todayStr),
    messages: [{ role: 'user', content: userContent }],
    output_config: {
      // Low effort keeps the round trip inside Apps Script's fetch timeout;
      // this is extraction, not reasoning-heavy work.
      effort: 'low',
      format: { type: 'json_schema', schema: draftSchema(people, cities, currencies) }
    },
    fallbacks: 'default'
  };

  var res = retryFetch(function () { return claudeFetch(key, body, true); });
  if (res.getResponseCode() === 400 && /fallback|beta/i.test(res.getContentText())) {
    delete body.fallbacks;
    res = retryFetch(function () { return claudeFetch(key, body, false); });
  }

  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code !== 200) {
    var msg = text;
    try { msg = JSON.parse(text).error.message; } catch (err) {}
    if (code === 401) msg = 'Anthropic rejected the API key. Re-run setApiKey with a valid key.';
    if (code === 429) msg = 'Anthropic rate limit hit. Wait a moment and try again.';
    if (code === 529 || code >= 500) msg = 'Anthropic is overloaded. It already retried with backoff — try again shortly.';
    throw new Error('Claude error ' + code + ': ' + msg);
  }

  var data = JSON.parse(text);

  // stop_details is only populated on a refusal, so guard before reading it.
  if (data.stop_reason === 'refusal') {
    var why = (data.stop_details && data.stop_details.explanation) ? ' (' + data.stop_details.explanation + ')' : '';
    throw new Error('Claude declined to process that image' + why + '. Enter the expense by hand.');
  }

  var out = '';
  (data.content || []).forEach(function (b) { if (b.type === 'text') out += b.text; });
  if (!out) throw new Error('Claude returned nothing to read. Try again, or enter it by hand.');

  return out;
}

function claudeFetch(key, body, withFallback) {
  var headers = {
    'x-api-key': key,
    'anthropic-version': '2023-06-01'
  };
  if (withFallback) headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';

  return UrlFetchApp.fetch(CLAUDE_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
}

/* ------------------------------------------------------------------ *
 * Gemini (Google AI Studio free tier)
 * ------------------------------------------------------------------ */

/**
 * Gemini takes an OpenAPI-subset schema, not JSON Schema: types are upper-case
 * enum names and additionalProperties is not allowed. Convert rather than
 * maintaining a second copy of the schema.
 */
function toGeminiSchema(node) {
  if (!node || typeof node !== 'object') return node;

  var out = {};
  if (node.type) out.type = String(node.type).toUpperCase();
  if (node.description) out.description = node.description;
  if (node.enum) out.enum = node.enum.slice();
  if (node.items) out.items = toGeminiSchema(node.items);

  if (node.properties) {
    out.properties = {};
    Object.keys(node.properties).forEach(function (k) {
      out.properties[k] = toGeminiSchema(node.properties[k]);
    });
    // Fixing the order makes the model fill fields in a sensible sequence.
    out.propertyOrdering = Object.keys(node.properties);
  }
  if (node.required) out.required = node.required.slice();
  return out;   // additionalProperties deliberately dropped
}

function geminiContent(parts) {
  return parts.map(function (p) {
    return p.image
      ? { inline_data: { mime_type: p.mediaType, data: p.image } }
      : { text: p.text };
  });
}

/**
 * Model names move. Try the stored/default one; if Google says it does not
 * exist, ask for the list, pick a flash model that can generateContent, and
 * remember it.
 */
function geminiModel() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || GEMINI_MODELS[0];
}

/** Is this failure about the model itself, rather than the key or the request? */
function isModelProblem(code, text) {
  return code === 404 || (code === 400 && /model|not found|not supported|no longer available/i.test(text));
}

/**
 * Google usually names the replacement right in the error text, e.g.
 * "Please update your code to use models/gemini-3.6-flash". That beats guessing.
 */
function suggestedModel(text, tried) {
  var hits = String(text).match(/models\/[a-zA-Z0-9._\-]+/g) || [];
  for (var i = 0; i < hits.length; i++) {
    var name = hits[i].replace(/^models\//, '');
    if (tried.indexOf(name) < 0) return name;
  }
  return null;
}

/** Ask Google what this key can actually use, skipping anything already tried. */
function findGeminiModel(key, tried) {
  tried = tried || [];
  var list = listGeminiModels(key);
  if (!list.length) return null;

  var usable = list.filter(function (n) { return tried.indexOf(n) < 0; });
  if (!usable.length) return null;

  // Preferred names first, then any plain flash model, then anything at all.
  for (var i = 0; i < GEMINI_MODELS.length; i++) {
    if (usable.indexOf(GEMINI_MODELS[i]) >= 0) return GEMINI_MODELS[i];
  }
  var flash = usable.filter(function (n) {
    return n.indexOf('flash') >= 0 &&
           n.indexOf('thinking') < 0 && n.indexOf('image') < 0 &&
           n.indexOf('tts') < 0 && n.indexOf('embedding') < 0 && n.indexOf('live') < 0;
  });
  return flash.length ? flash[0] : usable[0];
}

/** Run this in the editor to see every model your key can use. */
function listGeminiModels(key) {
  key = key || PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('No Gemini key set.');

  var res = UrlFetchApp.fetch(GEMINI_BASE + '/models?key=' + encodeURIComponent(key) + '&pageSize=200',
    { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    Logger.log('Could not list models: ' + res.getContentText());
    return [];
  }
  var names = (JSON.parse(res.getContentText()).models || []).filter(function (m) {
    return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0;
  }).map(function (m) {
    return String(m.name || '').replace(/^models\//, '');
  });
  Logger.log('Models usable with generateContent:\n  ' + names.join('\n  '));
  return names;
}

/**
 * Retry transient server-side failures (503 overloaded, 5xx, 429) with backoff.
 * Free tiers are the first thing shed when a provider is busy, so a couple of
 * seconds of patience beats handing the user an error they can do nothing with.
 * Anything below 500 comes straight back — a real error, or a 404 for the model
 * ladder to deal with.
 */
function retryFetch(fn, tries) {
  tries = tries || 3;
  var res;
  for (var i = 0; i < tries; i++) {
    res = fn();
    var code = res.getResponseCode();
    if (code < 500 && code !== 429) return res;
    if (i === tries - 1) return res;
    Utilities.sleep(retryDelayMs(res.getContentText(), i));
  }
  return res;
}

/** Honour the provider's own retryDelay when it sends one, else back off 1.5s, 3s, 6s. */
function retryDelayMs(text, attempt) {
  var m = String(text).match(/"retryDelay"\s*:\s*"(\d+(?:\.\d+)?)s"/);
  if (m) return Math.min(Number(m[1]) * 1000 + 250, 20000);
  return Math.min(1500 * Math.pow(2, attempt), 12000);
}

function geminiFetch(key, model, body) {
  return UrlFetchApp.fetch(
    GEMINI_BASE + '/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(key),
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
}

function geminiCall(system, parts, people, cities, currencies) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');

  var body = {
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: 'user', parts: geminiContent(parts) }],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json',
      responseSchema: toGeminiSchema(draftSchema(people, cities, currencies))
    }
  };

  // Model names get retired without notice, so walk candidates rather than
  // failing on the first dead one: the replacement Google names in the error
  // first, then whatever ListModels says this key can actually use.
  var tried = [];
  var model = geminiModel();
  var res, code, text;

  for (var attempt = 0; attempt < 4; attempt++) {
    tried.push(model);
    res  = retryFetch((function (m) { return function () { return geminiFetch(key, m, body); }; })(model));
    code = res.getResponseCode();
    text = res.getContentText();

    if (code === 200) {
      if (props.getProperty('GEMINI_MODEL') !== model) props.setProperty('GEMINI_MODEL', model);
      break;
    }
    // A model that is still overloaded after backing off is worth swapping out of —
    // a less fashionable one is usually free.
    if (!isModelProblem(code, text) && code !== 503) break;

    var next = suggestedModel(text, tried) || findGeminiModel(key, tried);
    if (!next) break;
    model = next;
  }

  if (code !== 200) {
    var msg = text;
    try { msg = JSON.parse(text).error.message; } catch (err) {}
    if (code === 400 && /API key not valid/i.test(msg)) {
      msg = 'Google rejected the key. Re-run setGeminiKey with a valid AI Studio key.';
    }
    if (code === 429) {
      msg = 'Free tier limit reached for now. Wait a few minutes, or enter this one by hand.';
    }
    if (code === 503 || code >= 500) {
      msg = 'Google is overloaded right now and the free tier is the first to get squeezed. ' +
            'It already retried ' + tried.length + ' model(s) with backoff. Try again in a minute, ' +
            'or just type this one in.';
    }
    if (isModelProblem(code, text)) {
      msg = 'No usable Gemini model found. Tried: ' + tried.join(', ') +
            '. Run listGeminiModels() in the script editor to see what this key can use, ' +
            'then set GEMINI_MODEL in Project Settings > Script Properties. Google said: ' + msg;
    }
    throw new Error('Gemini error ' + code + ': ' + msg);
  }

  var data = JSON.parse(text);
  var cand = (data.candidates || [])[0];

  if (!cand) {
    var blocked = data.promptFeedback && data.promptFeedback.blockReason;
    throw new Error(blocked
      ? 'Gemini declined that image (' + blocked + '). Enter the expense by hand.'
      : 'Gemini returned nothing. Try again, or enter it by hand.');
  }
  if (cand.finishReason && cand.finishReason !== 'STOP' && cand.finishReason !== 'MAX_TOKENS') {
    throw new Error('Gemini stopped early (' + cand.finishReason + '). Enter the expense by hand.');
  }

  var out = '';
  (((cand.content || {}).parts) || []).forEach(function (pt) { if (pt.text) out += pt.text; });
  if (!out) throw new Error('Gemini returned nothing to read. Try again, or enter it by hand.');
  return out;
}

/** Trust nothing from the model: clamp names, currency and numbers to what the app allows. */
function normaliseDraft(d, people, fallbackCurrency, fallbackCity, cities, currencies) {
  cities     = (cities && cities.length) ? cities : CITY_NAMES;
  currencies = (currencies && currencies.length) ? currencies : CURRENCIES();
  function person(n) { return people.indexOf(String(n)) >= 0 ? String(n) : null; }
  function num(v) { var x = Number(v); return isFinite(x) && x > 0 ? Math.round(x * 100) / 100 : 0; }

  var cur = String(d.currency || '').toUpperCase();
  if (currencies.indexOf(cur) < 0) cur = fallbackCurrency || currencies[0] || 'AUD';

  var city = String(d.city || '');
  if (cities.indexOf(city) < 0) city = fallbackCity || '';

  var type = ['equal', 'exact', 'percent', 'itemized'].indexOf(String(d.split_type)) >= 0
    ? String(d.split_type) : 'equal';

  var participants = (d.participants || []).map(person).filter(Boolean);

  var shares = {};
  (d.shares || []).forEach(function (s) {
    var p = person(s && s.person);
    if (p && num(s.value)) shares[p] = num(s.value);
  });

  var items = (d.items || []).map(function (i) {
    return {
      name: String((i && i.name) || 'Item').slice(0, 60),
      amount: num(i && i.amount),
      people: ((i && i.people) || []).map(person).filter(Boolean)
    };
  }).filter(function (i) { return i.amount > 0 && i.people.length; });

  // If the model picked a split it did not actually populate, fall back to equal
  // rather than handing the app an expense that splits to nothing.
  if (type === 'itemized' && !items.length) type = 'equal';
  if ((type === 'exact' || type === 'percent') && !Object.keys(shares).length) type = 'equal';
  if (type === 'equal' && !participants.length) participants = people.slice();

  var date = String(d.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) date = ymd(new Date());

  return {
    description:  String(d.description || '').slice(0, 80),
    date:         date,
    city:         city,
    currency:     cur,
    amount_local: num(d.amount_local),
    paid_by:      person(d.paid_by) || people[0],
    split_type:   type,
    participants: participants,
    shares:       shares,
    items:        items,
    extra:        num(d.extra),
    note:         String(d.note || '').slice(0, 200)
  };
}

function parseReceipt(p) {
  var people = getPeople();
  if (!people.length) throw new Error('Add people in Settings first.');

  var b64 = String(p.image || '');
  if (!b64) throw new Error('No image received.');
  var media = String(p.mediaType || 'image/jpeg');
  if (['image/jpeg', 'image/png', 'image/webp', 'image/gif'].indexOf(media) < 0) media = 'image/jpeg';

  var parts = [
    { image: b64, mediaType: media },
    { text: receiptPrompt(p) }
  ];
  return aiDraft(parts, people, p.city, p.currency, p.defaultPayer || people[0], p.cities, p.currencies);
}

function receiptPrompt(p) {
  var t = 'Read this receipt and turn it into one expense.';
  if (p.hint) {
    t += '\n\nThe person entering it added this, which overrides anything ambiguous ' +
         'on the receipt — especially who had what:\n"' + String(p.hint).slice(0, 500) + '"';
  }
  return t;
}

function parseText(p) {
  var people = getPeople();
  if (!people.length) throw new Error('Add people in Settings first.');

  var text = String(p.text || '').trim();
  if (!text) throw new Error('Nothing to read — say or type what the expense was.');

  var parts = [{
    text: 'Turn this description into one expense:\n\n"' + text.slice(0, 2000) + '"'
  }];
  return aiDraft(parts, people, p.city, p.currency, p.defaultPayer || people[0], p.cities, p.currencies);
}
