/**
 * Trip Split — Apps Script backend
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
var SHEET_RECURRING   = 'Recurring';
var SHEET_ACCESS      = 'Access';
var SHEET_CHANGES     = 'Changes';
var SHEET_DELETED     = 'Deleted';

// Quick categories for the spending breakdown. Blank means not chosen yet.
var CATEGORIES = ['food', 'transport', 'stay', 'activities', 'shopping', 'groceries', 'bills', 'other'];
function category(c) { c = String(c || '').toLowerCase(); return CATEGORIES.indexOf(c) >= 0 ? c : ''; }

// Only the starting set. The ExchangeRates tab is the real list from then on,
// so currencies are added from the app without touching this file.
var DEFAULT_CURRENCIES = ['AUD'];

// Bump when the app starts relying on something new here. The app compares it and
// says "redeploy" in plain words instead of failing on an unknown action.
var API_VERSION = 8;

/** Every currency the sheet knows about. AUD is always first and always present. */
function CURRENCIES() {
  var list = readRows(SHEET_RATES)
    .map(function (r) { return String(r.currency || '').trim().toUpperCase(); })
    .filter(function (c) { return /^[A-Z]{3}$/.test(c); });
  if (list.indexOf('AUD') < 0) list.unshift('AUD');
  return list.filter(function (c, i) { return list.indexOf(c) === i; });
}

var HEADERS = {};
// currency is who settles in what: blank means AUD, the shared pool.
HEADERS[SHEET_PEOPLE]      = ['name', 'currency'];
// project_id and category are appended, not inserted, so an existing sheet keeps its data.
HEADERS[SHEET_EXPENSES]    = ['id', 'date', 'city', 'description', 'paid_by', 'currency',
                              'amount_local', 'amount_aud', 'split_type',
                              'split_detail_json', 'item_breakdown_json', 'project_id', 'category',
                              'created_by', 'updated_by', 'updated_at'];
HEADERS[SHEET_RATES]       = ['currency', 'rate_to_aud', 'last_updated', 'is_manual_override'];
HEADERS[SHEET_SETTLEMENTS] = ['id', 'from', 'to', 'amount_aud', 'date', 'note', 'project_id',
                              'currency', 'amount_local', 'created_by', 'created_at'];
HEADERS[SHEET_RECURRING]   = ['id', 'project_id', 'kind', 'frequency', 'start_date', 'end_date',
                              'generated_through', 'paused', 'template_json'];
HEADERS[SHEET_PROJECTS]    = ['id', 'name', 'start_date', 'end_date',
                              'currencies', 'cities', 'members', 'archived'];
// One row per invite link. The token is the link's password; revoked turns it off.
// role: "guest" (only their own expenses in their projects) or "member" (everything, as themselves).
HEADERS[SHEET_ACCESS]      = ['token', 'person', 'project_ids', 'created', 'revoked', 'role'];
// One row per change, never edited: who, when, what, and the before and after.
HEADERS[SHEET_CHANGES]     = ['when', 'who', 'action', 'type', 'id', 'summary', 'before_json', 'after_json'];
// Everything deleted, kept whole so it can be restored. Rows stay; a restore is stamped on them.
HEADERS[SHEET_DELETED]     = ['deleted_at', 'deleted_by', 'type', 'id', 'project_id', 'summary', 'item_json',
                              'batch', 'restored_at', 'restored_by'];

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
    people.getRange(2, 1, 2, 1).setValues([['Person 1'], ['Person 2']]);
  }

  // Seed rate rows only if empty.
  var rates = ss.getSheetByName(SHEET_RATES);
  if (rates.getLastRow() < 2) {
    rates.getRange(2, 1, DEFAULT_CURRENCIES.length, 4).setValues(DEFAULT_CURRENCIES.map(function (c) {
      return [c, c === 'AUD' ? 1 : '', '', false];
    }));
  }

  // Seed one open-ended project only if there are none. Trips are added from the app.
  var projects = ss.getSheetByName(SHEET_PROJECTS);
  if (projects.getLastRow() < 2) {
    projects.getRange(2, 1, 1, 8).setValues([
      ['p_general', 'General', '', '', 'AUD', '', '', false]
    ]);
  }

  migrateProjectIds();

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
  var who = principal(String(req.token || ''), expected);
  if (!who) return json({ ok: false, error: 'Bad token' });
  if (who.revoked) return json({ ok: false, error: 'This invite link has been turned off. Ask for a new one.' });

  var action  = String(req.action || 'bootstrap');
  var payload = req.payload || {};
  ACTOR = who.owner ? (who.person || 'Main link') : who.person + ' (guest)';

  // The Claude calls take tens of seconds and touch no rows, so they must not
  // hold the script lock — that would block every other write for the duration.
  if (action === 'aiStatus' || action === 'parseReceipt' || action === 'parseText') {
    try {
      var aiOut;
      // A guest's draft may only name people they can see.
      var names = who.owner ? null : guestPeople(who, guestProjects(who));
      if (action === 'aiStatus')          aiOut = { ai: aiEnabled(), provider: aiProvider() };
      else if (action === 'parseReceipt') aiOut = { draft: parseReceipt(payload, names) };
      else                                aiOut = { draft: parseText(payload, names) };
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
    var before = CHANGE_TYPES[action] ? changeSnapshot(action, payload) : undefined;

    if (!who.owner) {
      out = guestAction(who, action, payload, req);
      logChange(action, payload, before);
      afterRequest(action);
      out.ok = true;
      return json(out);
    }

    switch (action) {
      case 'bootstrap':    out = bootstrap(!!req.refreshRates); break;
      case 'addExpense':   out = { expense: addExpense(payload) }; break;
      case 'addRecurring': out = { rule: saveRecurring(payload, true) }; break;
      case 'updateRecurring': out = { rule: saveRecurring(payload, false) }; break;
      case 'deleteRecurring': out = { id: deleteRecurring(payload.id) }; break;
      case 'updateExpense':out = { expense: updateExpense(payload) }; break;
      case 'deleteExpense':out = { id: deleteExpense(payload.id) }; break;
      case 'setPeople':    out = { people: setPeople(payload.people, payload.renames) }; break;
      case 'setPersonCurrency': out = { personCur: setPersonCurrency(payload.name, payload.currency) }; break;
      case 'createInvite': out = { invite: createInvite(payload) }; break;
      case 'revokeInvite': out = { invite: revokeInvite(payload.token) }; break;
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
      case 'restoreDeleted': out = { restored: restoreDeleted(payload.type, payload.id) }; break;
      default: return json({ ok: false, error: 'Unknown action: ' + action });
    }
    if (action === 'bootstrap') {
      out.whoami = { person: who.person || '', role: who.member ? 'member' : 'main' };
      out.changes = recentChanges(40);
      out.deleted = recentlyDeleted(50);
    }
    logChange(action, payload, before);
    afterRequest(action);
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

/**
 * A tab the code expects but the Sheet lacks (new code deployed without
 * re-running setup()) used to fail every request, including plain syncs. Create
 * it, and fill in any header cells an older Sheet is missing, once per run.
 */
var CHECKED_TABS = {};
function sheet(name) {
  var ss = book();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    if (!HEADERS[name]) throw new Error('Missing tab "' + name + '".');
    sh = ss.insertSheet(name);
  }
  if (HEADERS[name] && !CHECKED_TABS[name]) {
    var hdr = HEADERS[name];
    var need = hdr.length - sh.getMaxColumns();
    if (need > 0) sh.insertColumnsAfter(sh.getMaxColumns(), need);
    var cur = sh.getRange(1, 1, 1, hdr.length).getValues()[0];
    if (hdr.some(function (h, i) { return String(cur[i]) !== h; })) {
      sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    CHECKED_TABS[name] = true;
  }
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

/**
 * Sheets turns a typed '2030-01-15' into a date at midnight in the spreadsheet's
 * timezone. Formatting it back in the script's timezone — a separate setting that
 * often defaults elsewhere — moved every date a day earlier when the Sheet is
 * ahead (Sydney is). Read and write in the spreadsheet's own timezone.
 */
var SHEET_TZ = null;
function tz() {
  if (!SHEET_TZ) {
    try { SHEET_TZ = book().getSpreadsheetTimeZone(); } catch (e) {}
    if (!SHEET_TZ) SHEET_TZ = Session.getScriptTimeZone();
  }
  return SHEET_TZ;
}
function ymd(d) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), tz(), 'yyyy-MM-dd');
}

/* ------------------------------------------------------------------ *
 * Reads
 * ------------------------------------------------------------------ */

function bootstrap(forceRates) {
  var rates = refreshRates(!!forceRates);
  materializeRecurring(rates);
  return {
    apiVersion:  API_VERSION,
    recurring:   getRecurring(),
    projects:    getProjects(),
    people:      getPeople(),
    personCur:   personCurrencies(),
    invites:     getInvites(),
    expenses:    getExpenses(),
    rates:       rates,
    settlements: getSettlements(),
    serverDate:  ymd(new Date())
  };
}

/** { name: 'MYR' } for everyone who settles in something other than AUD. */
function personCurrencies() {
  var out = {};
  readRows(SHEET_PEOPLE).forEach(function (r) {
    var n = String(r.name).trim(), c = String(r.currency || '').trim().toUpperCase();
    if (n && /^[A-Z]{3}$/.test(c) && c !== 'AUD') out[n] = c;
  });
  return out;
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
      project_id:   String(r.project_id || ''),
      category:     category(r.category),
      created_by:   String(r.created_by || ''),
      updated_by:   String(r.updated_by || ''),
      updated_at:   r.updated_at instanceof Date ? stamp(r.updated_at) : String(r.updated_at || '')
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
      project_id: String(r.project_id || ''),
      // Rows from before per-person currencies are AUD.
      currency:     String(r.currency || 'AUD'),
      amount_local: r.amount_local === '' || r.amount_local == null
                      ? (Number(r.amount_aud) || 0) : (Number(r.amount_local) || 0),
      created_by:   String(r.created_by || ''),
      created_at:   r.created_at instanceof Date ? stamp(r.created_at) : String(r.created_at || '')
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
  var row = projectRow(p);
  if (r < 0) sheet(SHEET_PROJECTS).appendRow(row);          // upsert
  else sheet(SHEET_PROJECTS).getRange(r, 1, 1, row.length).setValues([row]);
  return getProjects().filter(function (x) { return x.id === p.id; })[0];
}

/** Deleting a project takes its expenses and settlements with it. */
function deleteProject(id) {
  var r = findRowById(SHEET_PROJECTS, id);
  if (r < 0) return id;                    // already gone: nothing to protect
  if (getProjects().length < 2) throw new Error('Keep at least one project.');

  // Recurring rules go too, or they would keep generating into a deleted project.
  // Everything goes to the Deleted tab as one batch, so restoring the project brings it all back.
  var batch = 'del' + Date.now();
  var project = getProjects().filter(function (p) { return p.id === String(id); })[0];
  getExpenses().filter(function (e) { return e.project_id === String(id); })
    .forEach(function (e) { deleteExpense(e.id, batch); });
  getSettlements().filter(function (x) { return x.project_id === String(id); })
    .forEach(function (x) { deleteSettlement(x.id, batch); });
  getRecurring().filter(function (x) { return x.project_id === String(id); })
    .forEach(function (x) { deleteRecurring(x.id, batch); });
  if (project) keepDeleted('project', project, batch);
  sheet(SHEET_PROJECTS).deleteRow(findRowById(SHEET_PROJECTS, id));
  return id;
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
 * Recurring expenses and transfers
 *
 * A rule is a template plus a schedule. Only this backend turns rules into
 * real rows — on every sync, inside the script lock — so several phones can never
 * each create the same month's rent, and a phone that has been offline for a
 * week cannot bring back an occurrence somebody deleted.
 *
 * generated_through is the high-water mark: occurrences on or before it have
 * been dealt with, whether they still exist or were deleted on purpose.
 * Each occurrence gets a deterministic id, rc_<rule>_<date>, so a retried
 * write lands on the same row instead of a new one.
 * ------------------------------------------------------------------ */

var FREQUENCIES = ['weekly', 'fortnightly', 'monthly', 'quarterly', 'yearly'];

function getRecurring() {
  return readRows(SHEET_RECURRING).map(function (r) {
    return {
      id:                String(r.id),
      project_id:        String(r.project_id || ''),
      kind:              String(r.kind || 'expense'),
      frequency:         String(r.frequency || 'monthly'),
      start_date:        r.start_date instanceof Date ? ymd(r.start_date) : String(r.start_date || ''),
      end_date:          r.end_date instanceof Date ? ymd(r.end_date) : String(r.end_date || ''),
      generated_through: r.generated_through instanceof Date ? ymd(r.generated_through) : String(r.generated_through || ''),
      paused:            r.paused === true || String(r.paused).toLowerCase() === 'true',
      template:          parseJson(r.template_json, {})
    };
  }).filter(function (x) { return x.id; });
}

function recurringRow(p) {
  return [
    String(p.id),
    String(p.project_id || ''),
    p.kind === 'transfer' ? 'transfer' : 'expense',
    FREQUENCIES.indexOf(p.frequency) >= 0 ? p.frequency : 'monthly',
    String(p.start_date || ymd(new Date())),
    String(p.end_date || ''),
    String(p.generated_through || ''),
    !!p.paused,
    JSON.stringify(p.template || {})
  ];
}

/**
 * Create or update a rule. generated_through only ever moves forward: a phone
 * holding an older copy must not rewind it and make the server regenerate
 * occurrences that were already handled.
 */
function saveRecurring(p, isNew) {
  if (!p.id) throw new Error('Recurring rule needs an id');
  var existing = getRecurring().filter(function (x) { return x.id === String(p.id); })[0];
  if (existing && existing.generated_through > String(p.generated_through || '')) {
    p.generated_through = existing.generated_through;
  }
  // Resuming continues from today; it does not backfill the paused months.
  if (existing && existing.paused && !p.paused) {
    var t = ymd(new Date());
    if (String(p.generated_through || '') < t) p.generated_through = t;
  }
  var row = recurringRow(p);
  var r = findRowById(SHEET_RECURRING, row[0]);
  if (r > 0) sheet(SHEET_RECURRING).getRange(r, 1, 1, row.length).setValues([row]);
  else sheet(SHEET_RECURRING).appendRow(row);
  return getRecurring().filter(function (x) { return x.id === row[0]; })[0];
}

/** Stops future occurrences. Ones already created are real money and stay. */
function deleteRecurring(id, batch) {
  var r = findRowById(SHEET_RECURRING, id);
  if (r < 0) return id;
  var rule = getRecurring().filter(function (x) { return x.id === String(id); })[0];
  if (rule) keepDeleted('recurring', rule, batch);
  sheet(SHEET_RECURRING).deleteRow(r);
  return id;
}

function pad2(n) { return (n < 10 ? '0' : '') + n; }
function lastDayOfMonth(y, m) { return new Date(Date.UTC(y, m + 1, 0)).getUTCDate(); }

/**
 * The n-th occurrence date (0-based) of a rule. Months count from the start
 * date's day, clamped to short months — a rule starting 31 Jan runs 28/29 Feb,
 * then back to 31 Mar, rather than drifting to the 28th forever.
 */
function nthOccurrence(start, frequency, n) {
  var p = start.split('-'), y = +p[0], m = +p[1] - 1, d = +p[2];
  if (frequency === 'weekly' || frequency === 'fortnightly') {
    var step = frequency === 'weekly' ? 7 : 14;
    var t = new Date(Date.UTC(y, m, d + step * n));
    return t.getUTCFullYear() + '-' + pad2(t.getUTCMonth() + 1) + '-' + pad2(t.getUTCDate());
  }
  var months = frequency === 'yearly' ? 12 * n : (frequency === 'quarterly' ? 3 * n : n);
  var total = m + months, yy = y + Math.floor(total / 12), mm = total % 12;
  var dd = Math.min(d, lastDayOfMonth(yy, mm));
  return yy + '-' + pad2(mm + 1) + '-' + pad2(dd);
}

/** Occurrence dates after `after` and on or before `until` (capped). */
function dueDates(rule, after, until, cap) {
  var out = [];
  for (var n = 0; n < 2000 && out.length < (cap || 60); n++) {
    var d = nthOccurrence(rule.start_date, rule.frequency, n);
    if (d > until) break;
    if (rule.end_date && d > rule.end_date) break;
    if (after && d <= after) continue;
    out.push(d);
  }
  return out;
}

function materializeRecurring(rateList) {
  var rules = getRecurring();
  if (!rules.length) return 0;
  var actor = ACTOR;
  ACTOR = 'Recurring';
  try { return makeOccurrences(rules, rateList); } finally { ACTOR = actor; }
}

function makeOccurrences(rules, rateList) {
  var rate = {};
  (rateList || []).forEach(function (r) { rate[r.currency] = Number(r.rate_to_aud) || 0; });
  rate.AUD = 1;

  var today = ymd(new Date());
  var made = 0;

  rules.forEach(function (rule) {
    if (rule.paused || !rule.start_date) return;
    var dates = dueDates(rule, rule.generated_through, today, 60);
    if (!dates.length) return;
    var t = rule.template || {};
    var doneThrough = '';

    for (var i = 0; i < dates.length; i++) {
      var d = dates[i];
      var id = 'rc_' + rule.id + '_' + d;
      if (rule.kind === 'transfer') {
        addSettlement({ id: id, project_id: rule.project_id, from: t.from, to: t.to,
                        amount_aud: Number(t.amount_aud) || 0, date: d,
                        note: t.note || 'Recurring' });
      } else {
        var local = Number(t.amount_local) || 0;
        var cur = String(t.currency || 'AUD');
        var r = rate[cur] || 0;
        // No rate yet: stop here WITHOUT advancing past this date, so it is
        // created on a later sync instead of being skipped for good.
        if (!r) break;
        addExpense({
          id: id, project_id: rule.project_id, date: d,
          city: t.city || '', description: t.description || 'Recurring',
          paid_by: t.paid_by, currency: cur,
          amount_local: local, amount_aud: Math.round(local * r * 100) / 100,
          split_type: t.split_type || 'equal',
          split_detail: t.split_detail || {}, items: t.items || null,
          category: t.category || ''
        });
      }
      made++;
      doneThrough = d;
    }

    if (!doneThrough) return;
    rule.generated_through = doneThrough;
    var row = findRowById(SHEET_RECURRING, rule.id);
    if (row > 0) {
      var col = HEADERS[SHEET_RECURRING].indexOf('generated_through') + 1;
      sheet(SHEET_RECURRING).getRange(row, col).setValue(rule.generated_through);
    }
  });

  if (made) { Logger.log('Created ' + made + ' recurring occurrence(s).'); markReportDirty(); }
  return made;
}

/* ------------------------------------------------------------------ *
 * Writes — people
 * ------------------------------------------------------------------ */

function setPeople(names, renames) {
  names = (names || []).map(function (n) { return String(n).trim(); })
                       .filter(function (n) { return n; });
  if (!names.length) throw new Error('Need at least one person');
  renames = renames || {};

  // Keep each person's settle currency, following any rename.
  var cur = personCurrencies();
  Object.keys(renames).forEach(function (old) {
    if (cur[old] && !cur[renames[old]]) cur[renames[old]] = cur[old];
  });

  var sh = sheet(SHEET_PEOPLE);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  sh.getRange(2, 1, names.length, 2).setValues(names.map(function (n) { return [n, cur[n] || '']; }));

  // An invite is tied to a name, so it follows a rename too.
  if (Object.keys(renames).length) {
    var ash = sheet(SHEET_ACCESS), last = ash.getLastRow();
    if (last > 1) {
      var vals = ash.getRange(2, 2, last - 1, 1).getValues();
      var changed = false;
      vals.forEach(function (v) { if (renames[v[0]]) { v[0] = renames[v[0]]; changed = true; } });
      if (changed) ash.getRange(2, 2, vals.length, 1).setValues(vals);
    }
  }
  return names;
}

/** Who settles in what. Blank or AUD puts them back in the shared AUD pool. */
function setPersonCurrency(name, currency) {
  name = String(name || '').trim();
  var c = String(currency || '').trim().toUpperCase();
  if (c === 'AUD') c = '';
  if (c && CURRENCIES().indexOf(c) < 0) throw new Error('Add ' + c + ' under Exchange rates first.');
  var r = findRowById(SHEET_PEOPLE, name);         // column 1 is the name
  if (r < 0) throw new Error(name + ' is not in the People list.');
  sheet(SHEET_PEOPLE).getRange(r, 2).setValue(c);
  return personCurrencies();
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
    String(p.project_id || ''),
    category(p.category)
  ];
}

function rowToExpense(row) {
  var o = {};
  HEADERS[SHEET_EXPENSES].forEach(function (h, i) { o[h] = row[i]; });
  return {
    id: String(o.id), date: o.date instanceof Date ? ymd(o.date) : String(o.date), city: o.city, description: o.description,
    paid_by: o.paid_by, currency: o.currency,
    amount_local: Number(o.amount_local), amount_aud: Number(o.amount_aud),
    split_type: o.split_type,
    split_detail: parseJson(o.split_detail_json, {}),
    items: parseJson(o.item_breakdown_json, null),
    project_id: String(o.project_id || ''),
    category: category(o.category),
    created_by: String(o.created_by || ''), updated_by: String(o.updated_by || ''),
    updated_at: o.updated_at instanceof Date ? stamp(o.updated_at) : String(o.updated_at || '')
  };
}

/**
 * Upsert by id. The phone retries anything it never got a reply for, so a save
 * that reached the Sheet but lost its response on a bad connection used to land
 * twice. Same id now means same row.
 */
function addExpense(p) {
  var row = expenseRow(p);
  var r = findRowById(SHEET_EXPENSES, row[0]);
  var n = HEADERS[SHEET_EXPENSES].length;
  // Who added it and who last changed it come from the link used, never from the phone.
  var old = r > 0 ? sheet(SHEET_EXPENSES).getRange(r, 1, 1, n).getValues()[0] : null;
  var created = old && old[row.length] ? old[row.length] : (KEEP_CREATED || ACTOR);
  var full = row.concat([created, ACTOR, stamp(new Date())]);
  if (old && sameRow(old.slice(0, row.length), row)) return rowToExpense(old);       // a retry: nothing changed
  if (r > 0) sheet(SHEET_EXPENSES).getRange(r, 1, 1, n).setValues([full]);
  else sheet(SHEET_EXPENSES).appendRow(full);
  return rowToExpense(full);
}

/** Who is making this request, for the change log and the added-by columns. */
var ACTOR = '';
var KEEP_CREATED = '';      // set while restoring, so "added by" stays with whoever first added it
function stamp(d) { return Utilities.formatDate(d instanceof Date ? d : new Date(d), tz(), 'yyyy-MM-dd HH:mm'); }
function sameRow(a, b) {
  for (var i = 0; i < b.length; i++) {
    var x = a[i] instanceof Date ? ymd(a[i]) : a[i];
    if (String(x) !== String(b[i]) && Number(x) !== Number(b[i])) return false;
  }
  return true;
}

function updateExpense(p) {
  if (!p.id) throw new Error('updateExpense needs an id');
  // Editing something another phone deleted (or that hasn't arrived yet) used to
  // throw and jam that phone's queue. Treat an update as an upsert.
  return addExpense(p);
}

function deleteExpense(id, batch) {
  var r = findRowById(SHEET_EXPENSES, id);
  if (r < 0) return id;                    // already gone: the goal is met
  keepDeleted('expense', rowToExpense(sheet(SHEET_EXPENSES).getRange(r, 1, 1, HEADERS[SHEET_EXPENSES].length).getValues()[0]), batch);
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
    String(p.project_id || ''),
    String(p.currency || 'AUD'),
    p.amount_local == null || p.amount_local === '' ? (Number(p.amount_aud) || 0) : (Number(p.amount_local) || 0)
  ];
  var r = findRowById(SHEET_SETTLEMENTS, row[0]);          // upsert, as with expenses
  if (r > 0) sheet(SHEET_SETTLEMENTS).getRange(r, 1, 1, row.length).setValues([row]);
  else sheet(SHEET_SETTLEMENTS).appendRow(row.concat([ACTOR, stamp(new Date())]));
  return { id: row[0], from: row[1], to: row[2], amount_aud: row[3],
           date: row[4], note: row[5], project_id: row[6], currency: row[7], amount_local: row[8] };
}

function deleteSettlement(id, batch) {
  var r = findRowById(SHEET_SETTLEMENTS, id);
  if (r < 0) return id;                    // already gone
  var st = getSettlements().filter(function (x) { return x.id === String(id); })[0];
  if (st) keepDeleted('settlement', st, batch);
  sheet(SHEET_SETTLEMENTS).deleteRow(r);
  return id;
}

/* ------------------------------------------------------------------ *
 * Change log
 *
 * Every write is recorded in the Changes tab: when, who (the person whose
 * link was used), what, a one-line summary, and the item before and after as
 * JSON — so a deleted expense can always be found again. A retry that
 * changes nothing (the phone re-sending something that already landed) is
 * not recorded. Rows are only ever appended.
 * ------------------------------------------------------------------ */

var CHANGE_TYPES = {
  addExpense: 'expense', updateExpense: 'expense', deleteExpense: 'expense',
  addSettlement: 'settlement', deleteSettlement: 'settlement',
  addProject: 'project', updateProject: 'project', deleteProject: 'project',
  addRecurring: 'recurring', updateRecurring: 'recurring', deleteRecurring: 'recurring',
  restoreDeleted: 'restore',
  setPeople: 'people', setPersonCurrency: 'person', setRate: 'rate', clearOverride: 'rate',
  addCurrency: 'rate', removeCurrency: 'rate', createInvite: 'invite', revokeInvite: 'invite'
};

function changeKey(action, p) {
  var t = CHANGE_TYPES[action];
  if (t === 'people') return 'people';
  if (t === 'person') return String(p.name || '');
  if (t === 'rate') return String(p.currency || '').toUpperCase();
  if (t === 'invite') return String(p.token || '');
  if (t === 'restore') return String(p.type || '') + ':' + String(p.id || '');
  return String(p.id || '');
}

/** The item as it stands, before or after a change. */
function changeSnapshot(action, p) {
  var t = CHANGE_TYPES[action], k = changeKey(action, p), row;
  if (t === 'restore') return changeSnapshot({ expense: 'addExpense', settlement: 'addSettlement', project: 'addProject',
                                               recurring: 'addRecurring' }[p.type] || '', { id: p.id });
  if (t === 'expense') { row = rowById(SHEET_EXPENSES, k); return row ? rowToExpense(row) : null; }
  if (t === 'settlement') return getSettlements().filter(function (s) { return s.id === k; })[0] || null;
  if (t === 'project') return getProjects().filter(function (x) { return x.id === k; })[0] || null;
  if (t === 'recurring') return getRecurring().filter(function (x) { return x.id === k; })[0] || null;
  if (t === 'people') return { people: getPeople() };
  if (t === 'person') return { name: k, currency: personCurrencies()[k] || 'AUD' };
  if (t === 'rate') return readRates()[k] || null;
  if (t === 'invite') {
    var inv = getInvites().filter(function (i) { return i.token === k; })[0];
    return inv ? Object.assign({}, inv, { token: inv.token.slice(0, 6) + '…' }) : null;   // the link itself stays out of the log
  }
  return null;
}

/** Leave out what changes on every save and means nothing to a reader. */
function changeView(o) {
  if (!o) return o;
  var c = JSON.parse(JSON.stringify(o));
  ['updated_by', 'updated_at', 'last_updated'].forEach(function (k) { delete c[k]; });
  return c;
}

function money(n) { return (Number(n) || 0).toFixed(2); }

function changeSummary(t, action, b, a, p) {
  var o = a || b || {};
  if (t === 'restore') return 'Restored ' + deletedSummary(p.type, o).replace(/^./, function (c) { return c.toLowerCase(); });
  if (t === 'expense') {
    var what = '"' + (o.description || 'Expense') + '" ' + o.currency + ' ' + money(o.amount_local) +
               (o.currency !== 'AUD' ? ' (A$' + money(o.amount_aud) + ')' : '') + ', paid by ' + o.paid_by;
    if (!b) return 'Added ' + what;
    if (!a) return 'Deleted ' + what;
    return 'Changed "' + (o.description || 'Expense') + '": ' + fieldDiff(b, a);
  }
  if (t === 'settlement') {
    var s = o.from + ' → ' + o.to + ' ' + (o.currency || 'AUD') + ' ' + money(o.amount_local || o.amount_aud);
    return (!b ? 'Logged settlement ' : !a ? 'Removed settlement ' : 'Changed settlement ') + s;
  }
  if (t === 'project') return (!b ? 'Added project "' : !a ? 'Deleted project "' : 'Changed project "') + o.name + '"' + (b && a ? ': ' + fieldDiff(b, a) : '');
  if (t === 'recurring') return (!b ? 'Set up' : !a ? 'Removed' : 'Changed') + ' a recurring ' + (o.kind || 'item') + (b && a ? ': ' + fieldDiff(b, a) : '');
  if (t === 'people') {
    var was = (b && b.people) || [], now = (a && a.people) || [];
    var add = now.filter(function (n) { return was.indexOf(n) < 0; }), gone = was.filter(function (n) { return now.indexOf(n) < 0; });
    return 'People: ' + [add.length ? 'added ' + add.join(', ') : '', gone.length ? 'removed ' + gone.join(', ') : ''].filter(String).join('; ');
  }
  if (t === 'person') return o.name + ' now settles in ' + o.currency;
  if (t === 'rate') return (!a ? 'Removed currency ' + (b && b.currency) : !b ? 'Added currency ' + a.currency
                            : 'Rate for ' + a.currency + ': ' + (b.rate_to_aud ? (1 / b.rate_to_aud).toFixed(4) : '—') + ' → ' + (a.rate_to_aud ? (1 / a.rate_to_aud).toFixed(4) : '—') + ' per AUD' + (a.is_manual_override ? ' (manual)' : ''));
  if (t === 'invite') return (action === 'revokeInvite' ? 'Turned off the link for ' : 'Made a ' + (o.role === 'member' ? 'full-access' : 'guest') + ' link for ') + o.person;
  return action;
}

/** "amount_aud 91.20 → 94.58; category activities → food" */
function fieldDiff(b, a) {
  var out = [];
  Object.keys(Object.assign({}, b, a)).forEach(function (k) {
    if (['created_by', 'updated_by', 'updated_at'].indexOf(k) >= 0) return;
    var x = JSON.stringify(b[k] === undefined ? null : b[k]), y = JSON.stringify(a[k] === undefined ? null : a[k]);
    if (x === y) return;
    var show = function (v) {
      v = v === undefined || v === null ? '' : Array.isArray(v) ? v.join(', ') : typeof v === 'object' ? JSON.stringify(v) : String(v);
      v = v === '' ? '(none)' : v;
      return v.length > 40 ? v.slice(0, 37) + '…' : v;
    };
    out.push(k + ' ' + show(b[k]) + ' → ' + show(a[k]));
  });
  return out.join('; ') || 'no visible change';
}

function logChange(action, p, before) {
  var t = CHANGE_TYPES[action];
  if (!t) return;
  try {
    var after = changeSnapshot(action, p);
    var b = changeView(before), a = changeView(after);
    if (JSON.stringify(b) === JSON.stringify(a)) return;          // a retry, or a no-op
    sheet(SHEET_CHANGES).appendRow([stamp(new Date()), ACTOR, action, t, changeKey(action, p).slice(0, 40),
      changeSummary(t, action, b, a, p), b ? JSON.stringify(b) : '', a ? JSON.stringify(a) : '']);
  } catch (err) { Logger.log('Change log failed: ' + err); }    // never fail a save over the log
}

/* ------------------------------------------------------------------ *
 * Deleted items
 *
 * Nothing is lost on delete: the whole item goes to the Deleted tab first,
 * with who and when. Deleting a project takes its expenses, settlements and
 * recurring rules as one batch, and restoring the project brings the batch
 * back. Restored rows stay in the tab, stamped with who restored them.
 * ------------------------------------------------------------------ */

function deletedSummary(type, o) {
  o = o || {};
  if (type === 'expense') return 'Expense "' + (o.description || 'Expense') + '" ' + o.currency + ' ' + money(o.amount_local) +
                                 ', paid by ' + o.paid_by + ', ' + o.date;
  if (type === 'settlement') return 'Settlement ' + o.from + ' → ' + o.to + ' ' + (o.currency || 'AUD') + ' ' +
                                    money(o.amount_local || o.amount_aud) + ', ' + o.date;
  if (type === 'project') return 'Project "' + o.name + '"';
  if (type === 'recurring') return 'Recurring ' + (o.kind || 'item') + ' "' + ((o.template || {}).description || (o.template || {}).note || o.id) + '"';
  return type + ' ' + (o.id || '');
}

function keepDeleted(type, item, batch) {
  sheet(SHEET_DELETED).appendRow([stamp(new Date()), ACTOR, type, String(item.id),
    String(type === 'project' ? item.id : item.project_id || ''), deletedSummary(type, item), JSON.stringify(item),
    batch || '', '', '']);
}

function deletedRows() {
  var sh = sheet(SHEET_DELETED), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS[SHEET_DELETED].length).getValues().map(function (r, i) {
    return { row: i + 2, deleted_at: r[0] instanceof Date ? stamp(r[0]) : String(r[0]), deleted_by: String(r[1]),
             type: String(r[2]), id: String(r[3]), project_id: String(r[4]), summary: String(r[5]),
             item: parseJson(r[6], null), batch: String(r[7] || ''), restored: !!String(r[8] || '') };
  });
}

/** Not yet restored, newest first, for the app. */
function recentlyDeleted(n) {
  return deletedRows().filter(function (d) { return !d.restored && d.item; }).reverse().slice(0, n)
    .map(function (d) { delete d.row; delete d.restored; return d; });
}

function restoreDeleted(type, id) {
  var rows = deletedRows().filter(function (d) { return d.type === String(type) && d.id === String(id) && !d.restored && d.item; });
  var d = rows[rows.length - 1];
  if (!d) {
    var live = type === 'expense' ? rowById(SHEET_EXPENSES, id) : type === 'settlement' ? rowById(SHEET_SETTLEMENTS, id)
             : type === 'project' ? rowById(SHEET_PROJECTS, id) : null;
    if (live) return { type: type, id: id };               // already back: a retry
    throw new Error('Nothing deleted with that id to restore.');
  }
  var list = [d];
  if (type === 'project' && d.batch) {
    list = list.concat(deletedRows().filter(function (x) { return x.batch === d.batch && !x.restored && x.type !== 'project'; }));
  } else if (d.project_id && !rowById(SHEET_PROJECTS, d.project_id)) {
    throw new Error('Its project was deleted — restore the project first.');
  }
  var done = stamp(new Date());
  list.forEach(function (x) {
    if (x.type === 'project') updateProject(x.item);
    else if (x.type === 'expense') { KEEP_CREATED = x.item.created_by || ''; try { addExpense(x.item); } finally { KEEP_CREATED = ''; } }
    else if (x.type === 'settlement') addSettlement(x.item);
    else if (x.type === 'recurring') saveRecurring(x.item, false);
    sheet(SHEET_DELETED).getRange(x.row, 9, 1, 2).setValues([[done, ACTOR]]);
  });
  return { type: type, id: id, count: list.length };
}

/** The newest changes, for the app's Settings screen. */
function recentChanges(n) {
  var sh = sheet(SHEET_CHANGES), last = sh.getLastRow();
  if (last < 2) return [];
  var from = Math.max(2, last - n + 1);
  return sh.getRange(from, 1, last - from + 1, 6).getValues().reverse().map(function (r) {
    return { when: r[0] instanceof Date ? stamp(r[0]) : String(r[0]), who: String(r[1]), type: String(r[3]),
             id: String(r[4]), summary: String(r[5]) };
  });
}

/* ------------------------------------------------------------------ *
 * Report and Summary tabs — an auditable view of the spending in the Sheet
 *
 * Report: one row per expense across every project, with each person's
 * share in AUD worked out exactly as the app does, and filters switched on.
 * Summary: totals by project, category, month and person, all formulas over
 * Report, driven by a project picker and a date range at the top.
 *
 * Both are rebuilt on the first sync after anything changes (a write marks
 * them stale), or from the Trip Split menu in the Sheet. They're outputs:
 * edits to them are overwritten, apart from the Summary's three picker cells.
 * ------------------------------------------------------------------ */

var SHEET_REPORT  = 'Report';
var SHEET_SUMMARY = 'Summary';
var CAT_LABELS = { food: 'Food & drink', transport: 'Transport', stay: 'Stay', activities: 'Activities',
                   shopping: 'Shopping', groceries: 'Groceries', bills: 'Bills', other: 'Other' };
var SPLIT_LABELS = { equal: 'equal', exact: 'exact', percent: '%', itemized: 'itemized' };

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Trip Split')
      .addItem('Rebuild Report and Summary', 'rebuildReport').addToUi();
  } catch (e) { /* opened without a UI */ }
}

function markReportDirty() {
  PropertiesService.getScriptProperties().setProperty('REPORT_DIRTY', '1');
}

/** Writes mark the report stale; the next sync rebuilds it. A failure never fails the sync. */
function afterRequest(action) {
  if (action === 'bootstrap') {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('REPORT_DIRTY') !== '1') return;
    try { rebuildReport(); props.deleteProperty('REPORT_DIRTY'); }
    catch (err) { Logger.log('Report rebuild failed: ' + err); }
  } else if (['aiStatus', 'parseReceipt', 'parseText'].indexOf(action) < 0) {
    markReportDirty();
  }
}

/** Split totalC cents by weights so the parts sum exactly — the app's allocate(), line for line. */
function allocateCents(totalC, weights) {
  var names = Object.keys(weights).filter(function (n) { return Number(weights[n]) > 0; });
  var out = {};
  if (!names.length) {
    var all = Object.keys(weights);
    if (!all.length) return out;
    var base = Math.floor(totalC / all.length);
    all.forEach(function (n, i) { out[n] = base + (i < totalC - base * all.length ? 1 : 0); });
    return out;
  }
  var sum = names.reduce(function (a, n) { return a + Number(weights[n]); }, 0);
  var used = 0, frac = [];
  names.forEach(function (n) {
    var raw = Number(weights[n]) / sum * totalC, fl = Math.floor(raw);
    out[n] = fl; used += fl; frac.push([n, raw - fl]);
  });
  frac.sort(function (a, b) { return b[1] - a[1]; });
  for (var i = 0, left = totalC - used; left > 0; i++, left--) out[frac[i % frac.length][0]]++;
  return out;
}

/** Each person's share of an expense in AUD cents — the app's sharesOf(), line for line. */
function shareCents(ex, people) {
  var totalC = Math.round((Number(ex.amount_aud) || 0) * 100);
  var d = ex.split_detail || {}, weights = {};
  if (ex.split_type === 'exact' || ex.split_type === 'percent') {
    var sh = d.shares || {};
    Object.keys(sh).forEach(function (n) { if (Number(sh[n]) > 0) weights[n] = Number(sh[n]); });
  } else if (ex.split_type === 'itemized') {
    var pack = ex.items || {}, items = pack.items || [], per = {}, itemsTotal = 0;
    items.forEach(function (it) {
      var amt = Number(it.amount) || 0, who = (it.people || []).filter(Boolean);
      if (!amt || !who.length) return;
      itemsTotal += amt;
      who.forEach(function (n) { per[n] = (per[n] || 0) + amt / who.length; });
    });
    var extra = Number(pack.extra) || 0;
    if (extra && itemsTotal > 0) {
      Object.keys(per).forEach(function (n) { per[n] += extra * (per[n] / itemsTotal); });
    } else if (extra) {
      var w = Object.keys(per).length ? Object.keys(per) : people;
      w.forEach(function (n) { per[n] = (per[n] || 0) + extra / w.length; });
    }
    weights = per;
  } else {
    var parts = (d.participants && d.participants.length) ? d.participants : people;
    parts.forEach(function (n) { weights[n] = 1; });
  }
  if (!Object.keys(weights).length) people.forEach(function (n) { weights[n] = 1; });
  return allocateCents(totalC, weights);
}

function colLetter(i) {
  var s = '';
  for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + (i - 1) % 26) + s;
  return s;
}

function freshTab(name) {
  var ss = book(), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var f = sh.getFilter && sh.getFilter();
  if (f) f.remove();
  sh.clear();
  return sh;
}

function rebuildReport() {
  var projects = getProjects(), byId = {};
  projects.forEach(function (p) { byId[p.id] = p; });
  var everyone = getPeople();
  var expenses = getExpenses().slice().sort(function (a, b) {
    return String(a.date).localeCompare(String(b.date)) || String(a.id).localeCompare(String(b.id));
  });

  // Everyone with a share somewhere gets a column, in People order; removed people last.
  var shares = expenses.map(function (ex) {
    var p = byId[ex.project_id];
    var members = (p && p.members && p.members.length) ? p.members.filter(function (n) { return everyone.indexOf(n) >= 0; }) : everyone;
    return shareCents(ex, members.length ? members : everyone);
  });
  var people = everyone.slice();
  expenses.forEach(function (ex, i) {
    [ex.paid_by].concat(Object.keys(shares[i])).forEach(function (n) { if (n && people.indexOf(n) < 0) people.push(n); });
  });

  var head = ['Date', 'Month', 'Project', 'Place', 'Category', 'Description', 'Paid by', 'Split', 'Currency',
              'Amount', 'Rate (per 1 AUD)', 'Amount (AUD)']
    .concat(people.map(function (p) { return 'Share: ' + p; }), ['Shares total', 'ID', 'Added by', 'Changed by', 'Changed at']);
  var rows = expenses.map(function (ex, i) {
    var rate = ex.currency === 'AUD' ? 1 : (ex.amount_aud > 0 ? ex.amount_local / ex.amount_aud : 0);
    var sh = shares[i], tot = 0;
    var per = people.map(function (p) { tot += sh[p] || 0; return (sh[p] || 0) / 100; });
    return [ex.date, String(ex.date).slice(0, 7), (byId[ex.project_id] || {}).name || '', ex.city,
            CAT_LABELS[ex.category] || 'Not set', ex.description, ex.paid_by, SPLIT_LABELS[ex.split_type] || ex.split_type,
            ex.currency, ex.amount_local, Math.round(rate * 1e4) / 1e4, ex.amount_aud]
      .concat(per, [tot / 100, ex.id, ex.created_by, ex.updated_by, ex.updated_at]);
  });

  var sh = freshTab(SHEET_REPORT);
  var ncol = head.length, nrow = rows.length + 1;
  sh.getRange(2, 2, Math.max(rows.length, 1), 1).setNumberFormat('@');          // "2026-10" stays text
  sh.getRange(1, 1, 1, ncol).setValues([head]).setFontWeight('bold').setBackground('#eef1f6');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, ncol).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 10, rows.length, 1).setNumberFormat('#,##0.00');
    sh.getRange(2, 11, rows.length, 1).setNumberFormat('0.0000');
    sh.getRange(2, 12, rows.length, people.length + 2).setNumberFormat('#,##0.00');
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, Math.max(nrow, 2), ncol).createFilter();

  buildSummary(projects, people, rows.map(function (r) { return r[1]; }), 13);
  return rows.length;
}

/** Totals as formulas over Report, driven by B3 (project), B4 and B5 (dates). */
function buildSummary(projects, people, months, firstShareCol) {
  var old = book().getSheetByName(SHEET_SUMMARY);
  var keep = old ? old.getRange(3, 2, 3, 1).getValues() : [['All'], [''], ['']];
  var names = projects.map(function (p) { return p.name; });
  if (names.indexOf(String(keep[0][0])) < 0) keep[0][0] = 'All';

  var sh = freshTab(SHEET_SUMMARY);
  var R = "Report!", L = R + '$L:$L';
  var filt = ',' + R + '$C:$C,$C$3,' + R + '$A:$A,">="&$C$4,' + R + '$A:$A,"<="&$C$5';
  var cells = [];                                       // [row, col, value, format] — written in one pass below
  var put = function (r, c, v, fmt) { cells.push([r, c, v, fmt]); };
  var MONEY = '#,##0.00', PCT = '0.0%';

  put(1, 1, 'Trip Split — summary');
  put(2, 1, 'Pick a project and dates; every total below follows. Amounts in AUD at each expense\'s own rate.');
  put(3, 1, 'Project');  put(3, 2, keep[0][0]);  put(3, 3, '=IF($B$3="All","*",$B$3)');
  put(4, 1, 'From');     put(4, 2, keep[1][0]);  put(4, 3, '=IF($B$4="",DATE(1900,1,1),$B$4)');
  put(5, 1, 'To');       put(5, 2, keep[2][0]);  put(5, 3, '=IF($B$5="",DATE(9999,12,31),$B$5)');
  put(7, 1, 'Total spent'); put(7, 2, '=SUMIFS(' + L + filt + ')', MONEY);
  put(7, 3, '=COUNTIFS(' + R + '$C:$C,$C$3,' + R + '$A:$A,">="&$C$4,' + R + '$A:$A,"<="&$C$5)&" expenses"');

  var r = 9, heads = [];
  var section = function (title, cols) { heads.push(r); put(r, 1, title); r++; cols.forEach(function (h, i) { put(r, i + 1, h); }); heads.push(r); r++; };
  var pct = function (row) { return '=IF($B$7=0,0,B' + row + '/$B$7)'; };

  section('By category', ['Category', 'Spent (AUD)', 'Share of total']);
  Object.keys(CAT_LABELS).map(function (k) { return CAT_LABELS[k]; }).concat(['Not set']).forEach(function (c) {
    put(r, 1, c); put(r, 2, '=SUMIFS(' + L + ',' + R + '$E:$E,$A' + r + filt + ')', MONEY); put(r, 3, pct(r), PCT); r++;
  });
  r++;
  section('By month', ['Month', 'Spent (AUD)', 'Share of total']);
  var monthList = months.filter(function (m, i) { return m && months.indexOf(m) === i; }).sort();
  monthList.forEach(function (m) {
    put(r, 1, m, '@'); put(r, 2, '=SUMIFS(' + L + ',' + R + '$B:$B,$A' + r + filt + ')', MONEY); put(r, 3, pct(r), PCT); r++;
  });
  r++;
  section('By person', ['Person', 'Paid (AUD)', 'Their share (AUD)', 'Paid − share']);
  people.forEach(function (p, i) {
    var col = colLetter(firstShareCol - 1 + i);
    put(r, 1, p);
    put(r, 2, '=SUMIFS(' + L + ',' + R + '$G:$G,$A' + r + filt + ')', MONEY);
    put(r, 3, '=SUMIFS(' + R + '$' + col + ':$' + col + filt + ')', MONEY);
    put(r, 4, '=B' + r + '-C' + r, MONEY);
    r++;
  });
  put(r, 1, 'Paid − share is before settlements and per-person currencies; the app\'s Settle screen has what is owed.'); r += 2;
  section('By project (dates apply, project picker doesn\'t)', ['Project', 'Spent (AUD)']);
  names.forEach(function (n) {
    put(r, 1, n);
    put(r, 2, '=SUMIFS(' + L + ',' + R + '$C:$C,$A' + r + ',' + R + '$A:$A,">="&$C$4,' + R + '$A:$A,"<="&$C$5)', MONEY);
    r++;
  });

  var grid = [];
  for (var i = 0; i < r; i++) { grid.push(['', '', '', '']); }
  cells.forEach(function (c) { grid[c[0] - 1][c[1] - 1] = c[2]; });
  // Formats go on first, so "2026-10" is kept as text rather than read as a date.
  var fmts = [];
  for (var j = 0; j < r; j++) { fmts.push(['General', 'General', 'General', 'General']); }
  cells.forEach(function (c) { if (c[3]) fmts[c[0] - 1][c[1] - 1] = c[3]; });
  fmts[3][1] = fmts[4][1] = fmts[3][2] = fmts[4][2] = 'yyyy-mm-dd';
  sh.getRange(1, 1, r, 4).setNumberFormats(fmts);
  sh.getRange(1, 1, grid.length, 4).setValues(grid);

  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1).setFontColor('#6b7280');
  sh.getRange(3, 3, 3, 1).setFontColor('#9ca3af');
  sh.getRange(3, 2, 3, 1).setBackground('#fff8db');
  heads.forEach(function (row) { sh.getRange(row, 1, 1, 4).setFontWeight('bold'); });
  sh.getRange(7, 1, 1, 2).setFontWeight('bold');
  sh.getRange(3, 2).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['All'].concat(names), true).build());
  sh.getRange(4, 2, 2, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireDate().setAllowInvalid(false).build());
  sh.setColumnWidth(1, 220);
}

/* ------------------------------------------------------------------ *
 * Invites — guests who see part of the data
 *
 * Your own token (API_TOKEN) sees and changes everything. Each invite in the
 * Access tab is a separate token for one person and a few projects. With it
 * the server sends only those projects and, within them, only the expenses
 * and settlements that person is part of — so nothing else ever reaches their
 * phone. They can add, edit and delete expenses they're part of, and log or
 * undo their own settlements; everything else is refused.
 * ------------------------------------------------------------------ */

function isTrue(v) { return v === true || String(v).toLowerCase() === 'true'; }

/** null for an unknown token; { revoked: true } for one that's been turned off. */
function principal(token, ownerToken) {
  if (!token) return null;
  if (token === ownerToken) return { owner: true };
  var row = readRows(SHEET_ACCESS).filter(function (r) { return String(r.token) === token; })[0];
  if (!row) return null;
  if (isTrue(row.revoked)) return { revoked: true };
  if (String(row.role) === 'member') return { owner: true, member: true, person: String(row.person) };
  return { owner: false, person: String(row.person), projects: csv(row.project_ids) };
}

function getInvites() {
  return readRows(SHEET_ACCESS).map(function (r) {
    return {
      token:       String(r.token),
      person:      String(r.person),
      project_ids: csv(r.project_ids),
      created:     r.created instanceof Date ? ymd(r.created) : String(r.created || ''),
      revoked:     isTrue(r.revoked),
      role:        String(r.role) === 'member' ? 'member' : 'guest'
    };
  }).filter(function (i) { return i.token; });
}

/** The phone makes the token, so an invite can be created offline. Upsert by token. */
function createInvite(p) {
  var token = String(p.token || '');
  if (!/^[a-f0-9]{32}$/.test(token)) throw new Error('Bad invite token.');
  var person = String(p.person || '').trim();
  if (getPeople().indexOf(person) < 0) throw new Error(person + ' is not in the People list.');
  var role = p.role === 'member' ? 'member' : 'guest';
  var ids = (p.project_ids || []).map(String).filter(function (x) { return x; });
  if (!ids.length && role === 'guest') throw new Error('Pick at least one project.');
  var row = [token, person, ids.join(','), String(p.created || ymd(new Date())), false, role];
  var r = findRowById(SHEET_ACCESS, token);
  if (r > 0) sheet(SHEET_ACCESS).getRange(r, 1, 1, row.length).setValues([row]);
  else sheet(SHEET_ACCESS).appendRow(row);
  return getInvites().filter(function (i) { return i.token === token; })[0];
}

function revokeInvite(token) {
  var r = findRowById(SHEET_ACCESS, String(token || ''));
  if (r < 0) return null;                         // never synced: nothing to turn off
  sheet(SHEET_ACCESS).getRange(r, 5).setValue(true);
  return getInvites().filter(function (i) { return i.token === token; })[0];
}

/** Is this person part of the expense — paying, or sharing in it? */
function involves(ex, name, members) {
  if (String(ex.paid_by) === name) return true;
  var d = ex.split_detail || {};
  var type = String(ex.split_type || 'equal');
  if (type === 'exact' || type === 'percent') return Number((d.shares || {})[name]) > 0;
  if (type === 'itemized') {
    var items = (ex.items && ex.items.items) || [];
    return items.some(function (it) { return (it.people || []).indexOf(name) >= 0; });
  }
  var parts = (d.participants && d.participants.length) ? d.participants : (members || []);
  return parts.indexOf(name) >= 0;
}

/** The guest's projects that still exist, by id. */
function guestProjects(who) {
  var out = {};
  getProjects().forEach(function (p) { if (who.projects.indexOf(p.id) >= 0) out[p.id] = p; });
  return out;
}

/** Everyone on the guest's projects — the names they need to split with. */
function guestPeople(who, projects) {
  var names = {};
  names[who.person] = true;
  Object.keys(projects).forEach(function (id) {
    projects[id].members.forEach(function (n) { names[n] = true; });
  });
  return getPeople().filter(function (n) { return names[n]; });
}

function guestBootstrap(who, projects) {
  var people = guestPeople(who, projects);
  var allCur = personCurrencies(), personCur = {};
  people.forEach(function (n) { if (allCur[n]) personCur[n] = allCur[n]; });
  return {
    apiVersion:  API_VERSION,
    guest:       { person: who.person, project_ids: Object.keys(projects) },
    whoami:      { person: who.person, role: 'guest' },
    recurring:   [],
    projects:    getProjects().filter(function (p) { return projects[p.id]; }),
    people:      people,
    personCur:   personCur,
    expenses:    getExpenses().filter(function (ex) {
                   var pr = projects[ex.project_id];
                   return pr && involves(ex, who.person, pr.members);
                 }),
    rates:       refreshRates(false),
    settlements: getSettlements().filter(function (s) {
                   return projects[s.project_id] && (s.from === who.person || s.to === who.person);
                 }),
    serverDate:  ymd(new Date())
  };
}

function rowById(name, id) {
  var r = findRowById(name, id);
  if (r < 0) return null;
  return sheet(name).getRange(r, 1, 1, HEADERS[name].length).getValues()[0];
}

function guestCheckExpense(who, projects, ex) {
  var pr = projects[String(ex.project_id || '')];
  if (!pr) throw new Error('That expense is in a project you haven’t been invited to.');
  if (!involves(ex, who.person, pr.members)) {
    throw new Error('You can only add or change expenses you’re part of.');
  }
}

function guestCheckSettlement(who, projects, s) {
  if (!projects[String(s.project_id || '')]) throw new Error('That settlement is in a project you haven’t been invited to.');
  if (s.from !== who.person && s.to !== who.person) throw new Error('You can only log settlements you’re part of.');
}

/** Everything a guest may do. The old row and the new one must both be theirs. */
function guestAction(who, action, p, req) {
  var projects = guestProjects(who);
  switch (action) {
    case 'bootstrap':
      return guestBootstrap(who, projects);
    case 'addExpense':
    case 'updateExpense':
      guestCheckExpense(who, projects, p);
      var old = rowById(SHEET_EXPENSES, p.id);
      if (old) guestCheckExpense(who, projects, rowToExpense(old));
      return { expense: addExpense(p) };
    case 'deleteExpense':
      var gone = rowById(SHEET_EXPENSES, p.id);
      if (gone) guestCheckExpense(who, projects, rowToExpense(gone));
      return { id: deleteExpense(p.id) };                   // kept in Deleted like anyone's
    case 'addSettlement':
      guestCheckSettlement(who, projects, p);
      var oldS = rowById(SHEET_SETTLEMENTS, p.id);
      if (oldS) guestCheckSettlement(who, projects, { from: oldS[1], to: oldS[2], project_id: oldS[6] });
      return { settlement: addSettlement(p) };
    case 'deleteSettlement':
      var s = rowById(SHEET_SETTLEMENTS, p.id);
      if (s) guestCheckSettlement(who, projects, { from: s[1], to: s[2], project_id: s[6] });
      return { id: deleteSettlement(p.id) };
    default:
      throw new Error('Only the trip organiser can change that.');
  }
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
        ? Utilities.formatDate(r.last_updated, tz(), "yyyy-MM-dd'T'HH:mm:ss")
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
        last_updated: Utilities.formatDate(new Date(), tz(), "yyyy-MM-dd'T'HH:mm:ss"),
        is_manual_override: false
      };
    });
  }

  // Last-resort seed so the app is never unusable with a zero rate.
  // No invented offline guesses: a currency with no rate stays at zero, the app
  // shows "no rate", and a rate can be typed on the expense itself.

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
    Utilities.formatDate(new Date(), tz(), "yyyy-MM-dd'T'HH:mm:ss"),
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
    last_updated: Utilities.formatDate(new Date(), tz(), "yyyy-MM-dd'T'HH:mm:ss"),
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
    text: 'Dinner, 60 dollars, split evenly',
    city: '', currency: 'AUD'
  });
  Logger.log('Parsed OK: ' + JSON.stringify(draft, null, 2));
  return draft;
}

/** The shape every parse returns. All fields required so the schema stays strict. */
function draftSchema(people, cities, currencies) {
  cities     = (cities && cities.length) ? cities : [];
  currencies = (currencies && currencies.length) ? currencies : CURRENCIES();
  return {
    type: 'object',
    additionalProperties: false,
    required: ['description', 'date', 'city', 'currency', 'amount_local', 'paid_by',
               'split_type', 'participants', 'shares', 'items', 'extra', 'note', 'category'],
    properties: {
      description:  { type: 'string', description: 'Short label, e.g. "Hotpot dinner". Never blank.' },
      date:         { type: 'string', description: 'YYYY-MM-DD. Use the date on the receipt if legible, otherwise today.' },
      city:         cities.length
                      ? { type: 'string', enum: cities, description: 'Which place this was in. Keep the one given in the prompt unless the receipt or description clearly says otherwise.' }
                      : { type: 'string', description: 'Place, if obvious. Empty string otherwise.' },
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
      note:  { type: 'string', description: 'One short sentence on anything uncertain, for the human to check. Empty if nothing.' },
      category: { type: 'string', enum: CATEGORIES,
                  description: 'food = restaurants, cafes, drinks; stay = hotels; activities = tickets, tours, attractions; ' +
                               'groceries = supermarkets; bills = utilities, rent, phone; other if none fit.' }
    }
  };
}

function draftSystemPrompt(people, city, currency, defaultPayer, todayStr, cities, currencies) {
  return [
    'You turn receipts and plain-English descriptions into a single expense for a',
    'group bill-splitting app.',
    '',
    'The people in this group are: ' + people.join(', ') + '. Use these names exactly.',
    'Currencies allowed here: ' + currencies.join(', ') + '.',
    (cities.length ? 'Places allowed here: ' + cities.join(', ') + '.' : 'There is no fixed list of places; leave the place empty unless it is obvious.'),
    '',
    'Context for this expense:',
    '- Place: ' + (city || 'not set'),
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
    '   If the split is not stated, prefer "equal" across everyone in the group.',
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
  cities     = (cities && cities.length) ? cities : [];
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
    system: system,
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
  cities     = (cities && cities.length) ? cities : [];
  currencies = (currencies && currencies.length) ? currencies : CURRENCIES();
  function person(n) { return people.indexOf(String(n)) >= 0 ? String(n) : null; }
  function num(v) { var x = Number(v); return isFinite(x) && x > 0 ? Math.round(x * 100) / 100 : 0; }

  var cur = String(d.currency || '').toUpperCase();
  if (currencies.indexOf(cur) < 0) cur = fallbackCurrency || currencies[0] || 'AUD';

  var city = String(d.city || '');
  if (cities.length ? cities.indexOf(city) < 0 : false) city = fallbackCity || '';
  city = city.slice(0, 40);

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
    note:         String(d.note || '').slice(0, 200),
    category:     category(d.category)
  };
}

function parseReceipt(p, people) {
  people = people || getPeople();
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

function parseText(p, people) {
  people = people || getPeople();
  if (!people.length) throw new Error('Add people in Settings first.');

  var text = String(p.text || '').trim();
  if (!text) throw new Error('Nothing to read — say or type what the expense was.');

  var parts = [{
    text: 'Turn this description into one expense:\n\n"' + text.slice(0, 2000) + '"'
  }];
  return aiDraft(parts, people, p.city, p.currency, p.defaultPayer || people[0], p.cities, p.currencies);
}
