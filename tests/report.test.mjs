// Tests for the Sheet's Report and Summary tabs. Run: node tests/report.test.mjs [out.xlsx]
// Runs Code.gs against the fake Sheet, checks the Report rows against the app's
// own share maths, and (given a path) writes Report + Summary to an .xlsx so the
// Summary formulas can be recalculated by a real spreadsheet.
import { readFileSync, writeFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { loadBackend } from './fake-gas.mjs';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const between = (a, b) => html.slice(html.indexOf(a), html.indexOf(b));
// The app's allocate() and sharesOf(), with the project's people passed in.
const appShares = new Function('PEOPLE', `
  const cents = n => Math.round((Number(n)||0)*100);
  const pPeople = () => PEOPLE, S = { people: PEOPLE };
  ${between('/** Split `totalC` cents', '/* ============================================================\n   Who settles in what')}
  return sharesOf;`);
const { xlsxBuild } = new Function(between('/* xlsx:start', '/* xlsx:end */') + '; return { xlsxBuild };')();

const { ctx, book, owner } = loadBackend();
const api = (action, payload) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token: owner, action, payload }) } }).s);
const ok = r => { assert.equal(r.ok, true, r.error); return r; };

ok(api('setPeople', { people: ['Alex', 'Sam', 'Jo', 'Mum'] }));
ok(api('addProject', { id: 'p_china', name: 'Greater China', members: ['Alex', 'Sam', 'Jo'], currencies: ['HKD', 'AUD'] }));
const home = ok(api('bootstrap', {})).projects.find(p => p.id === 'p_general');
ok(api('updateProject', Object.assign({}, home, { name: 'Household', members: ['Alex', 'Mum'] })));
const X = (id, project_id, date, paid_by, aud, split_type, extra) => ok(api('addExpense', Object.assign({
  id, project_id, date, paid_by, currency: 'HKD', amount_local: Math.round(aud * 512.34) / 100, amount_aud: aud,
  description: id, city: 'Hong Kong', split_type, split_detail: {} }, extra)));
X('e1', 'p_china', '2026-10-17', 'Alex', 100, 'equal', { split_detail: { participants: ['Alex', 'Sam', 'Jo'] }, category: 'food' });
X('e2', 'p_china', '2026-10-18', 'Sam', 33.33, 'equal', { split_detail: {} , category: 'transport' });          // no participants: everyone in the project
X('e3', 'p_china', '2026-10-24', 'Jo', 90, 'exact', { split_detail: { shares: { Alex: 100, Jo: 200 } } });
X('e4', 'p_china', '2026-11-01', 'Alex', 47.5, 'percent', { split_detail: { shares: { Alex: 30, Sam: 70 } }, category: 'stay' });
X('e5', 'p_china', '2026-10-30', 'Sam', 61.11, 'itemized', { items: { items: [{ name: 'a', amount: 80, people: ['Alex'] },
  { name: 'b', amount: 120, people: ['Sam', 'Jo'] }], extra: 30 }, category: 'food' });
ok(api('addExpense', { id: 'e6', project_id: 'p_general', date: '2026-10-05', paid_by: 'Mum', currency: 'AUD', amount_local: 2000,
  amount_aud: 2000, description: 'Rent', split_type: 'equal', split_detail: { participants: ['Alex', 'Mum'] }, category: 'bills' }));

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const report = () => {
  const rows = book.sheets.Report.rows().filter(r => r && r.some(v => v !== '' && v !== undefined));
  return { head: rows[0], rows: rows.slice(1) };
};

t('the rebuild runs without error and is wired to the next sync', () => {
  assert.equal(ctx.rebuildReport(), 6);
  ok(api('addExpense', { id: 'e7', project_id: 'p_general', date: '2026-10-06', paid_by: 'Alex', currency: 'AUD', amount_local: 80,
    amount_aud: 80, description: 'Power', split_type: 'equal', split_detail: { participants: ['Alex', 'Mum'] } }));
  assert.equal(report().rows.length, 6, 'not rebuilt until a sync');
  ok(api('bootstrap', {}));
  assert.equal(report().rows.length, 7, 'rebuilt on the sync after a change');
});

t('Report: one row per expense, oldest first, with categories and filters on', () => {
  const { head, rows } = report();
  assert.deepEqual(head.slice(0, 12), ['Date', 'Month', 'Project', 'Place', 'Category', 'Description', 'Paid by', 'Split',
    'Currency', 'Amount', 'Rate (per 1 AUD)', 'Amount (AUD)']);
  assert.deepEqual(head.slice(12, 16), ['Share: Alex', 'Share: Sam', 'Share: Jo', 'Share: Mum']);
  assert.deepEqual(rows.map(r => r[0]), rows.map(r => r[0]).slice().sort());
  assert.equal(rows.find(r => r[17] === 'e3')[4], 'Not set');
  assert.equal(rows.find(r => r[17] === 'e6')[2], 'Household');
  assert.ok(book.sheets.Report.getFilter(), 'filter switched on');
  assert.equal(book.sheets.Report.formats()['2,2'], '@', 'month column kept as text');
});

t('Report shares match the app to the cent, on every split type', () => {
  const { head, rows } = report();
  const people = head.slice(12, 16).map(h => h.replace('Share: ', ''));
  const exps = ok(api('bootstrap', {})).expenses;
  rows.forEach(r => {
    const ex = exps.find(e => e.id === r[17]);
    const members = ex.project_id === 'p_china' ? ['Alex', 'Sam', 'Jo'] : ['Alex', 'Mum'];
    const app = appShares(members)(ex);
    people.forEach((p, i) => assert.equal(Math.round(r[12 + i] * 100), app[p] || 0, `${ex.id} ${p}`));
    assert.equal(Math.round(r[16] * 100), Math.round(ex.amount_aud * 100), `${ex.id} shares add up`);
  });
});

t('Summary keeps the picker cells across rebuilds', () => {
  const sm = book.sheets.Summary;
  sm.getRange(3, 2).setValue('Greater China');
  ctx.rebuildReport();
  assert.equal(sm.getRange(3, 2).getValue(), 'Greater China');
  sm.getRange(3, 2).setValue('All');
});

// Optionally write Report + Summary out so a real spreadsheet can recalculate the Summary.
if (process.argv[2]) {
  const asSheet = (name, sh) => ({ name, rows: sh.rows().map((row, r) => (row || []).map((v, c) => {
    const f = sh.formats()[(r + 1) + ',' + (c + 1)];
    if (typeof v === 'string' && v.startsWith('=')) return { f: v.slice(1), v: '' };
    if (f === 'yyyy-mm-dd' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return { v, s: 'date' };
    return v;
  })) });
  writeFileSync(process.argv[2], xlsxBuild([asSheet('Summary', book.sheets.Summary), asSheet('Report', book.sheets.Report)]));
}

console.log(`\n${n} passed`);
