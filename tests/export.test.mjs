// Tests for the Excel export. Run: node tests/export.test.mjs [out.xlsx]
// Builds a workbook from sample data with the functions in index.html, checks
// its numbers, and (given a path) writes it out so it can be opened and recalculated.
import { readFileSync, writeFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const src = html.slice(html.indexOf('/* xlsx:start'), html.indexOf('/* xlsx:end */'));
const { xlsxBuild, exportWorkbook, xlCol, xlDate } =
  new Function(src + '; return { xlsxBuild, exportWorkbook, xlCol, xlDate };')();

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('column letters and dates', () => {
  assert.deepEqual([0, 25, 26, 27, 701, 702].map(xlCol), ['A', 'Z', 'AA', 'AB', 'ZZ', 'AAA']);
  assert.equal(xlDate('1900-03-01'), 61);
  assert.equal(xlDate('2026-10-16'), 46311);
  assert.equal(xlDate('nope'), null);
});

const people = ['Alex', 'Sam', 'Jo'];
const shares = ex => {
  const who = ex.split_detail.participants, total = Math.round(ex.amount_aud * 100);
  const out = {}; who.forEach((p, i) => out[p] = Math.floor(total / who.length) + (i < total % who.length ? 1 : 0));
  return out;
};
const E = (id, date, city, cat, desc, paid, hkd, who) => ({ id, date, city, category: cat, description: desc, paid_by: paid,
  currency: 'HKD', amount_local: hkd, amount_aud: Math.round(hkd / 5.1234 * 100) / 100, split_type: 'equal',
  split_detail: { participants: who || people } });
export const SAMPLE = {
  title: 'Trip Split — Greater China', filterNote: 'All expenses · 16 Oct – 1 Nov 2026', exportedOn: '2026-11-08', people,
  expenses: [
    E('e1', '2026-10-17', 'Hong Kong', 'food', 'Dim sum <for> "five" & co', 'Alex', 600),
    E('e2', '2026-10-17', 'Hong Kong', 'stay', 'Hotel', 'Sam', 4500),
    E('e3', '2026-10-24', 'Macau', 'transport', 'Turbojet', 'Jo', 1000, ['Jo', 'Sam']),
    E('e4', '2026-11-01', 'Shanghai', '', 'Mystery', 'Alex', 333.33),
  ],
  settlements: [{ id: 's1', date: '2026-10-30', from: 'Jo', to: 'Sam', currency: 'AUD', amount_local: 50, amount_aud: 50, note: 'PayID' },
                { id: 's2', date: '2026-10-31', from: 'Uncle', to: 'Alex', currency: 'MYR', amount_local: 60, amount_aud: 20, note: '' }],
  shares, catLabel: c => ({ food: 'Food & drink', stay: 'Stay', transport: 'Transport' })[c] || 'Not set',
  splitLabel: s => s,
  ledger: { pool: { Alex: 1234, Sam: -1234, Jo: 0 }, transfers: [{ from: 'Sam', to: 'Alex', cur: 'AUD', c: 1234 }, { from: 'Uncle', to: 'Alex', cur: 'MYR', c: 6000 }] },
};

const wb = exportWorkbook(SAMPLE);
const sheet = name => wb.find(s => s.name === name);
const val = c => (c && typeof c === 'object') ? c.v : c;

t('expenses sheet: a row each, shares add up, totals follow the filter', () => {
  const rows = sheet('Expenses').rows;
  assert.deepEqual(rows[0].map(val).slice(0, 4), ['Date', 'Month', 'Place', 'Category']);
  assert.equal(rows.length, 1 + 4 + 2);
  rows.slice(1, 5).forEach(r => {
    const aud = val(r[10]), sharesSum = val(r[11]) + val(r[12]) + val(r[13]);
    assert.ok(Math.abs(aud - sharesSum) < 0.005, 'shares add up to the AUD amount');
    assert.equal(val(r[15]), 0, 'check column is zero');
  });
  assert.equal(val(rows[4][3]), 'Not set');
  const tot = rows[6][10];
  assert.equal(tot.f, 'SUBTOTAL(109,K2:K5)');
  assert.ok(Math.abs(tot.v - rows.slice(1, 5).reduce((a, r) => a + val(r[10]), 0)) < 0.005);
  assert.equal(sheet('Expenses').filter, 'A1:T5');   // 11 columns + 3 shares + total, check, ID, added/changed by, changed at
});

t('summary: totals by category, month and person agree with the rows', () => {
  const sm = sheet('Summary').rows;
  const total = sm.find(r => val(r[0]) === 'Total spent');
  const catTotal = sm.find((r, i) => val(r[0]) === 'Total' && sm.slice(0, i).some(x => val(x[0]) === 'By category'));
  assert.ok(Math.abs(val(total[1]) - val(catTotal[1])) < 0.005);
  const alex = sm.find(r => val(r[0]) === 'Alex');
  assert.equal(alex[1].f.startsWith('SUMIFS('), true);
});

t('xlsx bytes: a valid zip with every part', () => {
  const bytes = xlsxBuild(wb);
  assert.equal(String.fromCharCode(bytes[0], bytes[1]), 'PK');
  const text = new TextDecoder().decode(bytes);
  ['[Content_Types].xml', 'xl/workbook.xml', 'xl/styles.xml', 'xl/worksheets/sheet4.xml'].forEach(p => assert.ok(text.includes(p), p));
  assert.ok(text.includes('Dim sum &lt;for&gt; &quot;five&quot; &amp; co'), 'text is escaped');
  if (process.argv[2]) writeFileSync(process.argv[2], bytes);
});

console.log(`\n${n} passed`);
