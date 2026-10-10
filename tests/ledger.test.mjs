// Tests for who-owes-whom across settle currencies. Run: node tests/ledger.test.mjs
// Pulls the pure functions out of index.html between the ledger markers.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const src = html.slice(html.indexOf('/* ledger:start'), html.indexOf('/* ledger:end */'));
const { buildLedger, pairCur } = new Function(src + '; return { buildLedger, pairCur };')();

// AUD per 1 unit. 1 AUD ≈ 3.0 MYR, ≈ 5.0 HKD.
let RATES = { AUD: 1, MYR: 1 / 3, HKD: 0.2 };
const rateNow = c => RATES[c] || 0;
const cents = n => Math.round(n * 100);

// Equal split in AUD cents, remainder to the first names — enough for these tests.
function shares(ex){
  const who = ex.split_detail.participants;
  const total = cents(ex.amount_aud);
  const base = Math.floor(total / who.length);
  const out = {};
  who.forEach((n, i) => out[n] = base + (i < total - base * who.length ? 1 : 0));
  return out;
}
const ex = (paid_by, aud, participants, fx) =>
  ({ paid_by, amount_aud: aud, split_detail: Object.assign({ participants }, fx ? { fx } : {}) });

const trio = ['Alex', 'Sam', 'Jo'];
const five = trio.concat(['Uncle', 'Aunt']);
const curOf = n => (n === 'Uncle' || n === 'Aunt') ? 'MYR' : 'AUD';
const allAud = () => 'AUD';
const find = (pairs, from, to) => pairs.find(p => p.from === from && p.to === to);

let n = 0;
const t = (name, fn) => { RATES = { AUD: 1, MYR: 1 / 3, HKD: 0.2 }; fn(); n++; console.log('ok', name); };

t('pairCur', () => {
  assert.equal(pairCur('Alex', 'Sam', curOf), 'AUD');
  assert.equal(pairCur('Uncle', 'Alex', curOf), 'MYR');
  assert.equal(pairCur('Alex', 'Uncle', curOf), 'MYR');
  assert.equal(pairCur('Uncle', 'Aunt', curOf), 'MYR');
  const mixed = n => ({ A: 'SGD', B: 'MYR' })[n] || 'AUD';
  assert.equal(pairCur('A', 'B', mixed), 'MYR');      // two foreign: the creditor's
});

t('all AUD behaves exactly as before', () => {
  const L = buildLedger([ex('Alex', 90, trio), ex('Sam', 30, ['Sam', 'Jo'])], [], trio, allAud, rateNow, shares);
  assert.deepEqual(L.pool, { Alex: 6000, Sam: -1500, Jo: -4500 });
  assert.deepEqual(L.pairs, []);
});

t('dinner for five paid by Alex: AUD pool gets a fifth each, MYR people owe one-to-one', () => {
  const L = buildLedger([ex('Alex', 100, five, { MYR: 1 / 3 })], [], five, curOf, rateNow, shares);
  assert.deepEqual(L.pool, { Alex: 4000, Sam: -2000, Jo: -2000 });
  assert.equal(find(L.pairs, 'Uncle', 'Alex').c, 6000);   // A$20 → RM60
  assert.equal(find(L.pairs, 'Aunt', 'Alex').c, 6000);
  assert.equal(find(L.pairs, 'Uncle', 'Alex').cur, 'MYR');
});

t('expenses only the three share are untouched by the MYR people being there', () => {
  const base = [ex('Alex', 90, trio), ex('Jo', 45, trio)];
  const before = buildLedger(base, [], trio, allAud, rateNow, shares).pool;
  const after = buildLedger(base.concat([ex('Uncle', 50, ['Uncle', 'Aunt'], { MYR: 1 / 3 })]), [],
                            five, curOf, rateNow, shares);
  assert.deepEqual(after.pool, before);
  assert.equal(find(after.pairs, 'Aunt', 'Uncle').c, 7500);  // A$25 → RM75
});

t('a locked rate keeps the MYR debt fixed when rates move', () => {
  const exps = [ex('Alex', 100, five, { MYR: 1 / 3 })];
  RATES.MYR = 1 / 3.5;
  const L = buildLedger(exps, [], five, curOf, rateNow, shares);
  assert.equal(find(L.pairs, 'Uncle', 'Alex').c, 6000);
});

t('no locked rate falls back to today', () => {
  RATES.MYR = 1 / 3.5;
  const L = buildLedger([ex('Alex', 100, five)], [], five, curOf, rateNow, shares);
  assert.equal(find(L.pairs, 'Uncle', 'Alex').c, 7000);
});

t('Uncle pays: everyone else owes him in MYR, nobody in the pool moves', () => {
  const L = buildLedger([ex('Uncle', 100, five, { MYR: 1 / 3 })], [], five, curOf, rateNow, shares);
  assert.deepEqual(L.pool, { Alex: 0, Sam: 0, Jo: 0 });
  ['Alex', 'Sam', 'Jo', 'Aunt'].forEach(p => assert.equal(find(L.pairs, p, 'Uncle').c, 6000));
});

t('debts both ways between two people net to one', () => {
  const L = buildLedger([
    ex('Alex', 100, five, { MYR: 1 / 3 }),          // Uncle owes Alex RM60
    ex('Uncle', 50, ['Alex', 'Uncle'], { MYR: 1 / 3 }) // Alex owes Uncle RM75
  ], [], five, curOf, rateNow, shares);
  assert.equal(find(L.pairs, 'Uncle', 'Alex'), undefined);
  assert.equal(find(L.pairs, 'Alex', 'Uncle').c, 1500);
});

t('an MYR settlement clears the pair; an AUD one in the pool works as before', () => {
  const exps = [ex('Alex', 100, five, { MYR: 1 / 3 })];
  const sets = [
    { from: 'Uncle', to: 'Alex', currency: 'MYR', amount_local: 60, amount_aud: 20 },
    { from: 'Sam', to: 'Alex', currency: 'AUD', amount_local: 20, amount_aud: 20 },
  ];
  const L = buildLedger(exps, sets, five, curOf, rateNow, shares);
  assert.equal(find(L.pairs, 'Uncle', 'Alex'), undefined);
  assert.equal(find(L.pairs, 'Aunt', 'Alex').c, 6000);
  assert.deepEqual(L.pool, { Alex: 2000, Sam: 0, Jo: -2000 });
});

t('an old AUD-only settlement between Alex and Uncle converts at today’s rate', () => {
  const L = buildLedger([ex('Alex', 100, five, { MYR: 1 / 3 })],
    [{ from: 'Uncle', to: 'Alex', amount_aud: 10 }], five, curOf, rateNow, shares);
  assert.equal(find(L.pairs, 'Uncle', 'Alex').c, 3000);
});

t('a currency with no rate at all is kept in AUD rather than dropped', () => {
  RATES.MYR = 0;
  const L = buildLedger([ex('Alex', 100, five)], [], five, curOf, rateNow, shares);
  const p = find(L.pairs, 'Uncle', 'Alex');
  assert.equal(p.cur, 'AUD');
  assert.equal(p.c, 2000);
});

t('with no rate, an AUD settlement still clears the AUD-held debt', () => {
  RATES.MYR = 0;
  const L = buildLedger([ex('Alex', 100, five)], [{ from: 'Uncle', to: 'Alex', amount_aud: 20 }],
                        five, curOf, rateNow, shares);
  assert.equal(find(L.pairs, 'Uncle', 'Alex'), undefined);
});

console.log(`\n${n} passed`);
