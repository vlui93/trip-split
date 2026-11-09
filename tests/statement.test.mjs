// Tests for reading bank and card statements. Run: node tests/statement.test.mjs
// The samples are invented, laid out like a bank PDF statement (dated rows,
// wrapped descriptions, running balance) and a card app's transaction list.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const src = html.slice(html.indexOf('/* stmt:start'), html.indexOf('/* stmt:end */'));
const S = new Function(src + '; return { parseStatementText, parseBankStatement, parseCardList, cleanMerchant, mergeOverlapping, fingerprints, findMatch, stmtLayout };')();

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

// As a screenshot of a bank statement reads: no columns, so the balance decides.
export const BANK = `
Opening Balance + $1,000.00
TRANSACTIONS
DATE TRANSACTION DESCRIPTION DEBIT CREDIT BALANCE
30/09/26 STATEMENT OPENING BALANCE 1,000.00
01/10/26 Deposit-Salary Example Employer
n12345678/10/2026 2,500.00 3,500.00
01/10/26 Debit Card Purchase Transportfornsw Opal
Sampletown Aus 10.00 3,490.00
02/10/26 Withdrawal-Osko Payment 1234567 Sam
Example 50.00 3,440.00
02/10/26 Withdrawal Mobile 7654321 Tfr Westpac Lif 1,000.00 2,440.00
03/10/26 Eftpos Debit 0337663 Coles 0825 Sample Hill
02/10 12.34 2,427.66
03/10/26 Eftpos Debit 0337664 Coles 0825 Sample Hill
02/10 12.34 2,415.32
04/10/26 Payment By Authority To Example Fitness
Ref1166540 35.90 2,379.42
05/10/26 Debit Card Purchase Dumpling House Pty Ltd
Sampletown Aus 88.00 2,291.42
06/10/26 Deposit Online 2852650 Tfr Westpac Lif 20.00 2,311.42
Example Banking Corporation ABN 00 000 000 000 AFSL and Australian credit licence 000000 Statement No. 1 Page 1 of 2
06/10/26 Debit Card Purchase Transportfornsw Opal
Sampletown Aus 10.00 2,301.42
06/10/26 Debit Card Purchase Transportfornsw Opal
Sampletown Aus 10.00 2,291.42
07/10/26 CLOSING BALANCE 2,291.42`.split('\n');

t('bank statement: rows, wrapped descriptions, balance check on every row', () => {
  const r = S.parseStatementText(BANK, { people: ['Alex', 'Sam', 'Jo'] });
  assert.equal(r.kind, 'bank');
  assert.equal(r.opening, 1000); assert.equal(r.closing, 2291.42);
  assert.equal(r.rows.length, 11);
  assert.ok(r.rows.every(x => x.ok === true), 'every row adds up');
  const kinds = r.rows.map(x => x.kind);
  assert.deepEqual(kinds, ['credit', 'purchase', 'person', 'transfer', 'purchase', 'purchase', 'directdebit', 'purchase', 'credit', 'purchase', 'purchase']);
  assert.equal(r.rows[2].person, 'Sam');
  const coles = r.rows[4];
  assert.deepEqual([coles.merchant, coles.date, coles.posted, coles.amount, coles.credit], ['Coles Sample Hill', '2026-10-02', '2026-10-03', 12.34, false]);
  assert.equal(r.rows[7].merchant, 'Dumpling House Sampletown');
  assert.equal(r.rows[0].credit, true);
});

t('bank statement: a misread amount is caught by the balance', () => {
  const bad = BANK.map(l => l.replace('88.00 2,291.42', '86.00 2,291.42'));
  const r = S.parseStatementText(bad, {});
  assert.equal(r.rows.find(x => /Dumpling/.test(x.merchant)).ok, false);
});

t('bank statement from a PDF: columns decide debit or credit', () => {
  const lines = [
    { cols: { debit: 300, credit: 360, balance: 420 } },
    { text: '30/09/26 STATEMENT OPENING BALANCE 100.00', xs: [420] },
    { text: '01/10/26 Refund Example Store 20.00 120.00', xs: [360, 420] },
    { text: '02/10/26 Debit Card Purchase Cafe Example 4.50', xs: [300] },
  ];
  const r = S.parseBankStatement(lines, {});
  assert.deepEqual(r.rows.map(x => [x.credit, x.amount]), [[true, 20], [false, 4.5]]);
  assert.equal(r.rows[0].ok, true);
});

// As a card app's list reads through the offline reader.
export const CARD1 = `9:41
26 Sep 2026 - 25 Oct 2026
Card Ending - 12345
Filter
ONLINE PAYMENT
RECEIVED - THANKYO... -$1,234.56
23 Oct 2026
EXAMPLE TRAVEL
TECHNOLOGY MELBOUR... $94.58
17 Oct 2026
LIVE PAYMENTS*1915
CARPARK MASCOT $10.00
16 Oct 2026
SEOUL BBQ
SAMPLEVILLE $54.00
15 Oct 2026
EXAMPLE AIR MASCOT`.split('\n');
export const CARD2 = `SEOUL BBQ
SAMPLEVILLE $54.00
15 Oct 2026
EXAMPLE AIR MASCOT
$230.58
15 Oct 2026
EXAMPLE AIR MASCOT
$230.58
15 Oct 2026
KPAY*NOODLE
BAR SYDNEY $17.00
14 Oct 2026`.split('\n');

t('bank statement whose dates sit in the middle of wrapped rows', () => {
  const lines = ['30/09/26 STATEMENT OPENING BALANCE 100.00', 'Eftpos Debit 0337663 Coles 0825 Sample', '03/10/26', 'Hill',
    '02/10 12.34 87.66', 'Debit Card Purchase Dumpling House Pty', '05/10/26', 'Ltd', 'Sampletown Aus 7.66 80.00',
    '06/10/26 Debit Card Purchase Cafe 5.00 75.00', 'Example Street'];
  const r = S.parseBankStatement(lines, {});
  assert.deepEqual(r.rows.map(x => [x.merchant, x.date, x.amount, x.ok]),
    [['Coles Sample Hill', '2026-10-02', 12.34, true], ['Dumpling House Sampletown', '2026-10-05', 7.66, true],
     ['Cafe Example Street', '2026-10-06', 5, true]]);
});

t('card list: blocks, payments skipped, a cut-off row left out', () => {
  const r = S.parseStatementText(CARD1, {});
  assert.equal(r.kind, 'card'); assert.equal(r.card, '12345');
  assert.deepEqual(r.rows.map(x => [x.date, x.amount, x.credit, x.kind]), [
    ['2026-10-23', 1234.56, true, 'credit'], ['2026-10-17', 94.58, false, 'purchase'],
    ['2026-10-16', 10, false, 'purchase'], ['2026-10-15', 54, false, 'purchase']]);
  assert.equal(r.rows[2].merchant, 'Carpark Mascot');
  assert.equal(r.rows[1].merchant, 'Example Travel Technology Melbour');
});

t('card list: the amount read on the date line when the name is one line', () => {
  const r = S.parseCardList(['Card Ending - 12345', 'EXAMPLE AIR MASCOT', '15 Oct 2026 $230.58', 'KPAY*NOODLE', 'BAR SYDNEY $17.00', '14 Oct 2026'], {});
  assert.deepEqual(r.rows.map(x => [x.merchant, x.date, x.amount]),
    [['Example Air Mascot', '2026-10-15', 230.58], ['Noodle Bar Sydney', '2026-10-14', 17]]);
});

t('overlapping screenshots: overlap counted once, real repeats kept', () => {
  const a = S.parseStatementText(CARD1, {}).rows, b = S.parseStatementText(CARD2, {}).rows;
  const m = S.mergeOverlapping([a, b]);
  assert.equal(m.filter(x => /Seoul/.test(x.merchant)).length, 1, 'Seoul BBQ was in both');
  assert.equal(m.filter(x => /Example Air/.test(x.merchant)).length, 2, 'two real flights the same day');
  assert.equal(m.length, 4 + 2 + 1);
  const fp = S.fingerprints(m, 'card:12345').filter(x => /Example Air/.test(x.merchant)).map(x => x.fp);
  assert.notEqual(fp[0], fp[1]);
});

t('matching an expense already in the app', () => {
  const row = { date: '2026-10-17', amount: 94.58, merchant: 'Example Travel' };
  const exps = [
    { id: 'a', paid_by: 'Alex', date: '2026-10-16', currency: 'HKD', amount_aud: 91.2 },     // estimate, within 6%
    { id: 'b', paid_by: 'Sam', date: '2026-10-17', currency: 'AUD', amount_aud: 94.58 },     // someone else paid
    { id: 'c', paid_by: 'Alex', date: '2026-09-01', currency: 'AUD', amount_aud: 94.58 },    // too long ago
  ];
  assert.deepEqual(S.findMatch(row, exps, 'Alex'), { id: 'a', diff: 3.38 });
  assert.equal(S.findMatch(row, exps, 'Jo'), null);
  exps[0].split_detail = { ref: 'x' };
  assert.equal(S.findMatch(row, exps, 'Alex'), null, 'already linked to another line');
});

console.log(`\n${n} passed`);
