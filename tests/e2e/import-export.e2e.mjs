// Browser test: import a PDF bank statement and card-app screenshots (read on the
// device, no network), then download the Excel export. Run: node tests/e2e/import-export.e2e.mjs
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { mkdtempSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { setup, runner, fixture } from './harness.mjs';

const T = await setup();
const { t, done } = runner();
const tmp = mkdtempSync(path.join(os.tmpdir(), 'tripsplit-'));
T.ok(T.api('setPeople', { people: ['Alex', 'Sam', 'Jo'] }));
T.ok(T.api('addProject', { id: 'p_trip', name: 'Greater China', start_date: '2026-09-25', end_date: '2026-11-01', members: ['Alex', 'Sam', 'Jo'], currencies: ['HKD', 'AUD'], cities: ['Hong Kong'] }));
T.book.sheets.ExchangeRates.appendRow(['HKD', 0.2, '2026-10-10T00:00:00', true]);
// Already entered by hand, at an estimated rate a little under what the card charged.
T.ok(T.api('addExpense', { id: 'e_klook', project_id: 'p_trip', paid_by: 'Alex', date: '2026-10-16', city: 'Hong Kong', description: 'Ocean Park tickets',
  currency: 'HKD', amount_local: 456, amount_aud: 91.2, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam', 'Jo'] }, category: 'activities' }));

const p = await T.phone();
await p.evaluate(() => { PREFS.me = 'Alex'; save(K_PREFS, PREFS); setActiveProject('p_trip'); });
const items = () => p.evaluate(() => IMP.items.map(i => `${i.date} ${i.merchant} ${i.amount} ${i.status}${i.why ? '/' + i.why : ''}${i.match ? ' match' : ''}`));
const idx = re => p.evaluate(src => IMP.items.findIndex(i => new RegExp(src).test(i.merchant) && i.status === 'todo'), re.source);
const openImportSheet = async () => { await p.click('#addBtn'); await p.click('#btnImport'); await p.waitForSelector('#impSheet.on'); };

await t('a PDF bank statement: every line read and checked against the balance', async () => {
  await openImportSheet();
  await p.setInputFiles('#impFiles', fixture('statement.pdf'));
  await p.waitForFunction(() => IMP && IMP.items, null, { timeout: 60000 });
  assert.deepEqual(await items(), [
    '2026-10-01 Deposit-Salary Example Employer n12345678/10/ 2500 skip/credit',
    '2026-10-01 Transportfornsw Opal Sampletown 10 todo',
    '2026-10-02 Withdrawal-Osko Payment Sam Example 50 todo',
    '2026-10-02 Withdrawal Mobile Tfr Westpac Lif 1000 skip/transfer',
    '2026-10-02 Coles Sample Hill 12.34 todo',
    '2026-10-04 Payment By Authority To Example Fitness Ref1166540 35.9 skip/directdebit',
    '2026-10-05 Dumpling House Sampletown 88 todo',
    '2026-10-06 Transportfornsw Opal Sampletown 10 todo',
    '2026-10-07 Deposit Online Tfr Westpac Lif 20 skip/credit']);
  assert.match(await p.textContent('#impBody .banner'), /Every line adds up/);
});

await t('split a line: form filled in, saved, shown back to check', async () => {
  await p.click(`[data-impsplit="${await idx(/Dumpling/)}"]`); await p.waitForSelector('#exSheet.on');
  assert.deepEqual(await p.evaluate(() => [F.city, F.date, F.desc, F.amount, F.cur, F.paid_by, F.category]),
    ['', '2026-10-05', 'Dumpling House Sampletown', '88.00', 'AUD', 'Alex', 'food']);
  await p.click('#exSave'); await p.waitForSelector('#detSheet.on');
  assert.match(await p.textContent('#detBody'), /Check this before moving on/);
  await p.click('#revOk'); await p.waitForSelector('#impSheet.on'); await p.settle();
  const ex = T.api('bootstrap', {}).expenses.find(e => /Dumpling/.test(e.description));
  assert.deepEqual([ex.amount_aud, ex.currency, ex.category, !!ex.split_detail.ref], [88, 'AUD', 'food', true]);
});

await t('ignore a merchant, mark personal, log a payment to a friend as a settlement', async () => {
  await p.click(`[data-impign="${await idx(/Opal/)}"]`);
  await p.click(`[data-imppers="${await idx(/Coles/)}"]`);
  await p.click(`[data-impsettle="${await p.evaluate(() => IMP.items.findIndex(i => i.kind === 'person'))}"]`);
  await p.settle();
  const s = (await items()).map(x => x.split(' ').pop());
  assert.deepEqual(s, ['skip/credit', 'skip/ignore', 'settled', 'skip/transfer', 'personal', 'skip/directdebit', 'split', 'skip/ignore', 'skip/credit']);
  const st = T.api('bootstrap', {}).settlements[0];
  assert.deepEqual([st.from, st.to, st.amount_aud, st.note], ['Alex', 'Sam', 50, 'From statement']);
  await p.click('#impFinish');
});

// Two overlapping views of a card app's list; the second in dark mode.
const row = (m1, m2, amt, date, green) => `<div style="display:flex;justify-content:space-between;align-items:center;padding:14px 0;border-bottom:1px solid #ccc;font-size:21px">
  <div><div style="${green ? 'color:#2e7d32' : ''}">${m1}</div>${m2 ? `<div style="${green ? 'color:#2e7d32' : ''}">${m2}</div>` : ''}<div style="color:#777">${date}</div></div><div style="${green ? 'color:#2e7d32' : ''}">${amt}</div></div>`;
async function shot(rows, file, dark){
  const s = await p.context().newPage();
  await s.setViewportSize({ width: 390, height: 700 });
  await s.setContent(`<body style="margin:0;font-family:Helvetica,Arial,sans-serif;background:${dark ? '#111' : '#f2f2f2'};color:${dark ? '#eee' : '#111'}">
    <div style="text-align:center;padding:14px 0 8px"><div style="font-weight:700;font-size:19px">26 Sep 2026 - 25 Oct 2026</div><div style="font-size:12px">Card Ending - 12345</div></div>
    <div style="margin:8px 14px;background:${dark ? '#1c1c1e' : '#fff'};border-radius:18px;padding:4px 18px">${rows}</div></body>`);
  await s.screenshot({ path: file }); await s.close();
  return file;
}

await t('card-app screenshots: overlap counted once, dark mode read, an existing expense matched', async () => {
  const a = await shot(row('ONLINE PAYMENT', 'RECEIVED - THANKYO...', '-$1,234.56', '23 Oct 2026', true) + row('KLOOK TRAVEL', 'TECHNOLOGY MELBOUR...', '$94.58', '17 Oct 2026') +
    row('SEOUL BBQ', 'SAMPLEVILLE', '$54.00', '15 Oct 2026') + row('EXAMPLE AIR MASCOT', '', '$230.58', '15 Oct 2026'), path.join(tmp, 'card1.png'));
  const b = await shot(row('EXAMPLE AIR MASCOT', '', '$230.58', '15 Oct 2026') + row('KPAY*NOODLE', 'BAR SYDNEY', '$17.00', '14 Oct 2026') +
    row('WORLD VISION', 'BURWOOD EAST', '$48.00', '14 Oct 2026'), path.join(tmp, 'card2.png'), true);
  await openImportSheet();
  await p.setInputFiles('#impFiles', [a, b]);
  await p.waitForFunction(() => IMP && IMP.items, null, { timeout: 180000 });
  assert.deepEqual(await items(), [
    '2026-10-23 Online Payment Received - Thankyo 1234.56 skip/credit',
    '2026-10-17 Klook Travel Technology Melbour 94.58 todo match',
    '2026-10-15 Seoul Bbq Sampleville 54 todo',
    '2026-10-15 Example Air Mascot 230.58 todo',
    '2026-10-14 Noodle Bar Sydney 17 todo',
    '2026-10-14 World Vision Burwood East 48 todo']);
  assert.match(await p.textContent('#impBody .banner'), /card ending 12345.*A\$444\.16/);
  await p.click(`[data-implink="${await idx(/Klook/)}"]`); await p.settle();
  assert.equal(T.api('bootstrap', {}).expenses.find(e => e.id === 'e_klook').amount_aud, 94.58, 'updated to what the card charged');
  await p.click('#impFinish');
});

await t('the Excel export downloads as one .xlsx holding the filtered rows', async () => {
  await p.click('nav [data-go="history"]');
  await p.fill('#fFrom', '2026-10-01'); await p.dispatchEvent('#fFrom', 'change');
  await p.fill('#fTo', '2026-10-10'); await p.dispatchEvent('#fTo', 'change');
  assert.match(await p.textContent('#histCount'), /^1 · A\$88\.00$/);
  // No share sheet in a desktop browser: falls back to a download.
  const [dl] = await Promise.all([p.waitForEvent('download'), p.evaluate(() => { navigator.canShare = undefined; exportExcel(); })]);
  assert.equal(dl.suggestedFilename(), 'Trip Split - Greater China - 2026-10-01 to 2026-10-10.xlsx');
  const file = path.join(tmp, 'export.xlsx'); await dl.saveAs(file);
  const text = (await readFile(file)).toString('latin1');
  assert.equal(text.slice(0, 2), 'PK');
  for (const part of ['xl/worksheets/sheet4.xml', 'Dumpling House Sampletown', 'SUBTOTAL(109,', 'Summary', 'Balances']) assert.ok(text.includes(part), part);
  assert.ok(!text.includes('Ocean Park'), 'rows outside the date range stay out');
});

await T.close();
done();
