// Browser test: invite a guest who settles in MYR, and check what their phone can see and do.
// Run: node tests/e2e/guests.e2e.mjs
import assert from 'node:assert/strict';
import { setup, runner } from './harness.mjs';

const T = await setup();
const { t, done } = runner();
T.ok(T.api('setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle', 'Mum'] }));
T.ok(T.api('addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle'], currencies: ['HKD', 'AUD'], cities: ['Hong Kong'] }));
const home = T.api('bootstrap', {}).projects.find(p => p.id === 'p_general');
T.ok(T.api('updateProject', Object.assign({}, home, { name: 'Household', members: ['Alex', 'Mum'] })));
T.book.sheets.ExchangeRates.appendRow(['HKD', 0.2, '2026-10-10T00:00:00', true]);
T.book.sheets.ExchangeRates.appendRow(['MYR', 1 / 3, '2026-10-10T00:00:00', true]);
T.ok(T.api('addExpense', { id: 'e_rent', project_id: 'p_general', paid_by: 'Alex', date: '2026-10-01', description: 'Rent', currency: 'AUD', amount_local: 2000, amount_aud: 2000, split_type: 'equal', split_detail: { participants: ['Alex', 'Mum'] } }));
T.ok(T.api('addExpense', { id: 'e_trio', project_id: 'p_trip', paid_by: 'Sam', date: '2026-10-17', description: 'Hotel deposit', currency: 'HKD', amount_local: 1500, amount_aud: 300, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam', 'Jo'] } }));

const A = await T.phone();
await A.evaluate(() => { PREFS.me = 'Alex'; save(K_PREFS, PREFS); setActiveProject('p_trip'); });
let link;

await t('the organiser makes a guest link that settles in MYR', async () => {
  await A.click('nav [data-go="more"]'); await A.click('#invNew');
  await A.click('[data-invrole="guest"]'); await A.click('[data-invwho="Uncle"]'); await A.click('[data-invcur="MYR"]');
  await A.click('#invCreate'); await A.waitForSelector('#invLinkSheet.on');
  link = await A.textContent('#invLinkText');
  assert.match(link, /#join=/);
  await A.click('#invLinkSheet button[data-close="invLinkSheet"]'); await A.settle();
});

await t('a dinner for five: the AUD pool and the MYR debt stay separate', async () => {
  await A.click('#addBtn');
  await A.fill('#exAmount', '500'); await A.dispatchEvent('#exAmount', 'input');
  await A.fill('#exDesc', 'Dim sum'); await A.dispatchEvent('#exDesc', 'input');
  await A.evaluate(() => { F.participants = ['Alex', 'Sam', 'Jo', 'Uncle']; renderForm(); });
  await A.click('#exSave'); await A.settle();
  const L = await A.evaluate(() => ledger());
  // Hotel: Alex and Jo each owe Sam A$100. Dinner: Sam and Jo each owe Alex A$25; Uncle's A$25 is the MYR debt.
  assert.deepEqual(L.pool, { Alex: -5000, Sam: 17500, Jo: -12500 });
  assert.deepEqual(L.pairs, [{ from: 'Uncle', to: 'Alex', cur: 'MYR', c: 7500 }]);   // A$25 at 3 MYR
});

const B = await T.phone({ link });
await t('the guest phone sees only their project and the expenses they are in', async () => {
  const v = await B.evaluate(() => ({ me: PREFS.me, guest: GUEST().person, projects: S.projects.map(p => p.name), people: S.people,
    expenses: S.expenses.map(e => e.description), invites: (S.invites || []).length,
    ownerOnly: [...document.querySelectorAll('.owner-only')].filter(el => el.offsetParent !== null).length }));
  assert.deepEqual([v.me, v.guest], ['Uncle', 'Uncle']);
  assert.deepEqual(v.projects, ['Greater China']);
  assert.deepEqual(v.expenses, ['Dim sum']);
  assert.ok(!v.people.includes('Mum'));
  assert.equal(v.invites, 0); assert.equal(v.ownerOnly, 0);
});

await t('the guest is stopped from adding an expense they are not in, then adds their own', async () => {
  await B.click('#addBtn');
  await B.fill('#exAmount', '200'); await B.dispatchEvent('#exAmount', 'input');
  await B.evaluate(() => { F.paid_by = 'Alex'; F.participants = ['Alex', 'Sam']; renderForm(); });
  await B.click('#exSave');
  assert.match(await B.textContent('#toast'), /Include yourself/);
  await B.fill('#exAmount', '100'); await B.dispatchEvent('#exAmount', 'input');
  await B.fill('#exDesc', 'Egg tarts'); await B.dispatchEvent('#exDesc', 'input');
  await B.evaluate(() => { F.paid_by = 'Uncle'; F.participants = ['Uncle', 'Alex']; renderForm(); });
  await B.click('#exSave'); await B.settle();
  const server = T.api('bootstrap', {}).expenses;
  assert.ok(server.some(e => e.description === 'Egg tarts' && e.paid_by === 'Uncle'));
  assert.ok(!server.some(e => e.amount_local === 200));
});

await t('the guest settles in MYR; the AUD pool does not move', async () => {
  await B.click('nav [data-go="settle"]');
  const i = await B.evaluate(() => window.__tx.findIndex(x => x.from === 'Uncle' && x.to === 'Alex'));
  await B.click(`[data-pay="${i}"]`);
  assert.equal(await B.textContent('#payAmtLbl'), 'Amount settled · MYR');
  await B.click('#paySave'); await B.settle();
  const st = T.api('bootstrap', {}).settlements[0];
  assert.deepEqual([st.from, st.to, st.currency, st.amount_local], ['Uncle', 'Alex', 'MYR', 45]);   // RM75 owed, less RM30 of egg tarts
  await A.evaluate(() => sync()); await A.settle();
  const L = await A.evaluate(() => ledger());
  assert.deepEqual(L.pool, { Alex: -5000, Sam: 17500, Jo: -12500 });
  assert.deepEqual(L.pairs, []);
});

await t('turning the link off locks the guest out', async () => {
  await A.click('nav [data-go="more"]'); await A.click('[data-invoff]'); await A.settle();
  await B.evaluate(() => sync()); await B.waitForFunction(() => !syncing);
  assert.match(await B.evaluate(() => lastSyncErr), /turned off/);
});

await T.close();
done();
