// Browser test: per-person links, the change log, undo and restore, settlement
// confirmation, and the needs-attention list. Run: node tests/e2e/audit.e2e.mjs
import assert from 'node:assert/strict';
import { setup, runner } from './harness.mjs';

const T = await setup();
const { t, done } = runner();
T.ok(T.api('setPeople', { people: ['Alex', 'Sam', 'Jo'] }));
T.ok(T.api('addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo'], currencies: ['AUD'] }));
T.ok(T.api('addProject', { id: 'p_side', name: 'Day trip', members: ['Alex', 'Sam'], currencies: ['AUD'] }));
const ALEX = 'a'.repeat(32), SAM = 'b'.repeat(32);
T.ok(T.api('createInvite', { token: ALEX, person: 'Alex', role: 'member', project_ids: [] }));
T.ok(T.api('createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));
const E = (id, pid, desc, extra) => T.ok(T.api('addExpense', Object.assign({ id, project_id: pid, date: '2026-10-17', description: desc, paid_by: 'Sam',
  currency: 'AUD', amount_local: 60, amount_aud: 60, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam'] }, category: 'food' }, extra), SAM));
E('e_hotel', 'p_trip', 'Hotel', { amount_local: 300, amount_aud: 300, category: 'stay', split_detail: { participants: ['Alex', 'Sam', 'Jo'] } });
E('e_lunch', 'p_side', 'Lunch'); E('e_park', 'p_side', 'Parking', { amount_local: 20, amount_aud: 20, category: 'transport' });

const A = await T.phone({ token: ALEX });
const B = await T.phone({ token: SAM, dialogs: d => (d.type() === 'prompt' ? d.accept('Not in my account yet') : d.accept()) });
for (const p of [A, B]) await p.evaluate(() => setActiveProject('p_trip'));
const boot = () => T.api('bootstrap', {});

await t('each phone knows whose link it is, and changes carry that name', async () => {
  assert.equal(await A.evaluate(() => PREFS.me), 'Alex');
  assert.equal(await B.evaluate(() => S.whoami.person), 'Sam');
  await A.click('#addBtn');
  await A.fill('#exAmount', '90'); await A.dispatchEvent('#exAmount', 'input');
  await A.fill('#exDesc', 'Dim sum'); await A.dispatchEvent('#exDesc', 'input');
  await A.click('#exSave'); await A.settle();
  const ex = boot().expenses.find(e => e.description === 'Dim sum');
  assert.deepEqual([ex.created_by, ex.updated_by, ex.category], ['Alex', 'Alex', 'food']);
  const logged = boot().changes.find(c => /^Added "Dim sum" AUD 90\.00/.test(c.summary));
  assert.equal(logged && logged.who, 'Alex');
});

await t('delete, then Undo, brings the expense back', async () => {
  const id = await A.evaluate(() => S.expenses.find(e => e.description === 'Dim sum').id);
  await A.evaluate(i => openEdit(i), id); await A.click('#exDelete');
  await A.click('#toastUndo'); await A.settle();
  const ex = boot().expenses.find(e => e.id === id);
  assert.equal(ex.created_by, 'Alex');
  assert.ok(boot().changes.some(c => /^Restored expense "Dim sum"/.test(c.summary) && c.who === 'Alex'));
});

await t('a deleted project is listed once and restored whole', async () => {
  await A.evaluate(() => { S.projects = S.projects.filter(x => x.id !== 'p_side'); enqueue('deleteProject', { id: 'p_side' }); });
  await A.settle(); await A.evaluate(() => sync()); await A.settle();
  await A.click('nav [data-go="history"]'); await A.click('[data-hview="activity"]');
  assert.match(await A.evaluate(() => document.querySelector('#deletedBox').innerText.replace(/\s+/g, ' ')), /Project "Day trip" and 2 items in it/);
  await A.click('[data-restore="project|p_side"]'); await A.settle();
  const b = boot();
  assert.ok(b.projects.some(p => p.id === 'p_side'));
  assert.deepEqual(b.expenses.filter(e => e.project_id === 'p_side').map(e => e.id).sort(), ['e_lunch', 'e_park']);
  assert.equal(b.deleted.length, 0);
});

await t('a payment waits for the person paid, who confirms it on their own phone', async () => {
  await A.click('nav [data-go="settle"]');
  const i = await A.evaluate(() => window.__tx.findIndex(x => x.from === 'Alex' && x.to === 'Sam'));
  await A.click(`[data-pay="${i}"]`); await A.click('#paySave'); await A.settle();
  assert.match(await A.evaluate(() => document.querySelector('#settleLog').innerText), /Waiting for Sam to confirm/);
  await B.evaluate(() => sync()); await B.settle();
  assert.deepEqual(await B.evaluate(() => attentionList().filter(x => x.kind === 'confirm').length), 1);
  await B.click('nav [data-go="settle"]'); await B.click('[data-stok]'); await B.settle();
  const st = boot().settlements[0];
  assert.deepEqual([st.status, st.confirmed_by], ['confirmed', 'Sam']);
});

await t('a disputed payment stops counting', async () => {
  await A.evaluate(() => { const st = { id: 's_x', project_id: 'p_trip', from: 'Alex', to: 'Sam', currency: 'AUD', amount_local: 50, amount_aud: 50, date: today(), status: 'pending' };
    S.settlements.push(st); enqueue('addSettlement', st); });
  await A.settle(); await B.evaluate(() => sync()); await B.settle(); await B.evaluate(() => renderSettle());
  const before = await B.evaluate(() => ledger().pool.Alex);
  await B.click('[data-stno="s_x"]'); await B.settle();
  assert.equal(await B.evaluate(() => ledger().pool.Alex), before - 5000);
  const st = boot().settlements.find(s => s.id === 's_x');
  assert.deepEqual([st.status, st.dispute_note], ['disputed', 'Not in my account yet']);
});

await t('needs attention: lists what is wrong, and its buttons fix it', async () => {
  T.ok(T.api('addExpense', { id: 'e_dup', project_id: 'p_trip', date: '2026-10-17', description: 'Hotel', paid_by: 'Sam', currency: 'AUD',
    amount_local: 300, amount_aud: 300, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam', 'Jo'] }, category: 'stay' }, SAM));
  T.ok(T.api('addExpense', { id: 'e_taxi', project_id: 'p_trip', date: '2026-10-18', description: 'Taxi', paid_by: 'Sam', currency: 'AUD',
    amount_local: 22, amount_aud: 22, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam'] }, category: '' }, SAM));
  await A.evaluate(() => sync()); await A.settle(); await A.click('nav [data-go="home"]');
  const kinds = () => A.evaluate(() => attentionList().map(x => x.kind));
  assert.deepEqual(await kinds(), ['disputed', 'dup', 'nocat']);
  assert.match(await A.textContent('#attnAll'), /3 to check/);
  await A.click('#attnAll');                                           // the full list, where an item can be dismissed
  await A.click('#attnBody [data-attnoff^="dup:"]');
  await A.click('#attnSheet [data-close="attnSheet"]');
  await A.click('#attnCard [data-attn="nocat:e_taxi"]'); await A.waitForSelector('#exSheet.on');
  assert.equal(await A.textContent('#fv-cat'), 'Not set');
  await A.click('[data-cat="transport"]'); await A.click('#exSave'); await A.settle();
  assert.deepEqual(await kinds(), ['disputed']);
  assert.equal(boot().expenses.find(e => e.id === 'e_taxi').category, 'transport');
});

await T.close();
done();
