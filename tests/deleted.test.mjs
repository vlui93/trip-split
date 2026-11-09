// Tests for the Deleted tab and restoring. Run: node tests/deleted.test.mjs
import assert from 'node:assert/strict';
import { loadBackend } from './fake-gas.mjs';

const { ctx, book, owner } = loadBackend();
const call = (token, action, payload) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token, action, payload }) } }).s);
const ok = r => { assert.equal(r.ok, true, r.error); return r; };
const boot = () => ok(call(owner, 'bootstrap', {}));
const deletedTab = () => book.sheets.Deleted.rows().slice(1).filter(r => r && r.length);

ok(call(owner, 'setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle'] }));
ok(call(owner, 'addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle'], currencies: ['HKD', 'AUD'] }));
ok(call(owner, 'addProject', { id: 'p_side', name: 'Day trip', members: ['Alex', 'Sam'], currencies: ['AUD'] }));
const SAM = 'b'.repeat(32), UNCLE = 'c'.repeat(32);
ok(call(owner, 'createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));
ok(call(owner, 'createInvite', { token: UNCLE, person: 'Uncle', project_ids: ['p_trip'] }));
const E = (id, project_id, paid_by, extra) => Object.assign({ id, project_id, date: '2026-10-17', description: 'Dim sum ' + id, paid_by,
  currency: 'HKD', amount_local: 600, amount_aud: 120, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam'] }, category: 'food' }, extra);

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('a deleted expense is kept whole, with who and when', () => {
  ok(call(SAM, 'addExpense', E('e1', 'p_trip', 'Sam')));
  ok(call(owner, 'deleteExpense', { id: 'e1' }));
  assert.equal(boot().expenses.some(e => e.id === 'e1'), false);
  const [row] = deletedTab();
  assert.deepEqual([row[1], row[2], row[3], row[4]], ['Main link', 'expense', 'e1', 'p_trip']);
  assert.equal(row[5], 'Expense "Dim sum e1" HKD 600.00, paid by Sam, 2026-10-17');
  const item = JSON.parse(row[6]);
  assert.equal(item.date, '2026-10-17'); assert.equal(item.created_by, 'Sam');
  assert.equal(boot().deleted[0].summary, row[5]);
});

t('restoring brings it back as it was, credited to whoever first added it', () => {
  ok(call(owner, 'restoreDeleted', { type: 'expense', id: 'e1' }));
  const ex = boot().expenses.find(e => e.id === 'e1');
  assert.deepEqual([ex.amount_local, ex.category, ex.created_by, ex.date], [600, 'food', 'Sam', '2026-10-17']);
  assert.deepEqual(ex.split_detail.participants, ['Alex', 'Sam']);
  const row = deletedTab()[0];
  assert.equal(row[9], 'Main link', 'the row stays, stamped with who restored it');
  assert.equal(boot().deleted.length, 0, 'restored items leave the app list');
  const log = book.sheets.Changes.rows().slice(1).filter(r => r && r.length).pop();
  assert.equal(log[5], 'Restored expense "Dim sum e1" HKD 600.00, paid by Sam, 2026-10-17');
  ok(call(owner, 'restoreDeleted', { type: 'expense', id: 'e1' }));     // a retry is fine
});

t('settlements are kept and restored too', () => {
  ok(call(owner, 'addSettlement', { id: 's1', project_id: 'p_trip', from: 'Jo', to: 'Sam', currency: 'AUD', amount_local: 40, amount_aud: 40, date: '2026-10-20' }));
  ok(call(SAM, 'deleteSettlement', { id: 's1' }));
  assert.equal(boot().deleted[0].summary, 'Settlement Jo → Sam AUD 40.00, 2026-10-20');
  assert.equal(boot().deleted[0].deleted_by, 'Sam');
  ok(call(SAM, 'restoreDeleted', { type: 'settlement', id: 's1' }));
  assert.equal(boot().settlements.find(s => s.id === 's1').amount_aud, 40);
});

t('deleting a project keeps everything in it as one batch, restored together', () => {
  ok(call(owner, 'addExpense', E('e2', 'p_side', 'Alex')));
  ok(call(owner, 'addExpense', E('e3', 'p_side', 'Sam')));
  ok(call(owner, 'addSettlement', { id: 's2', project_id: 'p_side', from: 'Sam', to: 'Alex', amount_aud: 10, date: '2026-10-21' }));
  ok(call(owner, 'addRecurring', { id: 'r1', project_id: 'p_side', kind: 'expense', frequency: 'monthly', start_date: '2030-01-01',
    template: { description: 'Parking', paid_by: 'Alex', currency: 'AUD', amount_local: 20, split_type: 'equal', split_detail: {} } }));
  ok(call(owner, 'deleteProject', { id: 'p_side' }));
  const b = boot();
  assert.equal(b.projects.some(p => p.id === 'p_side'), false);
  assert.equal(b.expenses.some(e => e.project_id === 'p_side'), false);
  const batch = deletedTab().filter(r => r[7] && r[8] === '');
  assert.deepEqual(batch.map(r => r[2]).sort(), ['expense', 'expense', 'project', 'recurring', 'settlement']);
  assert.equal(new Set(batch.map(r => r[7])).size, 1, 'one batch');
  assert.equal(call(owner, 'restoreDeleted', { type: 'expense', id: 'e2' }).error, 'Its project was deleted — restore the project first.');
  ok(call(owner, 'restoreDeleted', { type: 'project', id: 'p_side' }));
  const a = boot();
  assert.equal(a.projects.find(p => p.id === 'p_side').name, 'Day trip');
  assert.deepEqual(a.expenses.filter(e => e.project_id === 'p_side').map(e => e.id).sort(), ['e2', 'e3']);
  assert.ok(a.settlements.some(s => s.id === 's2'));
  assert.ok(a.recurring.some(r => r.id === 'r1'));
  assert.equal(a.deleted.length, 0);
});

t('a guest’s delete is kept too, but only full access can see or restore', () => {
  ok(call(UNCLE, 'addExpense', E('e4', 'p_trip', 'Uncle', { split_detail: { participants: ['Uncle', 'Alex'] } })));
  ok(call(UNCLE, 'deleteExpense', { id: 'e4' }));
  assert.equal(boot().deleted[0].deleted_by, 'Uncle (guest)');
  const g = ok(call(UNCLE, 'bootstrap', {}));
  assert.equal(g.deleted, undefined);
  assert.match(call(UNCLE, 'restoreDeleted', { type: 'expense', id: 'e4' }).error, /organiser/);
});

console.log(`\n${n} passed`);
