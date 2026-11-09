// Tests for per-person links, added-by/changed-by, and the Changes tab.
// Run: node tests/changes.test.mjs
import assert from 'node:assert/strict';
import { loadBackend } from './fake-gas.mjs';

const { ctx, book, owner } = loadBackend();
const call = (token, action, payload) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token, action, payload }) } }).s);
const ok = r => { assert.equal(r.ok, true, r.error); return r; };
const changes = () => (book.sheets.Changes ? book.sheets.Changes.rows().slice(1).filter(r => r && r.length) : []);

ok(call(owner, 'setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle'] }));
book.sheets.ExchangeRates.appendRow(['HKD', '', '', false]);      // as if added from the app
ok(call(owner, 'addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle'], currencies: ['HKD', 'AUD'] }));
const SAM = 'b'.repeat(32), UNCLE = 'c'.repeat(32);
ok(call(owner, 'createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));
ok(call(owner, 'createInvite', { token: UNCLE, person: 'Uncle', project_ids: ['p_trip'] }));
const dinner = { id: 'e1', project_id: 'p_trip', date: '2026-10-17', description: 'Dim sum', paid_by: 'Sam', currency: 'HKD',
                 amount_local: 600, amount_aud: 120, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam', 'Jo'] }, category: 'food' };

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('a member link sees and does everything, as that person', () => {
  const b = ok(call(SAM, 'bootstrap', {}));
  assert.deepEqual(b.whoami, { person: 'Sam', role: 'member' });
  assert.ok(Array.isArray(b.invites), 'members manage invites like the main link');
  assert.equal(b.guest, undefined);
  ok(call(SAM, 'setRate', { currency: 'HKD', rate_to_aud: 0.2 }));
  assert.deepEqual(ok(call(owner, 'bootstrap', {})).whoami, { person: '', role: 'main' });
});

t('added by and changed by come from the link used', () => {
  ok(call(SAM, 'addExpense', dinner));
  let ex = ok(call(owner, 'bootstrap', {})).expenses.find(e => e.id === 'e1');
  assert.equal(ex.created_by, 'Sam'); assert.equal(ex.updated_by, 'Sam');
  ok(call(owner, 'updateExpense', Object.assign({}, dinner, { amount_aud: 121.5, created_by: 'Mallory' })));
  ex = ok(call(owner, 'bootstrap', {})).expenses.find(e => e.id === 'e1');
  assert.equal(ex.created_by, 'Sam', 'the phone cannot rewrite who added it');
  assert.equal(ex.updated_by, 'Main link');
  assert.match(ex.updated_at, /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/);
});

t('the Changes tab: added, changed with a diff, a retry not logged', () => {
  const rows = changes().filter(r => r[4] === 'e1');
  assert.equal(rows.length, 2);
  assert.deepEqual([rows[0][1], rows[0][2]], ['Sam', 'addExpense']);
  assert.equal(rows[0][5], 'Added "Dim sum" HKD 600.00 (A$120.00), paid by Sam');
  assert.equal(rows[1][1], 'Main link');
  assert.equal(rows[1][5], 'Changed "Dim sum": amount_aud 120 → 121.5');
  ok(call(owner, 'updateExpense', Object.assign({}, dinner, { amount_aud: 121.5 })));        // the phone retrying
  assert.equal(changes().filter(r => r[4] === 'e1').length, 2, 'nothing changed, nothing logged');
});

t('a deleted expense can be recovered from the log', () => {
  ok(call(SAM, 'deleteExpense', { id: 'e1' }));
  const del = changes().filter(r => r[4] === 'e1').pop();
  assert.equal(del[5], 'Deleted "Dim sum" HKD 600.00 (A$121.50), paid by Sam');
  const before = JSON.parse(del[6]);
  assert.equal(before.amount_local, 600); assert.deepEqual(before.split_detail.participants, ['Alex', 'Sam', 'Jo']);
  assert.equal(del[7], '');
});

t('guests are labelled, refused writes are not logged, and links stay out of the log', () => {
  ok(call(UNCLE, 'addExpense', Object.assign({}, dinner, { id: 'e2', paid_by: 'Uncle', split_detail: { participants: ['Uncle', 'Alex'] } })));
  const g = changes().filter(r => r[4] === 'e2');
  assert.equal(g[0][1], 'Uncle (guest)');
  const count = changes().length;
  assert.equal(call(UNCLE, 'setRate', { currency: 'HKD', rate_to_aud: 1 }).ok, false);
  assert.equal(changes().length, count);
  const inv = changes().filter(r => r[3] === 'invite');
  assert.equal(inv.length, 2);
  assert.equal(inv[0][5], 'Made a full-access link for Sam');
  inv.forEach(r => assert.ok(!r[6].includes(SAM) && !r[7].includes(SAM) && !r[7].includes(UNCLE)));
});

t('settlements, rates and people are logged too, and the app gets the newest first', () => {
  ok(call(SAM, 'addSettlement', { id: 's1', project_id: 'p_trip', from: 'Jo', to: 'Sam', currency: 'AUD', amount_local: 40, amount_aud: 40 }));
  ok(call(owner, 'setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle', 'Aunt'] }));
  const b = ok(call(owner, 'bootstrap', {}));
  assert.equal(b.settlements.find(s => s.id === 's1').created_by, 'Sam');
  assert.equal(b.changes[0].summary, 'People: added Aunt');
  assert.equal(b.changes[1].summary, 'Logged settlement Jo → Sam AUD 40.00');
  assert.ok(b.changes.some(c => /^Rate for HKD: — → 5\.0000 per AUD \(manual\)$/.test(c.summary)));
  assert.equal(ok(call(UNCLE, 'bootstrap', {})).changes, undefined, 'guests do not see the log');
});

console.log(`\n${n} passed`);
