// Runs Code.gs against an in-memory fake of Google Sheets and checks what an
// invite link can and can't see or change. Run: node tests/backend-access.test.mjs
import assert from 'node:assert/strict';

import { loadBackend } from './fake-gas.mjs';

const { ctx, book, owner } = loadBackend();
const api = (token, action, payload, extra) =>
  JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(Object.assign({ token, action, payload }, extra || {})) } }).s);
const ok = r => { assert.equal(r.ok, true, r.error); return r; };
const refused = (r, re) => { assert.equal(r.ok, false, 'expected a refusal'); if (re) assert.match(r.error, re); };

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

/* ---------- the owner sets up two projects and five people ---------- */
ok(api(owner, 'setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle', 'Aunt', 'Mum'] }));
ok(api(owner, 'addProject', { id: 'p_china', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle', 'Aunt'], currencies: ['HKD', 'AUD'] }));
const home = ok(api(owner, 'bootstrap', {})).projects.find(p => p.id === 'p_general');
ok(api(owner, 'updateProject', Object.assign({}, home, { members: ['Alex', 'Sam', 'Jo', 'Mum'] })));
// MYR has no live rate offline, so add the rate row directly the way the sheet would hold it.
book.sheets.ExchangeRates.appendRow(['MYR', 0.3333, '2026-10-10T00:00:00', true]);

const E = (id, project_id, paid_by, participants) => ({
  id, project_id, paid_by, date: '2026-10-20', currency: 'HKD', amount_local: 500, amount_aud: 100,
  split_type: 'equal', split_detail: { participants }, items: null
});
ok(api(owner, 'addExpense', E('e_dinner5', 'p_china', 'Alex', ['Alex', 'Sam', 'Jo', 'Uncle', 'Aunt'])));
ok(api(owner, 'addExpense', E('e_trio', 'p_china', 'Sam', ['Alex', 'Sam', 'Jo'])));
ok(api(owner, 'addExpense', E('e_auntonly', 'p_china', 'Aunt', ['Aunt', 'Alex'])));
ok(api(owner, 'addExpense', E('e_rent', 'p_general', 'Alex', ['Alex', 'Mum'])));
ok(api(owner, 'addExpense', Object.assign(E('e_items', 'p_china', 'Jo', []), {
  split_type: 'itemized', items: { items: [{ name: 'Tea', amount: 50, people: ['Uncle'] }, { name: 'Beer', amount: 450, people: ['Jo', 'Sam'] }], extra: 0 }
})));
ok(api(owner, 'addSettlement', { id: 's_trio', project_id: 'p_china', from: 'Jo', to: 'Sam', amount_aud: 10 }));

const UNCLE = 'a'.repeat(32);
t('owner can set a settle currency and create an invite', () => {
  const r = ok(api(owner, 'setPersonCurrency', { name: 'Uncle', currency: 'MYR' }));
  assert.deepEqual(r.personCur, { Uncle: 'MYR' });
  ok(api(owner, 'createInvite', { token: UNCLE, person: 'Uncle', project_ids: ['p_china'] }));
  const b = ok(api(owner, 'bootstrap', {}));
  assert.equal(b.invites.length, 1);
  assert.equal(b.expenses.length, 5);
});

t('an unknown currency is refused', () => {
  refused(api(owner, 'setPersonCurrency', { name: 'Aunt', currency: 'JPY' }), /Add JPY/);
});

t('a guest sees only their project, and only expenses they are in', () => {
  const b = ok(api(UNCLE, 'bootstrap', {}));
  assert.deepEqual(b.guest, { person: 'Uncle', project_ids: ['p_china'] });
  assert.deepEqual(b.projects.map(p => p.id), ['p_china']);
  assert.deepEqual(b.expenses.map(e => e.id).sort(), ['e_dinner5', 'e_items']);
  assert.deepEqual(b.settlements, []);
  assert.deepEqual(b.recurring, []);
  assert.equal(b.invites, undefined);
  assert.ok(!b.people.includes('Mum'), 'people outside their projects stay hidden');
  assert.deepEqual(b.personCur, { Uncle: 'MYR' });
});

t('a guest can add, edit and delete an expense they are part of', () => {
  ok(api(UNCLE, 'addExpense', E('e_uncle1', 'p_china', 'Uncle', ['Uncle', 'Alex'])));
  ok(api(UNCLE, 'updateExpense', Object.assign(E('e_uncle1', 'p_china', 'Alex', ['Uncle', 'Alex']), { amount_aud: 120 })));
  ok(api(UNCLE, 'deleteExpense', { id: 'e_uncle1' }));
  assert.ok(!ok(api(owner, 'bootstrap', {})).expenses.some(e => e.id === 'e_uncle1'));
});

t('a guest cannot add an expense they are not in', () => {
  refused(api(UNCLE, 'addExpense', E('e_sneaky', 'p_china', 'Alex', ['Alex', 'Sam'])), /part of/);
});

t('a guest cannot overwrite, edit themselves into, or delete someone else’s expense', () => {
  refused(api(UNCLE, 'updateExpense', E('e_trio', 'p_china', 'Uncle', ['Uncle', 'Sam'])), /part of/);
  refused(api(UNCLE, 'addExpense', E('e_trio', 'p_china', 'Uncle', ['Uncle'])), /part of/);
  refused(api(UNCLE, 'deleteExpense', { id: 'e_trio' }), /part of/);
  refused(api(UNCLE, 'deleteExpense', { id: 'e_auntonly' }), /part of/);
  const trio = ok(api(owner, 'bootstrap', {})).expenses.find(e => e.id === 'e_trio');
  assert.equal(trio.paid_by, 'Sam');
});

t('a guest cannot write into a project they were not invited to', () => {
  refused(api(UNCLE, 'addExpense', E('e_x', 'p_general', 'Uncle', ['Uncle', 'Mum'])), /haven.t been invited/);
  refused(api(UNCLE, 'deleteExpense', { id: 'e_rent' }), /haven.t been invited/);
});

t('a guest can log and undo only their own settlements', () => {
  ok(api(UNCLE, 'addSettlement', { id: 's_u', project_id: 'p_china', from: 'Uncle', to: 'Alex', currency: 'MYR', amount_local: 60, amount_aud: 20 }));
  const s = ok(api(UNCLE, 'bootstrap', {})).settlements;
  assert.deepEqual(s.map(x => [x.id, x.currency, x.amount_local]), [['s_u', 'MYR', 60]]);
  refused(api(UNCLE, 'addSettlement', { id: 's_x', project_id: 'p_china', from: 'Jo', to: 'Alex', amount_aud: 5 }), /part of/);
  refused(api(UNCLE, 'deleteSettlement', { id: 's_trio' }), /part of/);
  ok(api(UNCLE, 'deleteSettlement', { id: 's_u' }));
});

t('a guest cannot touch people, rates, projects, recurring or invites', () => {
  const nope = /organiser/;
  refused(api(UNCLE, 'setPeople', { people: ['Uncle'] }), nope);
  refused(api(UNCLE, 'setRate', { currency: 'HKD', rate_to_aud: 1 }), nope);
  refused(api(UNCLE, 'refreshRates', {}), nope);
  refused(api(UNCLE, 'updateProject', { id: 'p_china', name: 'Mine' }), nope);
  refused(api(UNCLE, 'deleteProject', { id: 'p_general' }), nope);
  refused(api(UNCLE, 'addRecurring', { id: 'r1' }), nope);
  refused(api(UNCLE, 'createInvite', { token: 'b'.repeat(32), person: 'Uncle', project_ids: ['p_general'] }), nope);
  refused(api(UNCLE, 'setPersonCurrency', { name: 'Uncle', currency: 'AUD' }), nope);
  assert.equal(ok(api(owner, 'bootstrap', {})).people.length, 6);
});

t('a rename carries the settle currency and the invite across', () => {
  ok(api(owner, 'setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle Lim', 'Aunt', 'Mum'], renames: { Uncle: 'Uncle Lim' } }));
  ok(api(owner, 'updateProject', { id: 'p_china', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle Lim', 'Aunt'], currencies: ['HKD', 'AUD'] }));
  const b = ok(api(owner, 'bootstrap', {}));
  assert.deepEqual(b.personCur, { 'Uncle Lim': 'MYR' });
  assert.equal(b.invites[0].person, 'Uncle Lim');
  assert.equal(ok(api(UNCLE, 'bootstrap', {})).guest.person, 'Uncle Lim');
});

t('an unknown token is refused, and a turned-off invite says so', () => {
  refused(api('f'.repeat(32), 'bootstrap', {}), /Bad token/);
  ok(api(owner, 'revokeInvite', { token: UNCLE }));
  refused(api(UNCLE, 'bootstrap', {}), /turned off/);
  refused(api(UNCLE, 'addExpense', E('e_late', 'p_china', 'Uncle Lim', ['Uncle Lim'])), /turned off/);
});

t('older settlement rows read back as AUD', () => {
  const s = ok(api(owner, 'bootstrap', {})).settlements.find(x => x.id === 's_trio');
  assert.equal(s.currency, 'AUD');
  assert.equal(s.amount_local, 10);
});

console.log(`\n${n} passed`);
