// Tests for the one-request sync and "nothing has changed". Run: node tests/sync.test.mjs
import assert from 'node:assert/strict';
import { loadBackend } from './fake-gas.mjs';

const { ctx, book, owner, props } = loadBackend();
const post = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).s);
const call = (token, action, payload) => post({ token, action, payload });
const sync = (token, extra) => post(Object.assign({ token, action: 'sync' }, extra));
const ok = r => { assert.equal(r.ok, true, r.error); return r; };
const changes = () => book.sheets.Changes.rows().slice(1).filter(r => r && r.length);

ok(call(owner, 'setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle'] }));
ok(call(owner, 'addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle'], currencies: ['AUD'] }));
const SAM = 'b'.repeat(32), UNCLE = 'c'.repeat(32);
ok(call(owner, 'createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));
ok(call(owner, 'createInvite', { token: UNCLE, person: 'Uncle', project_ids: ['p_trip'] }));
const E = (id, paid_by, people, extra) => Object.assign({ id, project_id: 'p_trip', date: '2026-10-17', description: 'Lunch ' + id, paid_by,
  currency: 'AUD', amount_local: 60, amount_aud: 60, split_type: 'equal', split_detail: { participants: people }, category: 'food' }, extra);
const job = (qid, action, payload) => ({ qid, action, payload });

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('a whole queue goes in one request and the fresh copy comes back with it', () => {
  const r = ok(sync(SAM, { jobs: [
    job('q1', 'addExpense', E('e1', 'Sam', ['Alex', 'Sam'])),
    job('q2', 'addExpense', E('e2', 'Alex', ['Alex', 'Sam', 'Jo'])),
    job('q3', 'updateExpense', E('e1', 'Sam', ['Alex', 'Sam'], { amount_local: 75, amount_aud: 75 })),
    job('q4', 'deleteExpense', { id: 'e2' }),
    job('q5', 'addSettlement', { id: 's1', project_id: 'p_trip', from: 'Jo', to: 'Sam', currency: 'AUD', amount_local: 20, amount_aud: 20, date: '2026-10-18' }),
  ] }));
  assert.deepEqual(r.results, ['q1', 'q2', 'q3', 'q4', 'q5'].map(qid => ({ qid, ok: true })));
  assert.deepEqual(r.expenses.map(e => [e.id, e.amount_aud, e.created_by]), [['e1', 75, 'Sam']]);
  assert.equal(r.settlements.find(s => s.id === 's1').status, 'confirmed', 'Sam logged a payment to Sam');
  assert.deepEqual(r.whoami, { person: 'Sam', role: 'member' });
  assert.ok(r.rev);
  assert.equal(r.deleted[0].id, 'e2');
  const mine = changes().filter(c => c[1] === 'Sam').map(c => c[2]);
  assert.deepEqual(mine, ['addExpense', 'addExpense', 'updateExpense', 'deleteExpense', 'addSettlement'], 'each one is in the change log');
});

t('one refused change does not stop the ones after it', () => {
  const r = ok(sync(owner, { jobs: [
    job('a', 'confirmSettlement', { id: 'nope' }),
    job('b', 'addExpense', E('e3', 'Jo', ['Jo', 'Sam'])),
    job('c', 'bootstrap', {}),
    job('d', 'sync', { jobs: [] }),
  ] }));
  assert.deepEqual(r.results.map(x => x.ok), [false, true, false, false]);
  assert.ok(r.results[0].error);
  assert.match(r.results[2].error, /Unknown action/);
  assert.ok(r.expenses.some(e => e.id === 'e3'));
});

t('nothing new: the phone is told so instead of being sent everything', () => {
  const full = ok(sync(owner, {}));
  const again = ok(sync(owner, { rev: full.rev }));
  assert.deepEqual(again, { ok: true, unchanged: true, rev: full.rev, apiVersion: full.apiVersion, serverDate: full.serverDate });
  assert.equal(ok(sync(SAM, { rev: full.rev })).unchanged, true, 'the revision is the Sheet\'s, whoever asks');
  assert.equal(ok(sync(UNCLE, { rev: full.rev })).unchanged, true);
  assert.equal(ok(sync(owner, { rev: 'stale' })).unchanged, undefined);
  assert.equal(ok(sync(owner, {})).unchanged, undefined, 'a phone with no copy always gets one');
});

t('any change by anyone moves the revision on, including an older phone\'s single request', () => {
  const rev = ok(sync(owner, {})).rev;
  ok(call(SAM, 'addExpense', E('e4', 'Sam', ['Sam', 'Jo'])));
  const r = ok(sync(owner, { rev }));
  assert.equal(r.unchanged, undefined);
  assert.notEqual(r.rev, rev);
  assert.ok(r.expenses.some(e => e.id === 'e4'));
  const r2 = ok(sync(UNCLE, { rev: r.rev, jobs: [job('g1', 'addExpense', E('e5', 'Uncle', ['Uncle', 'Alex']))] }));
  assert.notEqual(r2.rev, r.rev);
  assert.equal(ok(sync(owner, { rev: r.rev })).unchanged, undefined);
  ok(call(owner, 'bootstrap', {}));
  assert.equal(ok(sync(owner, { rev: r2.rev })).unchanged, true, 'just looking changes nothing');
});

t('typing in the Sheet, a new day, or asking for live rates all send a full copy', () => {
  let rev = ok(sync(owner, {})).rev;
  ctx.onEdit({});
  let r = ok(sync(owner, { rev }));
  assert.equal(r.unchanged, undefined, 'an edit made in the Sheet itself');
  rev = r.rev;
  props.REV_DAY = '2000-01-01';
  r = ok(sync(owner, { rev }));
  assert.equal(r.unchanged, undefined, 'first sync of the day');
  assert.notEqual(r.rev, rev, 'and every other phone gets one too');
  rev = r.rev;
  assert.equal(ok(sync(owner, { rev })).unchanged, true);
  assert.equal(ok(sync(owner, { rev, refreshRates: true })).unchanged, undefined);
});

t('a guest\'s batch is checked change by change, and a turned-off link gets nothing', () => {
  const r = ok(sync(UNCLE, { jobs: [
    job('g1', 'addExpense', E('e6', 'Uncle', ['Uncle', 'Sam'])),
    job('g2', 'addExpense', E('e7', 'Alex', ['Alex', 'Sam'])),          // not one they are in
    job('g3', 'setPeople', { people: ['Uncle'] }),
    job('g4', 'deleteExpense', { id: 'e1' }),
  ] }));
  assert.deepEqual(r.results.map(x => x.ok), [true, false, false, false]);
  assert.ok(r.guest, 'and the copy that comes back is still the guest\'s view');
  assert.equal(r.expenses.some(e => e.id === 'e1'), false);
  assert.equal(r.changes, undefined);
  const all = ok(call(owner, 'bootstrap', {}));
  assert.deepEqual(all.people, ['Alex', 'Sam', 'Jo', 'Uncle']);
  assert.ok(all.expenses.some(e => e.id === 'e1') && !all.expenses.some(e => e.id === 'e7'));
  ok(call(owner, 'revokeInvite', { token: UNCLE }));
  const off = sync(UNCLE, { rev: r.rev, jobs: [job('g5', 'addExpense', E('e8', 'Uncle', ['Uncle', 'Sam']))] });
  assert.equal(off.ok, false);
  assert.equal(ok(call(owner, 'bootstrap', {})).expenses.some(e => e.id === 'e8'), false);
});

t('a long queue can go in parts, with the copy only on the last', () => {
  const part = ok(sync(owner, { noData: true, jobs: [job('p1', 'addExpense', E('e9', 'Jo', ['Jo', 'Alex']))] }));
  assert.deepEqual(Object.keys(part).sort(), ['apiVersion', 'ok', 'results']);
  assert.ok(ok(sync(owner, {})).expenses.some(e => e.id === 'e9'));
});

console.log(`\n${n} passed`);
