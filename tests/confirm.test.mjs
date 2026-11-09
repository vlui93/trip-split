// Tests for settlement confirmation. Run: node tests/confirm.test.mjs
import assert from 'node:assert/strict';
import { loadBackend } from './fake-gas.mjs';

const { ctx, book, owner } = loadBackend();
const call = (token, action, payload) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token, action, payload }) } }).s);
const ok = r => { assert.equal(r.ok, true, r.error); return r; };
const st = id => ok(call(owner, 'bootstrap', {})).settlements.find(s => s.id === id);

ok(call(owner, 'setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle'] }));
ok(call(owner, 'addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle'], currencies: ['AUD'] }));
const ALEX = 'a'.repeat(32), SAM = 'b'.repeat(32), UNCLE = 'c'.repeat(32);
ok(call(owner, 'createInvite', { token: ALEX, person: 'Alex', role: 'member', project_ids: [] }));
ok(call(owner, 'createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));
ok(call(owner, 'createInvite', { token: UNCLE, person: 'Uncle', project_ids: ['p_trip'] }));
const S = (id, from, to, extra) => Object.assign({ id, project_id: 'p_trip', from, to, currency: 'AUD', amount_local: 40, amount_aud: 40, date: '2026-10-20' }, extra);

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('the person paid logging it confirms it there and then', () => {
  ok(call(SAM, 'addSettlement', S('s1', 'Jo', 'Sam')));
  const s = st('s1');
  assert.deepEqual([s.status, s.confirmed_by], ['confirmed', 'Sam']);
});

t('anyone else logging it leaves it waiting for the person paid', () => {
  ok(call(ALEX, 'addSettlement', S('s2', 'Jo', 'Sam')));
  assert.equal(st('s2').status, 'pending');
  assert.match(call(ALEX, 'confirmSettlement', { id: 's2' }).error, /Only Sam, on their own link/);
  assert.match(call(owner, 'confirmSettlement', { id: 's2' }).error, /Only Sam/, 'the shared main link cannot vouch');
  ok(call(SAM, 'confirmSettlement', { id: 's2' }));
  const s = st('s2');
  assert.deepEqual([s.status, s.confirmed_by], ['confirmed', 'Sam']);
  assert.match(s.confirmed_at, /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/);
});

t('disputing records a note, and the change log says so', () => {
  ok(call(owner, 'addSettlement', S('s3', 'Alex', 'Sam', { amount_local: 75, amount_aud: 75 })));
  ok(call(SAM, 'disputeSettlement', { id: 's3', note: 'Nothing arrived yet' }));
  const s = st('s3');
  assert.deepEqual([s.status, s.dispute_note], ['disputed', 'Nothing arrived yet']);
  const log = book.sheets.Changes.rows().slice(1).filter(r => r && r.length).map(r => r[5]);
  assert.ok(log.includes('Disputed Alex → Sam AUD 75.00: "Nothing arrived yet"'));
  ok(call(SAM, 'confirmSettlement', { id: 's3' }));               // changed their mind
  assert.deepEqual([st('s3').status, st('s3').dispute_note], ['confirmed', '']);
  assert.ok(book.sheets.Changes.rows().slice(1).some(r => r && r[5] === 'Confirmed receiving Alex → Sam AUD 75.00'));
});

t('a guest confirms payments to them, and only those', () => {
  ok(call(ALEX, 'addSettlement', S('s4', 'Alex', 'Uncle')));
  ok(call(ALEX, 'addSettlement', S('s5', 'Alex', 'Jo')));
  ok(call(UNCLE, 'confirmSettlement', { id: 's4' }));
  assert.equal(st('s4').confirmed_by, 'Uncle');
  assert.equal(call(UNCLE, 'confirmSettlement', { id: 's5' }).ok, false);
});

t('editing a settlement keeps its status; a restore brings it back as it was', () => {
  ok(call(owner, 'addSettlement', Object.assign(S('s2', 'Jo', 'Sam'), { note: 'PayID' })));
  assert.equal(st('s2').status, 'confirmed');
  ok(call(owner, 'deleteSettlement', { id: 's2' }));
  ok(call(owner, 'restoreDeleted', { type: 'settlement', id: 's2' }));
  const s = st('s2');
  assert.deepEqual([s.status, s.confirmed_by, s.created_by, s.note], ['confirmed', 'Sam', 'Alex', 'PayID']);
});

t('a recurring transfer’s occurrences count as agreed', () => {
  ok(call(owner, 'addRecurring', { id: 'r1', project_id: 'p_trip', kind: 'transfer', frequency: 'monthly', start_date: '2026-01-05',
    end_date: '2026-02-05', template: { from: 'Jo', to: 'Alex', amount_aud: 100, note: 'Rent share' } }));
  const occ = ok(call(owner, 'bootstrap', {})).settlements.filter(s => /^rc_r1_/.test(s.id));
  assert.ok(occ.length >= 1);
  occ.forEach(s => assert.equal(s.status, ''));
});

console.log(`\n${n} passed`);
