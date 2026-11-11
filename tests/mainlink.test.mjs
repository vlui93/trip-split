// Tests for retiring the shared main link, and for the file the app and the
// script share. Run: node tests/mainlink.test.mjs
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import { loadBackend } from './fake-gas.mjs';

const { ctx, book, owner, props } = loadBackend();
const post = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).s);
const call = (token, action, payload) => post({ token, action, payload });
const ok = r => { assert.equal(r.ok, true, r.error); return r; };
const refused = (r, re) => { assert.equal(r.ok, false, 'expected a refusal'); if (re) assert.match(r.error, re); return r; };
const lastChange = () => book.sheets.Changes.rows().slice(1).filter(r => r && r.length).pop();

ok(call(owner, 'setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle'] }));
ok(call(owner, 'addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle'], currencies: ['AUD'] }));
const ALEX = 'a'.repeat(32), SAM = 'b'.repeat(32), UNCLE = 'c'.repeat(32);
ok(call(owner, 'createInvite', { token: ALEX, person: 'Alex', role: 'member', project_ids: [] }));
ok(call(owner, 'createInvite', { token: UNCLE, person: 'Uncle', project_ids: ['p_trip'] }));
const E = id => ({ id, project_id: 'p_trip', date: '2026-10-17', description: 'Lunch ' + id, paid_by: 'Alex', currency: 'AUD',
  amount_local: 60, amount_aud: 60, split_type: 'equal', split_detail: { participants: ['Alex', 'Uncle'] } });

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('the main link works until it is retired, and says so', () => {
  assert.equal(ok(call(owner, 'bootstrap', {})).mainLink, 'on');
  assert.equal(ok(call(ALEX, 'bootstrap', {})).mainLink, 'on');
  assert.equal(ok(call(UNCLE, 'bootstrap', {})).mainLink, undefined, 'not a guest\'s business');
});

t('only a personal full-access link can retire it', () => {
  refused(call(owner, 'retireMainLink', {}), /your own link/);
  refused(call(UNCLE, 'retireMainLink', {}), /organiser/);
  assert.equal(ok(call(owner, 'bootstrap', {})).mainLink, 'on');
});

t('retired: the main link is refused everywhere, personal and guest links carry on', () => {
  assert.equal(ok(call(ALEX, 'retireMainLink', {})).mainLink, 'off');
  assert.deepEqual([lastChange()[1], lastChange()[5]], ['Alex', 'Retired the shared main link']);
  const r = refused(call(owner, 'bootstrap', {}), /main link has been retired.*your own link/);
  assert.equal(r.mainLinkOff, true);
  refused(call(owner, 'addExpense', E('e1')), /retired/);
  refused(post({ token: owner, action: 'sync', jobs: [{ qid: 'q', action: 'addExpense', payload: E('e1') }] }), /retired/);
  refused(call(owner, 'restoreMainLink', {}), /retired/);
  ok(call(ALEX, 'addExpense', E('e2')));
  ok(call(UNCLE, 'addExpense', Object.assign(E('e3'), { paid_by: 'Uncle' })));
  const b = ok(call(ALEX, 'bootstrap', {}));
  assert.equal(b.mainLink, 'off');
  assert.deepEqual(b.expenses.map(e => e.id).sort(), ['e2', 'e3'], 'nothing from the main link got in');
  assert.equal(JSON.parse(ctx.doGet({ parameter: { action: 'version' } }).s).ok, true, 'the version check needs no link');
});

t('the last full-access link cannot be turned off while the main link is retired', () => {
  refused(call(ALEX, 'revokeInvite', { token: ALEX }), /last full-access link/);
  ok(call(ALEX, 'createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));
  ok(call(SAM, 'revokeInvite', { token: ALEX }));
  refused(call(ALEX, 'bootstrap', {}), /turned off/);
  refused(call(SAM, 'revokeInvite', { token: SAM }), /last full-access link/);
  ok(call(SAM, 'revokeInvite', { token: UNCLE }));                 // a guest link is not the way in
});

t('it can be brought back from a personal link, or from the script editor', () => {
  assert.equal(ok(call(SAM, 'restoreMainLink', {})).mainLink, 'on');
  assert.equal(lastChange()[5], 'Brought back the shared main link');
  assert.equal(ok(call(owner, 'bootstrap', {})).mainLink, 'on');
  ok(call(SAM, 'retireMainLink', {}));
  refused(call(owner, 'bootstrap', {}));
  assert.equal(ctx.restoreMainLink(), 'on');
  assert.equal(props.MAIN_LINK, undefined);
  ok(call(owner, 'bootstrap', {}));
});

t('the script says plainly when shared.js is missing', () => {
  const code = readFileSync(new URL('../Code.gs', import.meta.url), 'utf8');
  const bare = { PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'tok' }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; } }) } };
  vm.createContext(bare); vm.runInContext(code, bare);
  assert.deepEqual(JSON.parse(bare.doGet({ parameter: { action: 'version' } }).s).shared, false);
  assert.match(JSON.parse(bare.doGet({ parameter: { action: 'bootstrap', token: 'tok' } }).s).error, /missing shared\.js/);
  assert.equal(JSON.parse(ctx.doGet({ parameter: { action: 'version' } }).s).shared, true);
});

t('the app and the script read categories and shares from the one file', () => {
  const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
  assert.match(html, /<script src="shared\.js"><\/script>/);
  for (const dup of [/function allocate\s*\(/, /const CATS\s*=/, /function allocateCents/, /function shareCents/, /var CAT_LABELS/])
    assert.ok(!dup.test(html) && !dup.test(readFileSync(new URL('../Code.gs', import.meta.url), 'utf8')), 'no second copy: ' + dup);
  assert.deepEqual(Array.from(ctx.CATEGORIES), ['food', 'transport', 'stay', 'activities', 'shopping', 'groceries', 'bills', 'other']);
  assert.equal(JSON.stringify(ctx.sharesFor({ amount_aud: 100, split_type: 'equal', split_detail: {} }, ['A', 'B', 'C'])), '{"A":3334,"B":3333,"C":3333}');
  assert.equal(ctx.expenseRate({ currency: 'HKD', amount_local: 512.34, amount_aud: 100 }), 5.1234);
});

console.log(`\n${n} passed`);
