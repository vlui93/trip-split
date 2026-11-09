// Tests for the needs-attention list. Run: node tests/attention.test.mjs
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const src = html.slice(html.indexOf('/* attn:start'), html.indexOf('/* attn:end */'));
const { needsAttention } = new Function(src + '; return { needsAttention };')();

const E = (id, extra) => Object.assign({ id, date: '2026-10-17', description: 'Dim sum', paid_by: 'Sam', currency: 'HKD', amount_local: 600,
  amount_aud: 120, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam'] }, category: 'food' }, extra);
const base = { expenses: [], settlements: [], people: ['Alex', 'Sam', 'Jo'], me: 'Alex', ownLink: true, queueErrors: 0, importLeft: 0,
  today: '2026-11-09', dismissed: {}, catOf: c => ['food', 'stay', 'other'].includes(c) ? {} : null };
const run = over => needsAttention(Object.assign({}, base, over));
const kinds = list => list.map(i => i.kind);

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('nothing wrong, nothing listed', () => {
  assert.deepEqual(run({ expenses: [E('e1'), E('e2', { amount_local: 50 })] }), []);
});

t('payments: to confirm, disputed, and long-waiting ones', () => {
  const list = run({ settlements: [
    { id: 's1', from: 'Sam', to: 'Alex', amount_aud: 40, date: '2026-11-08', status: 'pending' },
    { id: 's2', from: 'Alex', to: 'Jo', amount_aud: 25, date: '2026-11-08', status: 'disputed', dispute_note: 'Nothing yet' },
    { id: 's3', from: 'Alex', to: 'Sam', amount_aud: 10, date: '2026-10-30', status: 'pending' },
    { id: 's4', from: 'Alex', to: 'Sam', amount_aud: 10, date: '2026-11-07', status: 'pending' },      // only two days
    { id: 's5', from: 'Jo', to: 'Sam', amount_aud: 10, date: '2026-10-01', status: 'pending' },        // not mine
  ] });
  assert.deepEqual(kinds(list), ['confirm', 'disputed', 'stale']);
  assert.equal(list[0].title, "Confirm Sam's payment of A$40.00");
  assert.equal(list[1].detail, '“Nothing yet”');
});

t('on the shared main link, payments to confirm point to making your own link instead', () => {
  const list = run({ ownLink: false, settlements: [
    { id: 's1', from: 'Sam', to: 'Alex', amount_aud: 40, date: '2026-11-08', status: 'pending' },
    { id: 's2', from: 'Jo', to: 'Alex', amount_aud: 15, date: '2026-11-08', status: 'pending' }] });
  assert.deepEqual(kinds(list), ['needlink']);
  assert.equal(list[0].title, '2 payments to you waiting to be confirmed');
});

t('likely duplicates: same payer, amount and currency within a day', () => {
  const list = run({ expenses: [E('a'), E('b', { date: '2026-10-18' }), E('c', { date: '2026-10-25' }), E('d', { paid_by: 'Jo' })] });
  assert.deepEqual(list.filter(i => i.kind === 'dup').map(i => i.ids), [['a', 'b']]);
});

t('missing rate, someone no longer in the project, no category', () => {
  const list = run({ expenses: [E('x', { amount_aud: 0 }), E('y', { amount_local: 1, split_detail: { participants: ['Alex', 'Mum'] } }),
                                E('z', { amount_local: 2, category: '' })] });
  assert.deepEqual(kinds(list), ['norate', 'gone', 'nocat']);
  assert.equal(list[1].title, 'Dim sum includes Mum');
});

t('sync errors and an unfinished import come near the top', () => {
  assert.deepEqual(kinds(run({ queueErrors: 2, importLeft: 5, expenses: [E('z', { category: '' })] })), ['sync', 'import', 'nocat']);
});

t('dismissed items stay hidden', () => {
  const list = run({ expenses: [E('a'), E('b')], dismissed: { 'dup:a+b': 1 } });
  assert.deepEqual(list, []);
});

console.log(`\n${n} passed`);
