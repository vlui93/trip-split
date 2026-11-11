// Browser test: the offline queue goes up in one request, a sync with nothing
// new fetches nothing, and an older backend still works. Run: node tests/e2e/sync.e2e.mjs
import assert from 'node:assert/strict';
import { setup, runner } from './harness.mjs';

const T = await setup();
const { t, done } = runner();
T.ok(T.api('setPeople', { people: ['Alex', 'Sam', 'Jo'] }));
T.ok(T.api('addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo'], currencies: ['AUD'] }));
const ALEX = 'a'.repeat(32), SAM = 'b'.repeat(32);
T.ok(T.api('createInvite', { token: ALEX, person: 'Alex', role: 'member', project_ids: [] }));
T.ok(T.api('createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));
const boot = () => T.api('bootstrap', {});

const A = await T.phone({ token: ALEX });
const B = await T.phone({ token: SAM });
for (const p of [A, B]) await p.evaluate(() => setActiveProject('p_trip'));
const add = async (p, amount, desc) => {
  await p.click('#addBtn');
  await p.fill('#exAmount', String(amount)); await p.dispatchEvent('#exAmount', 'input');
  await p.fill('#exDesc', desc); await p.dispatchEvent('#exDesc', 'input');
  await p.click('#exSave');
};
const since = (p, from) => p.requests.slice(from);

await t('three expenses added with no signal go up in one request when it returns', async () => {
  await A.settle();
  await A.offline(true);
  await add(A, 30, 'Coffee'); await add(A, 45, 'Taxi'); await add(A, 120, 'Dinner');
  await A.waitForFunction(() => !syncing);
  assert.equal(await A.evaluate(() => QUEUE.length), 3, 'kept on the phone');
  assert.equal(await A.evaluate(() => S.expenses.length), 3, 'and shown straight away');
  assert.equal(boot().expenses.length, 0);
  const mark = A.requests.length, sofar = A.batches.length;
  await A.offline(false);
  await A.settle();
  assert.deepEqual(A.batches[sofar], ['addExpense', 'addExpense', 'addExpense'], 'all three in the first request');
  assert.ok(since(A, mark).every(a => a === 'sync'), 'and none sent one by one');
  assert.deepEqual(boot().expenses.map(e => e.description).sort(), ['Coffee', 'Dinner', 'Taxi']);
  assert.ok(boot().expenses.every(e => e.created_by === 'Alex'));
});

await t('a sync with nothing new keeps the phone as it is; a change elsewhere comes through', async () => {
  await B.evaluate(() => sync()); await B.settle();
  assert.equal(await B.evaluate(() => S.expenses.length), 3);
  const before = await B.evaluate(() => JSON.stringify(S));
  let unchanged = 0;
  B.on('response', async r => { if (r.url().startsWith('https://script.google.com/')) { try { if ((await r.json()).unchanged) unchanged++; } catch {} } });
  await B.evaluate(() => sync()); await B.settle();
  await B.evaluate(() => sync()); await B.settle();
  assert.equal(unchanged, 2, 'the backend answered "unchanged" both times');
  assert.equal(await B.evaluate(() => JSON.stringify(S)), before);
  await add(A, 18, 'Snacks'); await A.settle();
  await B.evaluate(() => sync()); await B.settle();
  assert.equal(await B.evaluate(() => S.expenses.some(e => e.description === 'Snacks')), true);
  assert.equal(unchanged, 2);
});

await t('one refused change is flagged and the rest still arrive', async () => {
  const mark = A.requests.length;
  await A.evaluate(() => {
    QUEUE.push({ qid: 'bad1', action: 'confirmSettlement', payload: { id: 'missing' } });
    save(K_QUEUE, QUEUE);
  });
  await add(A, 9, 'Water');
  await A.waitForFunction(() => !syncing && QUEUE.length === 1 && QUEUE[0].err);
  assert.ok(boot().expenses.some(e => e.description === 'Water'));
  assert.ok(since(A, mark).every(a => a === 'sync'));
  assert.match(await A.textContent('#syncBar, #syncMsg, body'), /couldn.t sync/);
  await A.evaluate(() => { QUEUE = []; save(K_QUEUE, QUEUE); renderSync(); });
});

await t('"Sync now" always fetches a full copy', async () => {
  await B.evaluate(() => sync()); await B.settle();
  await B.evaluate(() => { S.expenses = []; });               // as if this phone's copy had gone wrong
  await B.evaluate(() => sync()); await B.settle();
  assert.equal(await B.evaluate(() => S.expenses.length), 0, 'an ordinary sync trusts its copy');
  await B.evaluate(() => go('more'));
  await B.click('#forceSync'); await B.settle();
  assert.equal(await B.evaluate(() => S.expenses.length), 5);
});

await t('a Sheet whose script is older still syncs, one request per change', async () => {
  const C = await T.phone({ token: ALEX, oldBackend: true });
  await C.evaluate(() => setActiveProject('p_trip'));
  const mark = C.requests.length;
  await add(C, 22, 'Bus'); await C.settle();
  assert.ok(boot().expenses.some(e => e.description === 'Bus'));
  const sent = since(C, mark);
  assert.ok(sent.includes('addExpense') && sent.includes('bootstrap'), sent.join(','));
  assert.equal(sent.filter(a => a === 'sync').length <= 1, true, 'it stops trying the new way');
  assert.equal(await C.evaluate(() => S.expenses.length), 6);
});

await T.close();
done();
