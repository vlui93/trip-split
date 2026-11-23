// Browser test: the app opens from the phone's saved copy when the site is
// unreachable or answers with an error page (GitHub Pages switched off).
// Run: node tests/e2e/offline-shell.e2e.mjs
import assert from 'node:assert/strict';
import { setup, runner } from './harness.mjs';

const T = await setup();
const { t, done } = runner();
T.ok(T.api('setPeople', { people: ['Alex', 'Sam'] }));
T.ok(T.api('addExpense', { id: 'e1', project_id: 'p_general', date: '2026-10-17', description: 'Groceries', paid_by: 'Alex',
  currency: 'AUD', amount_local: 40, amount_aud: 40, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam'] } }));

const P = await T.phone({});
await P.evaluate(() => navigator.serviceWorker.ready);
await P.reload(); await P.waitForSelector('#s-home.on');
await P.waitForFunction(() => !!navigator.serviceWorker.controller);
await P.waitForFunction(async () => !!(await (await caches.open('shell-v3')).match('index.html')));

await t('with the site answering 404, the saved app opens with its data', async () => {
  T.siteDown(true);
  await P.reload(); await P.waitForSelector('#s-home.on', { timeout: 15000 });
  assert.equal(await P.evaluate(() => S.expenses.some(e => e.description === 'Groceries')), true);
  assert.doesNotMatch(await P.content(), /There isn't a GitHub Pages site here/);
});

await t('and the saved copy is not overwritten by the error page', async () => {
  const cached = await P.evaluate(async () => (await (await caches.open('shell-v3')).match('index.html')).text());
  assert.match(cached, /<title>Trip Split<\/title>/);
});

await t('when the site is back, it is used again', async () => {
  T.siteDown(false);
  await P.reload(); await P.waitForSelector('#s-home.on');
  assert.equal(await P.evaluate(() => document.title), 'Trip Split');
});

await T.close();
done();
