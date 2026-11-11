// Browser test: retiring the shared main link, what a phone still on it sees,
// and moving that phone to a personal link without losing what it had queued.
// Run: node tests/e2e/mainlink.e2e.mjs
import assert from 'node:assert/strict';
import { setup, runner } from './harness.mjs';

const T = await setup();
const { t, done } = runner();
T.ok(T.api('setPeople', { people: ['Alex', 'Sam', 'Jo'] }));
T.ok(T.api('addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo'], currencies: ['AUD'] }));
const ALEX = 'a'.repeat(32), SAM = 'b'.repeat(32);
T.ok(T.api('createInvite', { token: ALEX, person: 'Alex', role: 'member', project_ids: [] }));
T.ok(T.api('createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));
const boot = token => T.api('bootstrap', {}, token);

const M = await T.phone({});                 // still on the shared main link
const A = await T.phone({ token: ALEX });
for (const p of [M, A]) await p.evaluate(() => setActiveProject('p_trip'));
const text = (p, sel) => p.evaluate(s => document.querySelector(s).innerText.replace(/\s+/g, ' ').trim(), sel);
const add = async (p, amount, desc) => {
  await p.click('#addBtn');
  await p.fill('#exAmount', String(amount)); await p.dispatchEvent('#exAmount', 'input');
  await p.fill('#exDesc', desc); await p.dispatchEvent('#exDesc', 'input');
  await p.click('#exSave');
};

await t('Settings says which kind of link this phone is on', async () => {
  await M.click('nav [data-go="more"]');
  assert.match(await text(M, '#inviteBox'), /You’re on the shared main link/);
  assert.equal(await M.locator('#mainOff').count(), 0, 'it cannot be retired from itself');
  assert.equal(await M.locator('#meBox').isVisible(), true, 'the shared link has to ask who you are');
  await A.click('nav [data-go="more"]');
  assert.match(await text(A, '#inviteBox'), /The shared main link still works/);
  assert.match(await text(A, '#inviteBox'), /No link yet: Jo\. They would be locked out\./);
  assert.equal(await A.locator('#meBox').isVisible(), false, 'a personal link already knows');
});

await t('retiring it from a personal link is recorded under that person', async () => {
  await A.click('#mainOff'); await A.settle();
  assert.match(await text(A, '#inviteBox'), /The shared main link is retired/);
  const b = boot(ALEX);
  assert.equal(b.mainLink, 'off');
  assert.ok(b.changes.some(c => c.who === 'Alex' && c.summary === 'Retired the shared main link'));
  assert.equal(boot().ok, false);
});

await t('a phone still on it is told, keeps its changes, and carries on with its own link', async () => {
  await M.click('nav [data-go="home"]');
  await add(M, 35, 'Noodles');
  await M.waitForFunction(() => !syncing && /retired/.test(lastSyncErr));
  assert.match(await text(M, '#syncBanner'), /1 change waiting to sync\. The shared main link has been retired\. Open the app with your own link\./);
  assert.equal(await M.evaluate(() => QUEUE.length), 1, 'kept on the phone');
  assert.equal(boot(ALEX).expenses.length, 0);
  const link = await A.evaluate(tok => inviteLink(tok), SAM);
  await M.click('#useOwnLink');
  assert.equal(await M.locator('#s-setup.on').count(), 1);
  assert.equal(await M.locator('#setupBack').isVisible(), true);
  await M.fill('#cfgUrl', link); await M.click('#cfgSave');
  await M.waitForSelector('#s-home.on'); await M.settle();
  assert.equal(await M.evaluate(() => S.whoami.person), 'Sam');
  const ex = boot(ALEX).expenses.find(e => e.description === 'Noodles');
  assert.equal(ex.created_by, 'Sam', 'what was waiting went up under the new link');
  assert.equal(await M.evaluate(() => $('#syncBanner').innerHTML), '');
});

await t('it can be brought back from the same screen', async () => {
  await M.click('nav [data-go="more"]');
  await M.click('#mainOn'); await M.settle();
  assert.equal(boot().mainLink, 'on');
  assert.match(await text(M, '#inviteBox'), /The shared main link still works/);
  assert.ok(boot().changes.some(c => c.who === 'Sam' && c.summary === 'Brought back the shared main link'));
});

await T.close();
done();
