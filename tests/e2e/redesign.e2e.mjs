// Browser test for the layout: your own number first on Home, the one-screen
// expense form, History's search and filter pills, the Activity view, the
// floating tab bar and the project colour. Run: node tests/e2e/redesign.e2e.mjs
import assert from 'node:assert/strict';
import { setup, runner } from './harness.mjs';

const T = await setup();
const { t, done } = runner();
T.ok(T.api('setPeople', { people: ['Alex', 'Sam', 'Jo', 'Uncle'] }));
T.ok(T.api('addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Jo', 'Uncle'], currencies: ['AUD'], cities: ['Hong Kong', 'Macau'] }));
const ALEX = 'a'.repeat(32), UNCLE = 'c'.repeat(32);
T.ok(T.api('createInvite', { token: ALEX, person: 'Alex', role: 'member', project_ids: [] }));
T.ok(T.api('createInvite', { token: UNCLE, person: 'Uncle', project_ids: ['p_trip'] }));
const E = (id, desc, paid_by, aud, extra) => T.ok(T.api('addExpense', Object.assign({ id, project_id: 'p_trip', date: '2026-10-17', description: desc, paid_by,
  city: 'Hong Kong', currency: 'AUD', amount_local: aud, amount_aud: aud, split_type: 'equal', split_detail: { participants: ['Alex', 'Sam', 'Jo'] }, category: 'food' }, extra)));
E('e1', 'Dim sum', 'Sam', 90);
E('e2', 'Ferry to Macau', 'Jo', 60, { category: 'transport', city: 'Macau', date: '2026-10-21' });
E('e3', 'Hotel', 'Jo', 300, { category: 'stay', date: '2026-11-02' });
for (let i = 0; i < 9; i++) E('x' + i, 'Coffee ' + i, 'Alex', 6, { date: '2026-10-0' + (i + 1) });

const A = await T.phone({ token: ALEX });
await A.evaluate(() => setActiveProject('p_trip'));
const text = sel => A.evaluate(s => document.querySelector(s).innerText.replace(/\s+/g, ' ').trim(), sel);

await t('Home leads with your own number and how to clear it', async () => {
  // Alex paid 54, owes a third of 504 = 168: owes 114.
  const hero = await text('#heroCard');
  assert.match(hero, /^You owe A\$114\.00 Pay Jo A\$114\.00/);
  assert.match(hero, /Your share A\$168\.00 of A\$504\.00/);
  assert.equal(await A.locator('#recent .ex').count(), 8, 'recent is capped');
  assert.match(await text('#recent'), /See all 12 in History/);
  await A.click('#heroCard [data-go="settle"]');
  assert.equal(await A.locator('#s-settle.on').count(), 1);
  await A.click('nav [data-go="home"]');
});

await t('a new expense is the amount, what for, and one line per choice', async () => {
  await A.click('#addBtn');
  assert.equal(await A.evaluate(() => document.activeElement.id), 'exAmount', 'the amount is ready to type');
  assert.deepEqual(await A.evaluate(() => ['paid', 'split', 'cat', 'date', 'city', 'repeat'].map(k => document.querySelector('#fv-' + k).textContent)),
    ['You', 'Equally, everyone', 'Not set', 'Today', 'Hong Kong', 'Never']);
  assert.equal(await A.locator('#exSheet .frow.open').count(), 0, 'every line starts shut');
  await A.fill('#exAmount', '48'); await A.dispatchEvent('#exAmount', 'input');
  await A.fill('#exDesc', 'Taxi to the airport'); await A.dispatchEvent('#exDesc', 'input');
  assert.equal(await A.textContent('#fv-cat'), 'Transport', 'the category is guessed from the words');
  await A.click('[data-frowt="paid"]'); await A.click('[data-paid="Sam"]');
  assert.equal(await A.textContent('#fv-paid'), 'Sam');
  assert.equal(await A.locator('#exSheet .frow.open').count(), 0, 'picking shuts the line again');
  await A.click('[data-frowt="split"]'); await A.click('[data-part="Uncle"]'); await A.click('[data-part="Jo"]');
  assert.equal(await A.textContent('#fv-split'), 'Equally, 2 of 4');
  await A.click('[data-st="percent"]');
  assert.equal(await A.textContent('#fv-split'), 'By percentage');
  await A.click('[data-st="equal"]');
  await A.click('[data-frowt="city"]'); await A.click('[data-city="Macau"]');
  await A.click('#exSave'); await A.settle();
  const ex = T.api('bootstrap', {}).expenses.find(e => e.description === 'Taxi to the airport');
  assert.deepEqual([ex.paid_by, ex.amount_aud, ex.category, ex.city, ex.split_detail.participants.slice().sort().join()], ['Sam', 48, 'transport', 'Macau', 'Alex,Sam']);
});

await t('History: search and filter pills narrow the list, Clear resets it', async () => {
  await A.click('nav [data-go="history"]');
  const count = () => A.textContent('#histCount');
  assert.match(await count(), /^13 · /);
  await A.fill('#fSearch', 'coffee'); await A.dispatchEvent('#fSearch', 'input');
  assert.match(await count(), /^9 · A\$54\.00$/);
  await A.fill('#fSearch', 'macau 60'); await A.dispatchEvent('#fSearch', 'input');
  assert.match(await count(), /^1 · A\$60\.00$/);
  await A.fill('#fSearch', ''); await A.dispatchEvent('#fSearch', 'input');
  await A.selectOption('[data-fsel="cat"]', 'transport');
  assert.match(await count(), /^2 · A\$108\.00$/);
  assert.equal(await A.locator('.fpill.on').count(), 1);
  await A.selectOption('[data-fsel="person"]', 'Jo');
  assert.match(await count(), /^1 · A\$60\.00$/, 'Jo paid for the ferry and is not in the taxi');
  await A.selectOption('[data-fsel="city"]', 'Hong Kong');
  assert.match(await count(), /^0 · /);
  await A.click('#fClear');
  assert.match(await count(), /^13 · /);
  assert.equal(await A.locator('#fClear').count(), 0);
  await A.click('#fDatesBtn');
  await A.fill('#fFrom', '2026-11-01'); await A.dispatchEvent('#fFrom', 'change');
  assert.match(await count(), /^2 · A\$348\.00$/, 'the hotel, and the taxi added today');
  assert.match(await A.textContent('#fDatesBtn'), /1 Nov – today/);
  await A.click('#fDatesBtn');
  assert.match(await count(), /^13 · /, 'tapping the dates pill again clears them');
});

await t('changes and deleted items live under History, not Settings', async () => {
  await A.click('[data-hview="activity"]');
  assert.match(await text('#changesBox'), /Added "Taxi to the airport"/);
  await A.click('nav [data-go="more"]');
  assert.equal(await A.locator('#s-more #changesBox, #s-more #deletedBox').count(), 0);
  await A.click('#toActivity');
  assert.equal(await A.locator('#s-history.on #actWrap').isVisible(), true);
  const G = await T.phone({ token: UNCLE });
  await G.click('nav [data-go="history"]');
  assert.equal(await G.locator('[data-hview="activity"]').isVisible(), false, 'a guest has no Activity');
  assert.match(await G.evaluate(() => document.querySelector('#heroCard').innerText), /all square/);
  await G.context().close();
});

await t('the tab bar tucks in on the way down, and + works throughout', async () => {
  await A.click('nav [data-go="history"]'); await A.click('[data-hview="list"]');
  const barH = () => A.evaluate(() => document.querySelector('.nav .tabs').getBoundingClientRect().height);
  assert.equal(await barH(), 64);
  await A.mouse.wheel(0, 500); await A.waitForTimeout(450);
  assert.equal(await A.evaluate(() => document.body.classList.contains('nav-min')), true);
  assert.equal(await barH(), 50);
  await A.mouse.wheel(0, -200); await A.waitForTimeout(450);
  assert.equal(await barH(), 64);
  await A.mouse.wheel(0, 400); await A.waitForTimeout(450);
  await A.click('#addBtn'); await A.waitForSelector('#exSheet.on');
  await A.click('#exSheet [data-close="exSheet"]');
});

await t('each project has its colour, and it can be changed on this phone', async () => {
  const accent = () => A.evaluate(() => getComputedStyle(document.documentElement).getPropertyValue('--accent-l').trim());
  const before = await accent();
  assert.match(before, /^#[0-9a-f]{6}$/);
  await A.evaluate(() => openProjectEditor('p_trip'));
  await A.click('[data-peaccent="3"]'); await A.click('#peSave'); await A.settle();
  assert.equal(await accent(), '#15803d');
  assert.equal(await A.evaluate(() => PREFS.accent.p_trip), 3);
  await A.evaluate(() => setActiveProject('p_general'));
  assert.equal(await accent(), '#2563eb', 'the everyday project keeps its blue');
});

await t('pinch to zoom is allowed, and the page works without blur', async () => {
  const vp = await A.getAttribute('meta[name="viewport"]', 'content');
  assert.ok(!/user-scalable|maximum-scale/.test(vp), vp);
  assert.equal(await A.locator('meta[name="theme-color"]').count(), 2, 'one for light, one for dark');
});

await T.close();
done();
