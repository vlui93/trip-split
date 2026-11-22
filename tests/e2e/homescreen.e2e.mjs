// Browser test: a link opened in the browser goes with the app to the home
// screen, and "Copy my link" / "Paste my link" covers a phone where it doesn't.
// The home-screen app keeps its own storage, so each launch here is a fresh
// phone opened at the address the icon was saved with.
// Run: node tests/e2e/homescreen.e2e.mjs
import assert from 'node:assert/strict';
import { setup, runner } from './harness.mjs';

const T = await setup();
const { t, done } = runner();
T.ok(T.api('setPeople', { people: ['Alex', 'Sam', 'Uncle'] }));
T.ok(T.api('addProject', { id: 'p_trip', name: 'Greater China', members: ['Alex', 'Sam', 'Uncle'], currencies: ['AUD'] }));
const ALEX = 'a'.repeat(32), UNCLE = 'c'.repeat(32), SAM = 'b'.repeat(32);
T.ok(T.api('createInvite', { token: ALEX, person: 'Alex', role: 'member', project_ids: [] }));
T.ok(T.api('createInvite', { token: UNCLE, person: 'Uncle', project_ids: ['p_trip'] }));
T.ok(T.api('createInvite', { token: SAM, person: 'Sam', role: 'member', project_ids: [] }));

const A = await T.phone({ token: ALEX });
const link = tok => A.evaluate(x => inviteLink(x), tok);
const UNCLE_LINK = await link(UNCLE), SAM_LINK = await link(SAM);

/** What Add to Home Screen would save: the manifest's start address, and the page's own address. */
const saved = p => p.evaluate(async () => {
  const href = document.querySelector('link[rel="manifest"]').href;
  const m = await (await fetch(href)).json();
  return { start: m.start_url, scope: m.scope, icon: m.icons[0].src, name: m.name, display: m.display, page: location.href };
});

let iconAddress;
await t('opened in the browser, the link stays in the address bar and in the manifest', async () => {
  const G = await T.phone({ link: UNCLE_LINK });
  assert.equal(await G.evaluate(() => S.guest.person), 'Uncle');
  const s = await saved(G);
  assert.match(s.page, /#join=/, 'the address bar keeps it');
  assert.equal(s.start, G.local(UNCLE_LINK).replace(/^http:\/\/127\.0\.0\.1:\d+\//, s.scope), 'the icon starts at the link');
  assert.equal(new URL(s.page).hash, new URL(s.start).hash);
  assert.deepEqual([s.name, s.display], ['Trip Split', 'standalone']);
  assert.ok(s.icon.startsWith(s.scope), 'icons still resolve from a generated manifest');
  assert.match(await G.textContent('#hsTip'), /Add to Home Screen/);
  iconAddress = s.start;
  await G.context().close();
});

await t('the home-screen icon opens already connected, with no prompt', async () => {
  let asked = 0;
  const H = await T.phone({ link: iconAddress, standalone: true, raw: true, dialogs: d => { asked++; d.accept(); } });
  await H.waitForSelector('#s-home.on'); await H.settle();
  assert.equal(await H.evaluate(() => S.guest.person), 'Uncle');
  assert.equal(asked, 0);
  assert.doesNotMatch(await H.evaluate(() => location.href), /join=/, 'the address is tidied inside the app');
  assert.equal(await H.evaluate(() => document.querySelector('link[rel="manifest"]').getAttribute('href')), 'manifest.webmanifest', 'nothing to carry from inside the app');
  assert.equal(await H.textContent('#hsTip'), '', 'no home-screen tip inside the home-screen app');

  // Later launches open at the same address. After Disconnect it must stay disconnected.
  await H.click('nav [data-go="more"]'); await H.click('#cfgClear');
  await H.goto(H.local(iconAddress)); await H.waitForSelector('#s-setup.on');
  assert.equal(await H.evaluate(() => CFG), null, 'an old icon address does not reconnect');
  // And a phone moved to another link is not pulled back to the icon's one.
  await H.click('#showManual'); await H.fill('#cfgUrl', SAM_LINK); await H.click('#cfgSave'); await H.waitForSelector('#s-home.on');
  await H.goto(H.local(iconAddress)); await H.waitForSelector('#s-home.on');
  assert.equal(await H.evaluate(() => S.whoami.person), 'Sam');
  assert.equal(asked, 0);
  await H.context().close();
});

await t('the creator\'s own phone, connected by URL and token, carries its link too', async () => {
  const s = await saved(A);
  const H = await T.phone({ link: s.start, standalone: true, raw: true });
  await H.waitForSelector('#s-home.on'); await H.settle();
  assert.equal(await H.evaluate(() => S.whoami.person), 'Alex');
  await H.context().close();
});

await t('where it doesn\'t carry over: Copy my link, then Paste my link', async () => {
  await A.click('#hsCopy');
  assert.equal(await A.evaluate(() => navigator.clipboard.readText()), await link(ALEX));
  const H = await T.phone({ standalone: true, raw: true });
  await H.waitForSelector('#s-setup.on');
  assert.equal(await H.locator('#pasteLink').isVisible(), true, 'the link comes first');
  assert.equal(await H.locator('#cfgToken').isVisible(), false, 'URL and token are tucked away');
  await H.evaluate(t => navigator.clipboard.writeText(t), await link(ALEX));
  await H.click('#pasteLink'); await H.waitForSelector('#s-home.on'); await H.settle();
  assert.equal(await H.evaluate(() => S.whoami.person), 'Alex');
  await H.context().close();
});

await t('pasting something else says so and opens the box', async () => {
  const H = await T.phone({ standalone: true, raw: true });
  await H.evaluate(() => navigator.clipboard.writeText('hello'));
  await H.click('#pasteLink');
  await H.waitForFunction(() => document.querySelector('#setupMsg').textContent.trim());
  assert.match(await H.textContent('#setupMsg'), /isn’t a Trip Split link/);
  assert.equal(await H.locator('#cfgUrl').isVisible(), true);
  await H.context().close();
});

await t('the tip can be put away, and a disconnected browser carries nothing', async () => {
  await A.click('#hsOff');
  assert.equal(await A.textContent('#hsTip'), '');
  await A.click('nav [data-go="more"]'); await A.click('#cfgClear');
  await A.waitForFunction(() => document.querySelector('link[rel="manifest"]').getAttribute('href') === 'manifest.webmanifest');
  assert.doesNotMatch(await A.evaluate(() => location.href), /join=/);
});

await T.close();
done();
