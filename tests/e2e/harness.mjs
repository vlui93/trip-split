// Shared setup for the browser tests: serves the site, runs the real Code.gs
// against the fake Sheet, and opens phone-sized pages wired to it.
import http from 'node:http';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import assert from 'node:assert/strict';
import { loadBackend } from '../fake-gas.mjs';

const require = createRequire(import.meta.url);
// CI installs playwright next to the repo; a dev box may have it elsewhere.
let pw;
try { pw = require('playwright'); }
catch { pw = require(path.join(process.env.PLAYWRIGHT_DIR || '/opt/node-tools/node_modules', 'playwright')); }

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.json': 'application/json',
  '.webmanifest': 'application/manifest+json', '.png': 'image/png', '.svg': 'image/svg+xml', '.gz': 'application/gzip', '.pdf': 'application/pdf' };
export const EXEC = 'https://script.google.com/macros/s/TEST/exec';
export const fixture = name => path.join(ROOT, 'tests', 'fixtures', name);

/** Everything a test needs. Call `await t.close()` at the end. */
export async function setup(){
  const server = http.createServer(async (req, res) => {
    const rel = decodeURIComponent(new URL(req.url, 'http://x').pathname).replace(/^\/+/, '') || 'index.html';
    const file = path.join(ROOT, rel);
    if (!file.startsWith(ROOT)) { res.writeHead(403).end(); return; }
    try { res.writeHead(200, { 'content-type': TYPES[path.extname(file)] || 'application/octet-stream' }).end(await readFile(file)); }
    catch { res.writeHead(404).end(); }
  });
  await new Promise(r => server.listen(0, '127.0.0.1', r));
  const site = `http://127.0.0.1:${server.address().port}/`;
  const { ctx: gas, book, owner } = loadBackend();
  const api = (action, payload, token) =>
    JSON.parse(gas.doPost({ postData: { contents: JSON.stringify({ token: token || owner, action, payload }) } }).s);
  const ok = r => { assert.equal(r.ok, true, r.error); return r; };
  const browser = await pw.chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
  const errors = [];

  /** A phone: its own storage, connected with `token` (or opened at `link`). */
  async function phone({ token, link, dialogs } = {}){
    const c = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2 });
    await c.route('https://script.google.com/**', r => r.fulfill({ status: 200, contentType: 'application/json',
      body: gas.doPost({ postData: { contents: r.request().postData() || '{}' } }).s }));
    const p = await c.newPage();
    p.on('pageerror', e => errors.push(e.message));
    p.on('dialog', d => (dialogs ? dialogs(d) : d.accept()));
    if (link){ await p.goto(link.replace(/^https?:\/\/[^/]+\//, site)); }
    else {
      await p.goto(site);
      await p.fill('#cfgUrl', EXEC); await p.fill('#cfgToken', token || owner); await p.click('#cfgSave');
    }
    await p.waitForSelector('#s-home.on');
    p.settle = () => p.waitForFunction(() => !syncing && QUEUE.length === 0, null, { timeout: 20000 });
    await p.settle();
    return p;
  }
  return { site, gas, book, owner, api, ok, phone, errors,
    async close(){ await browser.close(); server.close(); assert.deepEqual(errors, [], 'no page errors'); } };
}

/** Tiny test runner in the style of the unit tests. */
export function runner(){
  let n = 0;
  return { async t(name, fn){ await fn(); n++; console.log('ok', name); }, done(){ console.log(`\n${n} passed`); } };
}
