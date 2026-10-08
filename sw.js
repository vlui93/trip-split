// Keeps the app itself on the phone, so it opens with no connection or when
// GitHub Pages is slow or blocked. Only this site's own files are touched —
// Google, the rate APIs and everything else go to the network untouched; the
// app's offline queue already handles data.
//
// The app shell (page, manifest, icons) is network-first: a fresh copy when
// the site answers within a few seconds, the saved one otherwise. The offline
// receipt reader (ocr/) is large and never changes within a version, so it's
// cache-first and only stored once downloaded from Settings.
const OCR_CACHE   = 'ocr-7.0.0';   // keep in step with OCR_CACHE in index.html
const SHELL_CACHE = 'shell-v1';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon.svg',
  'icon-32-v2.png', 'icon-120-v2.png', 'icon-152-v2.png', 'icon-167-v2.png',
  'icon-180-v2.png', 'icon-192-v2.png', 'icon-512-v2.png'];
const NETWORK_WAIT = 4000;

self.addEventListener('install', e => e.waitUntil((async () => {
  const c = await caches.open(SHELL_CACHE);
  // One at a time and forgiving, so one missing icon can't block the install.
  await Promise.all(SHELL.map(u => c.add(new Request(u, { cache: 'reload' })).catch(() => {})));
  await self.skipWaiting();
})()));

self.addEventListener('activate', e => e.waitUntil((async () => {
  for (const k of await caches.keys())
    if ((k.startsWith('ocr-') && k !== OCR_CACHE) || (k.startsWith('shell-') && k !== SHELL_CACHE)) await caches.delete(k);
  await self.clients.claim();
})()));

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin) return;
  if (url.pathname.includes('/ocr/')) return e.respondWith(ocrFirst(req));
  e.respondWith(networkFirst(req));
});

async function ocrFirst(req){
  const c = await caches.open(OCR_CACHE);
  const hit = await c.match(req, { ignoreSearch: true });
  if (hit) return hit;
  const res = await fetch(req);
  if (res.ok) c.put(req, res.clone());
  return res;
}

async function networkFirst(req){
  const c = await caches.open(SHELL_CACHE);
  // Every page load is the one app page, however it was opened.
  const key = req.mode === 'navigate' ? 'index.html' : req;
  const saved = () => c.match(key, { ignoreSearch: true });

  const net = fetch(req).then(res => {
    if (res.ok && res.type === 'basic') c.put(key, res.clone());
    return res;
  });
  const slow = new Promise(r => setTimeout(r, NETWORK_WAIT, null));

  // Prefer the network, but don't wait past NETWORK_WAIT if we have a copy.
  const first = await Promise.race([net.catch(() => null), slow]);
  if (first) return first;
  const hit = await saved();
  if (hit) return hit;
  return net;   // nothing saved yet: wait it out, or fail as the browser would
}
