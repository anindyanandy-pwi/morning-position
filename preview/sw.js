/* §25.4 PWA service worker. NETWORK-FIRST for the app shell (index.html and the app's own scripts/CSS) so a fix
 * reaches the office on the very next open; the cache is only the offline fallback. The report the viewer opened is
 * cached per signed-in user by the page (caches 'amp-report'), never other offices' data; sign-out clears it.
 * CACHE_VERSION is replaced with the commit by publish-preview.sh so each publish gets a fresh shell cache. */
const CACHE_VERSION = 'dev';
const SHELL = 'amp-shell-' + CACHE_VERSION;
const SHELL_FILES = ['./', 'index.html', 'preview.css', 'style.css', 'app.js', 'reader.js', 'motion.js',
  'config.js', 'docx-compat.js', 'report-docx.js', 'docx-build.js', 'ir-logo-256.png', 'cover-frame.png', 'cover-seal.png', 'cover-rule.png', 'cover-contents.png',
  'vendor/docx-9.8.1.iife.js', 'vendor/jszip-3.10.1.min.js',
  'vendor/docx-preview-0.3.5.min.js', 'manifest.webmanifest', 'icon-180.png', 'icon-192.png'];

self.addEventListener('install', e => { e.waitUntil(caches.open(SHELL).then(c => c.addAll(SHELL_FILES.map(u => new Request(u, { cache: 'reload' }))).catch(() => {})).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== SHELL && k !== 'amp-report').map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('message', e => { if (e.data === 'clear') caches.keys().then(ks => ks.forEach(k => caches.delete(k))); });

self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;   // never the API POSTs (no other office's data)
  const isShell = u.pathname.endsWith('/') || SHELL_FILES.some(f => u.pathname.endsWith('/' + f.replace('./', '')));
  if (!isShell) return;
  // network-first: try the network, update the cache, fall back to cache only when offline
  e.respondWith(
    fetch(e.request).then(r => { const copy = r.clone(); caches.open(SHELL).then(c => c.put(e.request, copy)); return r; })
      .catch(() => caches.match(e.request).then(c => c || caches.match('index.html')))
  );
});
