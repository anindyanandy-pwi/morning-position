// The redesign briefly registered a service worker at this scope (13.11). Remove it so the portal loads normally.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil((async () => { await self.registration.unregister(); const cs = await self.clients.matchAll(); cs.forEach(c => c.navigate(c.url)); })()));
