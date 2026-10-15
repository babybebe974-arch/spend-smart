// Service worker — Spend Smart
// Page : réseau d'abord, cache en secours. Les appels Apps Script ne sont jamais mis en cache.
// « Partager → Spend Smart » (WhatsApp, Galerie) : les fichiers reçus sont gardés puis envoyés par la page.
const CACHE = 'spendsmart-v1';
const PARTAGE = 'spendsmart-partage';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE && k !== PARTAGE).map(k => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);

  if (req.method === 'POST' && url.pathname.endsWith('/partage')) {
    e.respondWith((async () => {
      const data = await req.formData();
      const fichiers = data.getAll('fichiers').filter(f => f && f.size);
      const c = await caches.open(PARTAGE);
      let i = 0;
      for (const f of fichiers) {
        await c.put('./partage/' + Date.now() + '-' + (i++), new Response(f, {
          headers: { 'Content-Type': f.type || 'image/jpeg', 'X-Nom': encodeURIComponent(f.name || 'photo') }
        }));
      }
      return Response.redirect(new URL('./?partage=' + fichiers.length, self.registration.scope).href, 303);
    })());
    return;
  }

  if (req.method !== 'GET') return;
  if (url.hostname.includes('script.google') || url.hostname.includes('googleusercontent')) return;

  e.respondWith(
    fetch(req).then(res => {
      if (res && res.ok && res.type === 'basic') {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy));
      }
      return res;
    }).catch(() => caches.match(req).then(hit =>
      hit || (req.mode === 'navigate' ? caches.match('./index.html') : undefined)))
  );
});
