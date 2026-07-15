/*
 * BudgetIQ Service Worker
 * Enables offline use + installability (PWA).
 * Strategy:
 *   - HTML pages : cache-first + background refresh (instant after first load)
 *   - Everything : cache-first  (fast load, works offline after first visit)
 * Bump CACHE_VERSION whenever you change the template to push updates.
 */
const CACHE_VERSION = 'budgetiq-v59-github-pages';
const APP_BASE = '/budgetiq/';

// App shell — precached on install so the app opens offline right away.
const CORE_ASSETS = [
  `${APP_BASE}splash.html`,
  `${APP_BASE}dashboard.html`,
  `${APP_BASE}cart.html`,
  `${APP_BASE}plan.html`,
  `${APP_BASE}profile.html`,
  `${APP_BASE}history.html`,
  `${APP_BASE}css/tailwind.min.css`,
  `${APP_BASE}css/tailwind.min.css?v=1`,
  `${APP_BASE}css/style.css?v=4`,
  `${APP_BASE}js/budgetiq-core.js?v=59`,
  `${APP_BASE}fonts/material-icons.woff2`,
  `${APP_BASE}img/logo.png`,
  `${APP_BASE}img/budgetiq-logo.png`,
  `${APP_BASE}icons/apple-touch-icon.png`,
  `${APP_BASE}icons/icon-192.png`,
  `${APP_BASE}icons/icon-512.png`,
  `${APP_BASE}icons/icon-maskable-512.png`,
  `${APP_BASE}icons/favicon-32.png`,
  `${APP_BASE}manifest.webmanifest`
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) =>
      // addAll fails hard if any file 404s; add individually so one miss
      // never breaks the whole install.
      Promise.all(
        CORE_ASSETS.map((url) =>
          cache.add(url).catch((err) => console.warn('[sw] skip precache', url, err))
        )
      )
    )
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const request = event.request;

  // Only handle GET; let the browser deal with the rest.
  if (request.method !== 'GET') return;

  const isHTML =
    request.mode === 'navigate' ||
    (request.headers.get('accept') || '').includes('text/html');

  if (isHTML) {
    // iOS Safari cannot use a redirected Response returned by a service worker.
    // Refetch the final same-origin URL and cache only non-redirected pages.
    const networkUpdate = caches.open(CACHE_VERSION).then(async (cache) => {
      let response = await fetch(request);

      if (response.redirected && response.url) {
        const finalURL = new URL(response.url);
        if (finalURL.origin === self.location.origin) {
          response = await fetch(finalURL.pathname + finalURL.search, {
            credentials: 'same-origin'
          });
        }
      }

      if (response.ok && !response.redirected && response.type !== 'opaqueredirect') {
        await cache.put(request, response.clone());
      }

      return response;
    });

    event.waitUntil(networkUpdate.then(() => undefined).catch(() => undefined));
    event.respondWith(
      caches.open(CACHE_VERSION)
        .then(async (cache) => {
          const cached = await cache.match(request, { ignoreSearch: true });
          if (cached && !cached.redirected && cached.type !== 'opaqueredirect') return cached;
          if (cached) await cache.delete(request);
          return networkUpdate;
        })
        .catch(async () => {
          const fallback = await caches.match(`${APP_BASE}splash.html`);
          return fallback || Response.error();
        })
    );
    return;
  }

  // Cache-first for assets (CSS/JS/img/fonts, incl. cross-origin CDN files).
  event.respondWith(
    caches.match(request).then((cached) => {
      if (cached) return cached;
      return fetch(request)
        .then((response) => {
          if (response.ok && !response.redirected && response.type !== 'opaqueredirect') {
            const copy = response.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(request, copy)).catch(() => {});
          }
          return response;
        })
        .catch(() => cached);
    })
  );
});
