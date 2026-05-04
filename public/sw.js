self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('fetch', (event) => {
  // Minimal passthrough fetch listener for PWA compliance
  event.respondWith(fetch(event.request));
});
