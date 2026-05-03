// Service Worker minimal untuk mengaktifkan fitur PWA (Installable)
const CACHE_NAME = 'absensi-v1';

self.addEventListener('install', (event) => {
  console.log('Service Worker: Installed');
});

self.addEventListener('activate', (event) => {
  console.log('Service Worker: Activated');
});

self.addEventListener('fetch', (event) => {
  // Biarkan browser menangani permintaan secara normal
  // Ini hanya untuk memenuhi syarat 'installable' di Chrome/Edge
});
