const CACHE_NAME = 'listrik-kita-pos-v35';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './logo.png',
  './logo-receipt.png',
  './manifest.json',
  './tailwind.min.js',
  './lucide.min.js',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

// Listener pesan dari aplikasi (untuk tombol "Perbarui Sekarang")
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// Install Event - Caching semua aset aplikasi
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[Service Worker] Caching app shell (v35 - Fix Modal Form & Kamera HP)...');
      return cache.addAll(ASSETS_TO_CACHE).catch(err => {
        console.warn('[Service Worker] Partial cache failure:', err);
      });
    })
  );
  self.skipWaiting();
});

// Activate Event - Hapus semua cache lama (paksa update)
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keyList) => {
      return Promise.all(
        keyList.map((key) => {
          if (key !== CACHE_NAME) {
            console.log('[Service Worker] Removing old cache:', key);
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// Fetch Event - Network-First untuk HTML/Navigasi, Cache-First untuk Aset Statis
self.addEventListener('fetch', (event) => {
  // Call ke Google Apps Script selalu live (Network Only)
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com') || event.request.url.includes('googleusercontent.com')) {
    return;
  }

  const isHtml = event.request.mode === 'navigate' || 
                 event.request.headers.get('accept')?.includes('text/html') ||
                 event.request.url.endsWith('index.html') || 
                 event.request.url.endsWith('/');

  if (isHtml) {
    // Network-First untuk HTML agar update tampilan langsung terlihat saat refresh
    event.respondWith(
      fetch(event.request)
        .then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseToCache));
          }
          return networkResponse;
        })
        .catch(() => caches.match(event.request).then(cached => cached || caches.match('./index.html')))
    );
    return;
  }

  // Cache-First untuk aset lokal lainnya (gambar, skrip JS statis)
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      const fetchPromise = fetch(event.request).then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200) {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseToCache));
        }
        return networkResponse;
      }).catch(() => cachedResponse);

      return cachedResponse || fetchPromise;
    })
  );
});
