const CACHE_NAME = 'makspilot-cache-v1.0.0';

const urlsToCache = [
  '/',
  '/index.html',
  '/style.css',
  '/javascript/script.js',
  '/javascript/functions.js',
  '/javascript/redirect.js',
  '/img/main.png',
  '/img/maxpilot.png',
  '/img/python.svg',
  '/img/shiba.svg',
  '/img/nerd.svg',
  '/img/duck.svg',
  '/img/linkedin.svg',
  '/img/blog.svg',
  '/favicon.ico'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Открыт кэш');
        return cache.addAll(urlsToCache);
      })
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        if (response) {
          return response;
        }
        return fetch(event.request);
      }
    )
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
});