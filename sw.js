const CACHE_NAME = 'makspilot_v0.6.40';

const urlsToCache = [
  '/style.css',
  '/javascript/script.js',
  '/javascript/functions.js',
  '/javascript/redirect.js',
  '/img/main.webp',
  '/img/main_mobile.webp',
  '/img/makspilot.svg',
  '/img/python.svg',
  '/img/shiba.svg',
  '/img/nerd.svg',
  '/img/duck.svg',
  '/img/linkedin.svg',
  '/img/blog.svg',
  '/favicon.ico'
];

const cacheFirstStrategy = async (request) => {
  if (request.url.startsWith('chrome-extension://') || request.url.startsWith('moz-extension://')) {
    return fetch(request);
  }
  
  const cachedResponse = await caches.match(request);
  if (cachedResponse) {
    return cachedResponse;
  }

  try {
    const networkResponse = await fetch(request);
    
    if (networkResponse.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, networkResponse.clone());
    }
    
    return networkResponse;
  } catch (error) {
    console.error('Fetch failed:', error);
    throw error;
  }
};

const networkFirstStrategy = async (request) => {
  if (request.url.startsWith('chrome-extension://') || request.url.startsWith('moz-extension://')) {
    return fetch(request);
  }
  
  try {
    const networkResponse = await fetch(request);
    
    if (networkResponse.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, networkResponse.clone());
    }
    
    return networkResponse;
  } catch (error) {
    const cachedResponse = await caches.match(request);
    if (cachedResponse) {
      return cachedResponse;
    }
    
    console.error('Network and cache failed:', error);
    throw error;
  }
};

const getCacheStrategy = (request) => {
  const url = new URL(request.url);
  
  // Allow external libraries (DuckDB, etc.) to bypass cache
  if (url.hostname.includes('jsdelivr.net') || 
      url.hostname.includes('unpkg.com') ||
      url.pathname.includes('duckdb') ||
      url.pathname.includes('duckdb-wasm.js')) {
    return networkFirstStrategy;
  }
  
  if (request.destination === 'document' || url.pathname === '/' || url.pathname.endsWith('.html')) {
    return networkFirstStrategy;
  }
  
  return cacheFirstStrategy;
};

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('A cache opened');
        return cache.addAll(urlsToCache);
      })
  );
  self.skipWaiting();
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') {
    return;
  }
  
  const strategy = getCacheStrategy(event.request);
  
  event.respondWith(strategy(event.request));
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
  return self.clients.claim();
});