// Service worker mínimo: permite instalar la app y deja el dashboard
// disponible (última versión cargada) aunque no haya señal al abrirlo.
const CACHE_NAME = 'sar-corani-v1';
const ARCHIVOS_APP = ['./sistema_sar_corani.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (evento) => {
  evento.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ARCHIVOS_APP))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (evento) => {
  evento.waitUntil(
    caches.keys().then((nombres) =>
      Promise.all(nombres.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

// Estrategia: intenta la red primero (para tener datos frescos del dashboard);
// si no hay conexión, sirve la última copia guardada en caché.
self.addEventListener('fetch', (evento) => {
  evento.respondWith(
    fetch(evento.request)
      .then((respuesta) => {
        const copia = respuesta.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(evento.request, copia));
        return respuesta;
      })
      .catch(() => caches.match(evento.request))
  );
});
