// Service Worker pour Groove In — gère le cache et le mode hors ligne
// Le numéro de version est synchronisé avec BUILD_VERSION du HTML.
// À chaque nouveau build, le cache est invalidé et l'app se rafraîchit.

const VERSION = "20260508-1812";
const CACHE_NAME = "groove-in-" + VERSION;

// Ressources à pré-cacher au moment de l'install
const PRECACHE_URLS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-512-maskable.png",
  "./apple-touch-icon.png",
  "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      // addAll en mode "best effort" : si une URL externe échoue, on ne bloque pas l'install
      return Promise.all(
        PRECACHE_URLS.map((url) =>
          cache.add(url).catch((err) => console.warn("Cache miss au precache:", url, err))
        )
      );
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((k) => k.startsWith("groove-in-") && k !== CACHE_NAME)
          .map((k) => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Ne jamais cacher les appels Apps Script (toujours frais, vers Google)
  if (url.hostname.includes("script.google.com") || url.hostname.includes("googleusercontent.com")) {
    return; // Le navigateur fait sa requête directe
  }

  // Stratégie network-first pour la navigation HTML — toujours essayer la version fraîche, fallback sur le cache si offline
  if (req.mode === "navigate" || (req.method === "GET" && req.headers.get("accept")?.includes("text/html"))) {
    event.respondWith(
      fetch(req)
        .then((resp) => {
          // Stocke la version la plus fraîche
          const clone = resp.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          return resp;
        })
        .catch(() => caches.match(req).then((c) => c || caches.match("./index.html")))
    );
    return;
  }

  // Stratégie cache-first pour les assets statiques (icônes, manifest, scripts CDN)
  if (req.method === "GET") {
    event.respondWith(
      caches.match(req).then((cached) => {
        if (cached) return cached;
        return fetch(req).then((resp) => {
          if (resp.ok) {
            const clone = resp.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          }
          return resp;
        });
      })
    );
  }
});

// Permet à la page d'envoyer un message "skipWaiting" pour forcer une mise à jour
self.addEventListener("message", (event) => {
  if (event.data && event.data.type === "SKIP_WAITING") self.skipWaiting();
});
