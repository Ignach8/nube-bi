// Service worker: deja la app disponible sin conexión y abre al instante.
const VERSION = "nubebi-v2";
const SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./config.js",
  "./worker.js",
  "./manifest.webmanifest",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];
const CDN_HOSTS = ["cdnjs.cloudflare.com"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL.map((u) => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);

  // Librerías del CDN: versión fija, se guardan para siempre.
  if (CDN_HOSTS.includes(url.hostname)) {
    event.respondWith(
      caches.match(req).then((hit) => hit || fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(VERSION).then((c) => c.put(req, copy));
        return res;
      }))
    );
    return;
  }

  // Archivos propios: primero la red (siempre la última versión), caché si no hay conexión.
  if (url.origin === self.location.origin) {
    event.respondWith(
      // "no-cache" obliga a preguntar al servidor si hay versión nueva (GitHub Pages guarda 10 min).
      fetch(req, { cache: "no-cache" }).then((res) => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(VERSION).then((c) => c.put(req, copy));
        }
        return res;
      }).catch(() => caches.match(req, { ignoreSearch: true }))
    );
  }
  // Google APIs y login: siempre por red (no se tocan).
});
