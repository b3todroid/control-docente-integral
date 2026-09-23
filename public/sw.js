/* ============================================================
   SERVICE WORKER
   Guarda la aplicación en el dispositivo para que abra sin internet.
   No toca los datos escolares: esos viven en IndexedDB y nunca se
   borran desde aquí.
   ============================================================ */

const VERSION = "cdi-v2";
const CACHE = `${VERSION}-app`;

// Al instalar una versión nueva, entra de inmediato.
self.addEventListener("install", (e) => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then((c) =>
      c.addAll(["./", "./index.html", "./manifest.webmanifest", "./icono-192.png", "./icono-512.png"]).catch(() => {})
    )
  );
});

// Al activarse, limpia solo cachés de versiones anteriores de la app.
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches
      .keys()
      .then((llaves) => Promise.all(llaves.filter((k) => k.startsWith("cdi-") && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // no intercepta recursos externos

  // Navegación: intenta la red y, si no hay, entrega la app guardada.
  if (req.mode === "navigate") {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copia = res.clone();
          caches.open(CACHE).then((c) => c.put("./index.html", copia)).catch(() => {});
          return res;
        })
        .catch(() => caches.match("./index.html").then((r) => r || caches.match("./")))
    );
    return;
  }

  // Archivos de la app: primero lo guardado, y se actualiza en segundo plano.
  e.respondWith(
    caches.match(req).then((guardado) => {
      const red = fetch(req)
        .then((res) => {
          if (res && res.status === 200 && res.type === "basic") {
            const copia = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copia)).catch(() => {});
          }
          return res;
        })
        .catch(() => guardado);
      return guardado || red;
    })
  );
});
