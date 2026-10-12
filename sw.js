// App-Hülle offline verfügbar halten. Netz zuerst, damit Änderungen sofort ankommen;
// Graph/OneDrive-Anfragen laufen nicht über den Cache.
const CACHE = "tasksnap-v1";
const SHELL = [
  "./", "index.html", "style.css", "app.js", "config.js", "manifest.webmanifest",
  "vendor/msal-browser.min.js", "icon-180.png", "icon-192.png", "icon-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET" || url.origin !== location.origin) return;
  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    try {
      const response = await Promise.race([
        fetch(event.request),
        new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), 4000)),
      ]);
      if (response.ok && !url.search) cache.put(event.request, response.clone());
      return response;
    } catch {
      // ignoreSearch: Rückkehr von der Anmeldung (…/?code=…) auch offline beantworten
      return (await cache.match(event.request, { ignoreSearch: true })) || Response.error();
    }
  })());
});
