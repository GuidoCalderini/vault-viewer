// Service worker minimal : réseau d'abord, cache en secours (fichiers de la page seulement).
// Si le réseau ne répond pas en 4 s (réseau faible), on sert la copie gardée : le chargeur s'ouvre
// quand même et peut proposer le hors ligne. Ne met jamais en cache les réponses de Microsoft ni de OneDrive.
const CACHE = "vv-chargeur-4";
const FICHIERS = ["./", "index.html", "loader.js", "graphfs.js", "config.js", "essai.html", "essai.js", "lib/msal-browser.min.js", "manifest.webmanifest", "icons/icon-180.png", "icons/icon-192.png", "icons/icon-512.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FICHIERS)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== self.location.origin) return;
  const garde = () => caches.match(e.request, { ignoreSearch: true });
  const reseau = fetch(e.request).then((r) => {
    if (r.ok) { const copie = r.clone(); caches.open(CACHE).then((c) => c.put(e.request, copie)); }
    return r;
  });
  const delai = new Promise((res) => setTimeout(() => res(null), 4000));
  e.respondWith(
    Promise.race([reseau, delai])
      .then((r) => r || garde().then((c) => c || reseau))
      .catch(() => garde())
  );
});
