/* Appli installable + hors ligne : garde une copie du code. Les données (OneDrive) ne sont jamais mises en cache ici. */
const V = "yousdam-v44";
const CODE = ["./", "index.html", "moteur.js", "app.js", "local.js", "config.js", "manifest.json", "icon-192.png", "icon-512.png",
  "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"];
self.addEventListener("install", e => { e.waitUntil(caches.open(V).then(c => c.addAll(CODE)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (u.pathname.endsWith("version.json") || u.pathname.includes("/__natif/")) return;   // toujours la vraie version en ligne
  if (e.request.method !== "GET" || u.hostname.endsWith("microsoft.com") || u.hostname.endsWith("microsoftonline.com") || u.hostname.includes("live.com") || u.hostname.includes("sharepoint")) return;
  // le code : d'abord le réseau (toujours la dernière version), sinon la copie
  e.respondWith(fetch(e.request, {cache: "no-cache"}).then(r => { const c = r.clone(); caches.open(V).then(x => x.put(e.request, c)); return r; }).catch(() => caches.match(e.request)));
});
