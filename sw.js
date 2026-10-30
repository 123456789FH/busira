const PREFIX='basirah-static-'+new URL(self.registration.scope).pathname+':';
const CACHE=PREFIX+'v1.0.0';
const FILES=['./','./index.html','./style.css','./app.js','./omr.js','./manifest.webmanifest','./assets/icon.svg','./assets/arabic-regular.ttf','./assets/arabic-bold.ttf'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES)).then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith(PREFIX)&&k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{const u=new URL(e.request.url),base=new URL(self.registration.scope);if(e.request.method!=='GET'||u.origin!==base.origin||!u.pathname.startsWith(base.pathname))return;
  if(e.request.mode==='navigate'){e.respondWith(fetch(e.request).then(r=>{if(r.ok){const copy=r.clone();caches.open(CACHE).then(c=>c.put('./index.html',copy));}return r;}).catch(()=>caches.match('./index.html')));return;}
  if(!FILES.some(f=>new URL(f,base).pathname===u.pathname))return;
  e.respondWith(caches.match(e.request).then(cached=>cached||fetch(e.request)));
});
