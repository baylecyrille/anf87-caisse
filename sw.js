const VER='caisse-v2';
const CDN=['https://unpkg.com/react@18.2.0/umd/react.production.min.js','https://unpkg.com/react-dom@18.2.0/umd/react-dom.production.min.js','https://unpkg.com/@babel/standalone@7.23.5/babel.min.js'];
self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(VER).then(c=>c.addAll(['./','./index.html','./manifest.json',...CDN])));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==VER).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{const url=new URL(e.request.url);
  if(url.pathname.endsWith('index.html')||url.pathname==='/'){
    e.respondWith(fetch(e.request).then(r=>{const c=r.clone();caches.open(VER).then(ca=>ca.put(e.request,c));return r;}).catch(()=>caches.match(e.request)));return;}
  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(res=>{const c=res.clone();caches.open(VER).then(ca=>ca.put(e.request,c));return res;})));
});
