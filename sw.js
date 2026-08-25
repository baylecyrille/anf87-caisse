const VER='caisse-v3'; // version incrémentée : purge l'ancien cache qui gardait les
// anciennes synchronisations Google Sheets figées pour toujours.
const CDN=[
  'https://unpkg.com/react@18.2.0/umd/react.production.min.js',
  'https://unpkg.com/react-dom@18.2.0/umd/react-dom.production.min.js',
  'https://unpkg.com/@babel/standalone@7.23.5/babel.min.js',
  'https://unpkg.com/@zxing/library@0.20.0/umd/index.min.js'
];
self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(VER).then(c=>c.addAll(['./','./index.html','./manifest.json',...CDN])));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==VER).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{
  const url=new URL(e.request.url);
 
  // IMPORTANT : on ne met JAMAIS en cache les appels vers Google (Apps Script /
  // Google Sheets). Sans cette exclusion, la toute première synchronisation réussie
  // restait figée en cache indéfiniment, et chaque "actualisation" suivante resservait
  // ces données périmées au lieu d'aller chercher les vraies données à jour — c'était
  // la cause du "ça enregistre bien dans la feuille, mais la synchro remet les
  // anciennes infos".
  const isGoogleApi = url.hostname.endsWith('script.google.com') || url.hostname.endsWith('googleusercontent.com');
  if(isGoogleApi){ e.respondWith(fetch(e.request)); return; }
 
  if(url.pathname.endsWith('index.html')||url.pathname==='/'){
    e.respondWith(fetch(e.request).then(r=>{const c=r.clone();caches.open(VER).then(ca=>ca.put(e.request,c));return r;}).catch(()=>caches.match(e.request)));return;}
  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(res=>{const c=res.clone();caches.open(VER).then(ca=>ca.put(e.request,c));return res;})));
});
 
