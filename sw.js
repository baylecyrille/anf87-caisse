const VER='caisse-v6'; // v6 : installation résiliente (voir plus bas) — corrige le cas où
// l'appli devenait TOTALEMENT indisponible hors connexion ("Vous êtes hors connexion"
// affiché par le navigateur lui-même à la place de l'appli).
const CDN=[
  'https://unpkg.com/react@18.2.0/umd/react.production.min.js',
  'https://unpkg.com/react-dom@18.2.0/umd/react-dom.production.min.js',
  'https://unpkg.com/@babel/standalone@7.23.5/babel.min.js',
  'https://unpkg.com/@zxing/library@0.20.0/umd/index.min.js',
  'https://cdn.jsdelivr.net/npm/qrcode@1.5.3/build/qrcode.min.js',
  'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js'
];
self.addEventListener('install',e=>{
  self.skipWaiting();
  // IMPORTANT : jamais caches.open(VER).then(c=>c.addAll([...])) — addAll() est TOUT OU
  // RIEN : si un seul des 6 fichiers (page comprise) échoue à se charger pendant cette
  // installation (accroc wifi, CDN temporairement lent, coupure d'une seconde...),
  // l'INSTALLATION ENTIÈRE échoue et RIEN n'est mis en cache, pas même la page elle-même.
  // La prochaine fois que l'appareil est hors connexion, il n'y a alors plus rien en
  // réserve : c'est le navigateur qui affiche sa page générique "Vous êtes hors
  // connexion" à la place de l'appli, qui n'a jamais la chance de démarrer son propre
  // code. En mettant chaque fichier en cache indépendamment (et en avalant l'erreur
  // d'un fichier en échec plutôt que de la laisser remonter), un seul CDN capricieux ne
  // prive plus jamais l'appli de tout ce qui a pu, par ailleurs, être mis en cache.
  e.waitUntil(caches.open(VER).then(c=>Promise.all(
    ['./','./index.html','./manifest.json',...CDN].map(url=>c.add(url).catch(()=>{}))
  )));
});
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
