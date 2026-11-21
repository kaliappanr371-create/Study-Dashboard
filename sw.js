/* Study Dashboard service worker. Version bcaa152983 is derived from the app files, so any code change gives a new cache. */
const CACHE='study-dash-bcaa152983';
const PRECACHE=["./", "index.html", "manifest.webmanifest", "assets/style.css?v=bcaa152983", "assets/app.js?v=bcaa152983", "assets/seed.js?v=bcaa152983", "icons/icon-192.png", "icons/icon-512.png", "icons/maskable-512.png", "icons/apple-touch-180.png", "icons/favicon.ico"];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(PRECACHE.map(u=>new Request(u,{cache:'reload'})))))});
self.addEventListener('activate',e=>{e.waitUntil((async()=>{
  for(const k of await caches.keys())if(k.startsWith('study-dash-')&&k!==CACHE)await caches.delete(k);
  await self.clients.claim()})())});
self.addEventListener('message',e=>{if(e.data==='skip-waiting')self.skipWaiting()});
self.addEventListener('fetch',e=>{
  const r=e.request;if(r.method!=='GET')return;
  const u=new URL(r.url);if(u.origin!==location.origin)return;
  e.respondWith((async()=>{
    const c=await caches.open(CACHE);
    const hit=await c.match(r,{ignoreSearch:r.mode==='navigate'});if(hit)return hit;
    if(r.mode==='navigate')return (await c.match('index.html'))||Response.error();
    try{return await fetch(r)}catch(_){return Response.error()}
  })())});
