// Офлайн-доступ: при наличии сети всегда берём свежую версию, без сети — из кеша.
const C='npa-v2';
self.addEventListener('install',e=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(
  caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',e=>{
  const url=new URL(e.request.url);
  if(e.request.method!=='GET'||url.origin!==location.origin) return;
  const key=url.origin+url.pathname;               // без ?query
  e.respondWith(
    fetch(e.request,{cache:'no-cache'})              // перепроверка по ETag — дёшево
      .then(r=>{ if(r.ok){const c=r.clone(); caches.open(C).then(ca=>ca.put(key,c))} return r })
      .catch(()=>caches.open(C).then(ca=>ca.match(key)).then(h=>h||caches.match(e.request,{ignoreSearch:true})))
  );
});
