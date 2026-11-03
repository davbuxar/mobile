const C='coupon-v1',F=['./','index.html','manifest.json'];
self.addEventListener('install',e=>e.waitUntil(caches.open(C).then(c=>c.addAll(F))));
self.addEventListener('fetch',e=>{if(e.request.url.includes('script.google.com')||e.request.url.includes('googleusercontent'))return;
 e.respondWith(fetch(e.request).catch(()=>caches.match(e.request)))});
