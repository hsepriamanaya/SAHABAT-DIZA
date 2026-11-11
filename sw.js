const C="sahabat-v1";
self.addEventListener("install",e=>e.waitUntil(caches.open(C).then(c=>c.addAll(["./","index.html"]))));
self.addEventListener("fetch",e=>{if(e.request.method==="GET"&&new URL(e.request.url).origin===location.origin)e.respondWith(fetch(e.request).catch(()=>caches.match(e.request)))});
