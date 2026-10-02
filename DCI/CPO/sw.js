const CACHE_NAME='cpo-supply-v2.10.1';
const APP_SHELL=[
  './',
  './index.html',
  './manifest.webmanifest',
  './assets/cpo-app-icon.svg',
  './assets/optical_fanout_assumed_dimensions.json',
  './assets/optical_fanout_cad_groups.json'
];

self.addEventListener('install',event=>{
  event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.addAll(APP_SHELL)).then(()=>self.skipWaiting()));
});

self.addEventListener('activate',event=>{
  event.waitUntil(
    caches.keys()
      .then(keys=>Promise.all(
        keys
          .filter(key=>key.startsWith('cpo-supply-') && key!==CACHE_NAME)
          .map(key=>caches.delete(key))
      ))
      .then(()=>self.clients.claim())
  );
});

self.addEventListener('fetch',event=>{
  const req=event.request;
  if(req.method!=='GET') return;
  const url=new URL(req.url);

  if(url.pathname.startsWith('/socket.io/')){
    event.respondWith(fetch(req));
    return;
  }

  if(url.pathname.startsWith('/api/')){
    event.respondWith(fetch(req,{cache:'no-store'}).catch(()=>new Response(JSON.stringify({ok:false,error:'offline'}),{status:503,headers:{'Content-Type':'application/json'}})));
    return;
  }

  if(url.origin===self.location.origin){
    const isNavigation=req.mode==='navigate' || url.pathname.endsWith('/index.html') || url.pathname==='/' || url.pathname.endsWith('/DCI/CPO/');
    if(isNavigation){
      event.respondWith(
        fetch(req,{cache:'no-store'}).then(res=>{
          if(res && res.ok){
            const clone=res.clone();
            caches.open(CACHE_NAME).then(cache=>cache.put('./index.html',clone));
          }
          return res;
        }).catch(()=>caches.match(req).then(cached=>cached||caches.match('./index.html')))
      );
      return;
    }

    event.respondWith(
      caches.match(req).then(cached=>{
        const network=fetch(req).then(res=>{
          if(res && res.ok){
            const clone=res.clone();
            caches.open(CACHE_NAME).then(cache=>cache.put(req,clone));
          }
          return res;
        }).catch(()=>cached);
        return cached || network;
      })
    );
  }
});