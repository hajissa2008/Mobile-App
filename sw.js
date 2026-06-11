var CACHE='maseel-v2';
var CDN='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
self.addEventListener('install',function(e){
  e.waitUntil(
    caches.open(CACHE).then(function(c){return c.addAll(['maseel-app.html',CDN])}).then(function(){return self.skipWaiting()})
  );
});
self.addEventListener('activate',function(e){
  e.waitUntil(
    Promise.all([
      caches.keys().then(function(l){return Promise.all(l.filter(function(k){return k!==CACHE}).map(function(k){return caches.delete(k)}))}),
      self.clients.claim()
    ])
  );
});
self.addEventListener('fetch',function(e){
  if(e.request.method!=='GET')return;
  if(new URL(e.request.url).hostname.indexOf('sharepoint')>-1)return;
  e.respondWith(
    caches.match(e.request).then(function(r){
      if(r){
        fetch(e.request).then(function(resp){if(resp&&resp.status===200){caches.open(CACHE).then(function(c){c.put(e.request,resp)})}}).catch(function(){});
        return r;
      }
      return fetch(e.request).then(function(resp){var c2=resp.clone();caches.open(CACHE).then(function(c){c.put(e.request,c2)});return resp}).catch(function(){return new Response('Offline',{status:503})});
    })
  );
});
