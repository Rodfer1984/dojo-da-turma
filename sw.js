/* Service Worker — Dojo da Turma
   Deixa o app abrir offline. A sincronização com a planilha (Google Apps Script)
   NÃO é interceptada: requisições POST e de outros domínios passam direto pela rede. */
const CACHE = "dojo-turma-v4";
const ARQUIVOS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-512-maskable.png"
];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQUIVOS)));
  self.skipWaiting();
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  // Só intercepta GET do próprio site — o envio para a planilha (POST, script.google.com) passa direto.
  if (e.request.method !== "GET") return;
  const url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return;

  // Estratégia: rede primeiro (pega atualizações), cache como reserva offline.
  e.respondWith(
    fetch(e.request)
      .then(resp => {
        const copia = resp.clone();
        caches.open(CACHE).then(c => c.put(e.request, copia));
        return resp;
      })
      .catch(() => caches.match(e.request).then(r => r || caches.match("./index.html")))
  );
});
