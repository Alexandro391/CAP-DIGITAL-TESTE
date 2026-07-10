/* ═══════════════════════════════════════════════════════════
   Service Worker — PMC · Painel de Monitoramento de Chamados
   Cache só do "shell" do app (HTML/manifest/ícone).
   Os dados dos chamados (Google Sheets / Apps Script) NUNCA
   são cacheados aqui — sempre buscados direto da rede, senão
   o painel mostraria chamados desatualizados.
   ═══════════════════════════════════════════════════════════ */

const CACHE_VERSION = '2026-07-10-2'; // troque essa data ao publicar uma nova versão do painel
const CACHE_NAME = 'pmc-shell-' + CACHE_VERSION;

const APP_SHELL = [
  './painel-pmc.html',
  './manifest-pmc.json',
  './icon-192.png'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(
        names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const url = event.request.url;

  // Nunca intercepta/cacheia dados ao vivo: planilha (gviz) e Apps Script
  if (url.includes('docs.google.com') || url.includes('script.google.com')) {
    return; // deixa passar direto pra rede
  }

  // Shell do app: network-first, com fallback pro cache se ficar offline
  event.respondWith(
    fetch(event.request)
      .then((res) => {
        const resClone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClone)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
