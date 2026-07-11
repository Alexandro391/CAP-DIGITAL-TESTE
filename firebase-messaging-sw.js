// firebase-messaging-sw.js
// Service Worker do Firebase Cloud Messaging — Painel CAP Digital (PMC-FMC-PUSH)
// IMPORTANTE: este arquivo deve ficar na RAIZ do repositório
// (mesmo nível do index.html / dashboard.html), nunca dentro de subpasta.

importScripts('https://www.gstatic.com/firebasejs/12.16.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/12.16.0/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: "AIzaSyA_HJfDMyPEYewfbGgQIJ90LmUvYK36OKE",
  authDomain: "pmc-fmc-push.firebaseapp.com",
  projectId: "pmc-fmc-push",
  storageBucket: "pmc-fmc-push.firebasestorage.app",
  messagingSenderId: "1020191800841",
  appId: "1:1020191800841:web:ffa56e8854c6c72fad82d8"
});

const messaging = firebase.messaging();

// Notificação recebida com o app FECHADO ou em segundo plano
messaging.onBackgroundMessage((payload) => {
  console.log('[SW] Notificação em background:', payload);

  const titulo = payload.notification?.title || 'Novo chamado CAP Digital';
  const opcoes = {
    body: payload.notification?.body || 'Um novo chamado foi registrado.',
    icon: '/CAP-DIGITAL-TESTE/icon-192.png',
    badge: '/CAP-DIGITAL-TESTE/icon-192.png',
    vibrate: [200, 100, 200],
    tag: 'cap-digital-chamado', // evita empilhar notificações repetidas do mesmo tipo
    data: payload.data || {}
  };

  self.registration.showNotification(titulo, opcoes);
});

// Ao clicar na notificação, abre ou foca a aba do painel
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window' }).then((clientList) => {
      for (const client of clientList) {
        if (client.url.includes('dashboard.html') && 'focus' in client) {
          return client.focus();
        }
      }
      if (clients.openWindow) {
        return clients.openWindow('/CAP-DIGITAL-TESTE/dashboard.html');
      }
    })
  );
});
