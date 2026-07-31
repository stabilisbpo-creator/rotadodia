// Service worker do Firebase Cloud Messaging.
// ⚠️ Troque os valores abaixo pelo firebaseConfig real (Firebase Console → Configurações do projeto → Geral)

importScripts("https://www.gstatic.com/firebasejs/10.12.0/firebase-app-compat.js");
importScripts("https://www.gstatic.com/firebasejs/10.12.0/firebase-messaging-compat.js");

firebase.initializeApp({
  apiKey: "AIzaSyBO__2BO5kGHGekDg9IH9uldGthJiFmdN8",
  authDomain: "rota-do-dia.firebaseapp.com",
  projectId: "rota-do-dia",
  storageBucket: "rota-do-dia.firebasestorage.app",
  messagingSenderId: "428174221981",
  appId: "1:428174221981:web:5c893bf42cd7efc24b6a74"
});

const messaging = firebase.messaging();

// dispara quando a notificação chega com o app fechado / em segundo plano
messaging.onBackgroundMessage((payload) => {
  const titulo = payload.data?.title || "Rota do Dia";
  const opcoes = {
    body: payload.data?.body || "",
    icon: "icon-192.png",
    badge: "icon-192.png"
  };
  self.registration.showNotification(titulo, opcoes);
});
