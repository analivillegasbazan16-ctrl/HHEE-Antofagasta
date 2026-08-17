/* ============================================================
   HHEE ANTOFAGASTA
   FIREBASE CLOUD MESSAGING - SERVICE WORKER
   ============================================================ */

self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(self.clients.claim());
});


/* ============================================================
   FIREBASE
   ============================================================ */

importScripts(
  "https://www.gstatic.com/firebasejs/10.13.2/firebase-app-compat.js"
);

importScripts(
  "https://www.gstatic.com/firebasejs/10.13.2/firebase-messaging-compat.js"
);


firebase.initializeApp({
  apiKey: "AIzaSyCtZny0My5plZKF3VIOFY0MgTdY8NN82ek",
  authDomain: "hhee-antofagasta.firebaseapp.com",
  projectId: "hhee-antofagasta",
  storageBucket: "hhee-antofagasta.firebasestorage.app",
  messagingSenderId: "1004191970210",
  appId: "1:1004191970210:web:c8be9ae46af57d2b7138fb",
  measurementId: "G-9PD7MCBCRS"
});


const messaging = firebase.messaging();


/* ============================================================
   MENSAJES EN SEGUNDO PLANO

   NO mostramos nosotros otra notificación.
   Firebase la mostrará automáticamente y manejará
   el enlace fcm_options.link enviado por Apps Script.
   ============================================================ */

messaging.onBackgroundMessage(payload => {
  console.log(
    "[HHEE] Notificación recibida en segundo plano:",
    payload
  );
});
