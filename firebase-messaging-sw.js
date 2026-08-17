/* ============================================================
   HHEE ANTOFAGASTA
   SERVICE WORKER - FIREBASE CLOUD MESSAGING
   ============================================================ */

const APP_BASE = "/HHEE-Antofagasta/";

/* ============================================================
   INSTALACION / ACTUALIZACION
   ============================================================ */

self.addEventListener("install", event => {
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(
    self.clients.claim()
  );
});


/* ============================================================
   CLICK EN LA NOTIFICACION

   IMPORTANTE:
   Este evento debe declararse ANTES de cargar Firebase.
   ============================================================ */

self.addEventListener("notificationclick", event => {

  event.notification.close();

  const destino =
    event.notification &&
    event.notification.data &&
    event.notification.data.url
      ? event.notification.data.url
      : APP_BASE;

  const urlDestino =
    new URL(
      destino,
      self.location.origin
    ).href;


  event.waitUntil(

    clients.matchAll({
      type: "window",
      includeUncontrolled: true
    })
    .then(async ventanas => {

      /*
        Si la aplicación HHEE ya está abierta,
        reutilizamos esa ventana.
      */

      for (const ventana of ventanas) {

        if (
          ventana.url.startsWith(
            self.location.origin + APP_BASE
          )
        ) {

          try {

            await ventana.navigate(
              urlDestino
            );

          } catch (error) {

            console.error(
              "No fue posible navegar a la solicitud:",
              error
            );
          }


          if ("focus" in ventana) {
            return ventana.focus();
          }
        }
      }


      /*
        Si HHEE no está abierta,
        se abre desde cero.
      */

      if (clients.openWindow) {

        return clients.openWindow(
          urlDestino
        );
      }


      return null;
    })
  );
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

  apiKey:
    "AIzaSyCtZny0My5plZKF3VIOFY0MgTdY8NN82ek",

  authDomain:
    "hhee-antofagasta.firebaseapp.com",

  projectId:
    "hhee-antofagasta",

  storageBucket:
    "hhee-antofagasta.firebasestorage.app",

  messagingSenderId:
    "1004191970210",

  appId:
    "1:1004191970210:web:c8be9ae46af57d2b7138fb",

  measurementId:
    "G-9PD7MCBCRS"
});


const messaging =
  firebase.messaging();


/* ============================================================
   NOTIFICACIONES EN SEGUNDO PLANO
   ============================================================ */

messaging.onBackgroundMessage(payload => {

  console.log(
    "[HHEE] Notificación recibida:",
    payload
  );


  const data =
    payload && payload.data
      ? payload.data
      : {};


  const titulo =
    data.title ||
    "Gestión de Horas Extras";


  const cuerpo =
    data.body ||
    "Tienes una nueva notificación.";


  const destino =
    data.url ||
    APP_BASE;


  const solicitudId =
    data.solicitudId ||
    "";


  const tipo =
    data.tipo ||
    "";


  const opciones = {

    body:
      cuerpo,

    tag:
      data.tag ||
      (
        solicitudId
          ? "hhee-" + solicitudId
          : "hhee-notificacion"
      ),

    /*
      Solicita mantener la notificación visible
      hasta que el usuario interactúe con ella.
    */
    requireInteraction:
      true,

    renotify:
      true,

    data: {

      url:
        destino,

      solicitudId:
        solicitudId,

      tipo:
        tipo
    }
  };


  return self.registration.showNotification(
    titulo,
    opciones
  );
});
