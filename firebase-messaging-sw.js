/* ============================================================
   HHEE ANTOFAGASTA
   SERVICE WORKER UNIFICADO

   - Android / PC: Firebase Cloud Messaging
   - iPhone / iPad: Web Push estándar
   - Apertura de solicitud desde notificación
   ============================================================ */

const APP_BASE =
  "/HHEE-Antofagasta/";

const APP_URL =
  "https://analivillegasbazan16-ctrl.github.io/HHEE-Antofagasta/";


/* ============================================================
   DETECCIÓN DEL ENTORNO
   ============================================================ */

const USER_AGENT =
  self.navigator?.userAgent
  ||
  "";

const ES_IOS =
  /iPhone|iPad|iPod/i.test(
    USER_AGENT
  )
  ||
  /Macintosh/i.test(
    USER_AGENT
  )
  &&
  /Mobile/i.test(
    USER_AGENT
  );


/* ============================================================
   INSTALACIÓN
   ============================================================ */

self.addEventListener(
  "install",
  () => {
    self.skipWaiting();
  }
);


/* ============================================================
   ACTIVACIÓN
   ============================================================ */

self.addEventListener(
  "activate",
  event => {
    event.waitUntil(
      self.clients.claim()
    );
  }
);


/* ============================================================
   OBTENER URL DESDE UNA NOTIFICACIÓN
   ============================================================ */

function obtenerUrlNotificacion(
  notification
) {

  const data =
    notification?.data
    ||
    {};


  if (data.url) {
    return String(
      data.url
    );
  }


  const fcm =
    data.FCM_MSG
    ||
    data.fcmMessage
    ||
    {};


  if (fcm?.data?.url) {
    return String(
      fcm.data.url
    );
  }


  if (fcm?.fcmOptions?.link) {
    return String(
      fcm.fcmOptions.link
    );
  }


  if (fcm?.fcm_options?.link) {
    return String(
      fcm.fcm_options.link
    );
  }


  return APP_URL;
}


/* ============================================================
   CLICK EN NOTIFICACIÓN
   ============================================================ */

self.addEventListener(
  "notificationclick",
  event => {

    event.notification.close();


    const destino =
      obtenerUrlNotificacion(
        event.notification
      );


    let urlDestino;


    try {
      urlDestino =
        new URL(
          destino,
          self.location.origin
        ).href;
    }
    catch {
      urlDestino =
        APP_URL;
    }


    event.waitUntil(
      self.clients
        .matchAll({
          type:
            "window",

          includeUncontrolled:
            true
        })
        .then(
          async ventanas => {

            for (
              const ventana of ventanas
            ) {

              if (
                ventana.url.startsWith(
                  self.location.origin
                  +
                  APP_BASE
                )
              ) {

                try {
                  await ventana.navigate(
                    urlDestino
                  );
                }
                catch (error) {
                  console.warn(
                    "[HHEE] No se pudo navegar la ventana existente:",
                    error
                  );
                }


                try {
                  return await ventana.focus();
                }
                catch {
                  return ventana;
                }
              }
            }


            if (
              self.clients.openWindow
            ) {
              return self.clients.openWindow(
                urlDestino
              );
            }


            return null;
          }
        )
    );
  }
);


/* ============================================================
   WEB PUSH ESTÁNDAR
   iPhone / iPad
   ============================================================ */

self.addEventListener(
  "push",
  event => {

    let datos =
      null;


    try {
      datos =
        event.data
          ?
            event.data.json()
          :
            null;
    }
    catch {
      datos =
        null;
    }


    /*
      Solo procesamos aquí los mensajes Web Push propios.
      Los mensajes de Firebase en Android/PC quedan a cargo
      de Firebase Messaging.
    */

    const esWebPushHhee =
      datos
      &&
      typeof datos ===
        "object"
      &&
      typeof datos.title ===
        "string"
      &&
      typeof datos.url ===
        "string"
      &&
      !datos.from
      &&
      !datos.notification;


    if (!esWebPushHhee) {
      return;
    }


    const titulo =
      datos.title
      ||
      "Gestión de Horas Extras";


    const cuerpo =
      datos.body
      ||
      "Tienes una nueva notificación.";


    const url =
      datos.url
      ||
      APP_URL;


    const tag =
      datos.tag
      ||
      (
        datos.solicitudId
          ?
            "hhee-"
            +
            datos.solicitudId
          :
            "hhee-webpush"
      );


    const opciones = {
      body:
        cuerpo,

      tag:
        tag,

      data: {
        url:
          url,

        solicitudId:
          datos.solicitudId
          ||
          "",

        tipo:
          datos.tipo
          ||
          "",

        canal:
          "webpush"
      },

      requireInteraction:
        true
    };


    event.waitUntil(
      self.registration
        .showNotification(
          titulo,
          opciones
        )
    );
  }
);


/* ============================================================
   FIREBASE MESSAGING

   IMPORTANTE:
   No cargamos Firebase Messaging en iPhone/iPad porque
   el canal de iOS es Web Push estándar.
   ============================================================ */

if (!ES_IOS) {

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


  messaging.onBackgroundMessage(
    payload => {
      console.log(
        "[HHEE] FCM segundo plano:",
        payload
      );
    }
  );
}
