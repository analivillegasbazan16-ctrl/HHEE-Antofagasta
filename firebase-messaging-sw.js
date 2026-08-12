/* ============================================================
   HHEE ANTOFAGASTA
   FIREBASE CLOUD MESSAGING - SERVICE WORKER
   ============================================================ */

const APP_BASE = "/HHEE-Antofagasta/";


/*
  IMPORTANTE:
  notificationclick queda definido ANTES de importar FCM.
*/

self.addEventListener(
  "notificationclick",
  event => {

    event.notification.close();


    const destino =
      event.notification?.data?.url
      ||
      APP_BASE;


    const urlDestino =
      new URL(
        destino,
        self.location.origin
      ).href;


    event.waitUntil(

      clients
        .matchAll({
          type: "window",
          includeUncontrolled: true
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
                    "[HHEE] No se pudo navegar la ventana:",
                    error
                  );
                }


                if (
                  "focus" in ventana
                ) {

                  return ventana.focus();
                }
              }
            }


            if (
              clients.openWindow
            ) {

              return clients.openWindow(
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
   FIREBASE
   ============================================================ */

importScripts(
  "https://www.gstatic.com/firebasejs/12.16.0/firebase-app-compat.js"
);

importScripts(
  "https://www.gstatic.com/firebasejs/12.16.0/firebase-messaging-compat.js"
);


const firebaseConfig = {

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
};


firebase.initializeApp(
  firebaseConfig
);


const messaging =
  firebase.messaging();


/* ============================================================
   MENSAJES EN SEGUNDO PLANO
   ============================================================ */

messaging.onBackgroundMessage(
  payload => {

    console.log(
      "[HHEE] Mensaje recibido en segundo plano:",
      payload
    );


    const data =
      payload?.data
      ||
      {};


    const titulo =
      data.title
      ||
      payload?.notification?.title
      ||
      "Gestión de Horas Extras";


    const cuerpo =
      data.body
      ||
      payload?.notification?.body
      ||
      "Tienes una nueva notificación.";


    const destino =
      data.url
      ||
      APP_BASE;


    const opciones = {

      body:
        cuerpo,

      tag:
        data.tag
        ||
        "hhee-notificacion",

      renotify:
        true,

      data: {

        url:
          destino,

        solicitudId:
          data.solicitudId
          ||
          "",

        tipo:
          data.tipo
          ||
          ""
      }
    };


    return self.registration
      .showNotification(
        titulo,
        opciones
      );
  }
);


self.addEventListener(
  "activate",
  event => {

    event.waitUntil(
      clients.claim()
    );
  }
);
