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
   INSTALACIÓN
   ============================================================ */

self.addEventListener(
  "install",
  event => {

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


  /*
    Notificaciones Web Push propias.
  */

  if (
    data.url
  ) {

    return String(
      data.url
    );
  }


  /*
    Firebase puede guardar el mensaje original
    dentro de FCM_MSG.
  */

  const fcm =
    data.FCM_MSG
    ||
    data.fcmMessage
    ||
    {};


  /*
    URL enviada dentro de data.url.
  */

  if (
    fcm?.data?.url
  ) {

    return String(
      fcm.data.url
    );
  }


  /*
    fcm_options.link.
  */

  if (
    fcm?.fcmOptions?.link
  ) {

    return String(
      fcm.fcmOptions.link
    );
  }


  /*
    Otra variante posible.
  */

  if (
    fcm?.fcm_options?.link
  ) {

    return String(
      fcm.fcm_options.link
    );
  }


  return APP_URL;
}


/* ============================================================
   CLICK EN NOTIFICACIÓN

   IMPORTANTE:
   Este listener está ANTES de importar Firebase.
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
        .matchAll(
          {
            type:
              "window",

            includeUncontrolled:
              true
          }
        )
        .then(
          async ventanas => {

            /*
              Primero buscamos una ventana de HHEE
              que ya esté abierta.
            */

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

                  /*
                    La llevamos a la URL exacta:

                    ?modulo=...
                    &solicitud=...
                  */

                  await ventana.navigate(
                    urlDestino
                  );

                }
                catch (
                  error
                ) {

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


            /*
              Si HHEE estaba cerrada,
              intentamos abrirla.

              En una PWA instalada Android/iOS,
              el sistema puede abrir la aplicación
              instalada correspondiente al scope.
            */

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

   Los mensajes enviados desde Cloudflare contienen:

   {
     title,
     body,
     url,
     tag,
     tipo,
     solicitudId
   }
   ============================================================ */

self.addEventListener(
  "push",
  event => {

    /*
      Intentamos leer el contenido del Push.
    */

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
      IMPORTANTE:

      Si no tiene nuestra estructura,
      NO hacemos nada.

      Esto permite que Firebase Messaging
      siga procesando sus propios mensajes.
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


    if (
      !esWebPushHhee
    ) {

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
            "hhee-" +
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
      }
    };


    /*
      requireInteraction funciona donde el
      navegador/sistema operativo lo admita.
    */

    opciones.requireInteraction =
      true;


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
   FIREBASE

   Se carga DESPUÉS de notificationclick.
   ============================================================ */

importScripts(
  "https://www.gstatic.com/firebasejs/10.13.2/firebase-app-compat.js"
);


importScripts(
  "https://www.gstatic.com/firebasejs/10.13.2/firebase-messaging-compat.js"
);


/* ============================================================
   CONFIGURACIÓN FIREBASE
   ============================================================ */

firebase.initializeApp(
  {

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
  }
);


/* ============================================================
   FIREBASE MESSAGING
   ============================================================ */

const messaging =
  firebase.messaging();


/* ============================================================
   MENSAJES FCM EN SEGUNDO PLANO

   Firebase muestra automáticamente las notificaciones
   que contienen notification + fcm_options.link.

   Por eso aquí NO mostramos una segunda notificación.
   ============================================================ */

messaging.onBackgroundMessage(
  payload => {

    console.log(
      "[HHEE] FCM segundo plano:",
      payload
    );
  }
);
