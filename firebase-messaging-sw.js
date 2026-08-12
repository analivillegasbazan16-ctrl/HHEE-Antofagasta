importScripts(
  "https://www.gstatic.com/firebasejs/12.16.0/firebase-app-compat.js"
);

importScripts(
  "https://www.gstatic.com/firebasejs/12.16.0/firebase-messaging-compat.js"
);


const firebaseConfig = {
  apiKey: "AIzaSyCtZny0My5plZKF3VIOFY0MgTdY8NN82ek",
  authDomain: "hhee-antofagasta.firebaseapp.com",
  projectId: "hhee-antofagasta",
  storageBucket: "hhee-antofagasta.firebasestorage.app",
  messagingSenderId: "1004191970210",
  appId: "1:1004191970210:web:c8be9ae46af57d2b7138fb",
  measurementId: "G-9PD7MCBCRS"
};


firebase.initializeApp(firebaseConfig);

const messaging = firebase.messaging();


messaging.onBackgroundMessage((payload) => {

  console.log(
    "Notificación recibida en segundo plano:",
    payload
  );


  const titulo =
    payload?.notification?.title
    ||
    payload?.data?.title
    ||
    "Gestión de Horas Extras";


  const cuerpo =
    payload?.notification?.body
    ||
    payload?.data?.body
    ||
    "Tienes una nueva notificación.";


  const url =
    payload?.data?.url
    ||
    "/HHEE-Antofagasta/";


  const opciones = {

    body: cuerpo,

    icon:
      "/HHEE-Antofagasta/icon-192.png",

    badge:
      "/HHEE-Antofagasta/icon-192.png",

    data: {
      url: url
    },

    tag:
      payload?.data?.tag
      ||
      "hhee-notificacion",

    renotify: true
  };


  return self.registration.showNotification(
    titulo,
    opciones
  );
});


self.addEventListener(
  "notificationclick",
  (event) => {

    event.notification.close();


    const destino =
      event.notification?.data?.url
      ||
      "/HHEE-Antofagasta/";


    event.waitUntil(

      clients.matchAll({
        type: "window",
        includeUncontrolled: true
      })
      .then((ventanas) => {

        for (
          const ventana of ventanas
        ) {

          if (
            ventana.url.includes(
              "/HHEE-Antofagasta/"
            )
            &&
            "focus" in ventana
          ) {

            ventana.navigate(destino);

            return ventana.focus();
          }
        }


        if (
          clients.openWindow
        ) {

          return clients.openWindow(
            destino
          );
        }


        return null;
      })
    );
  }
);
