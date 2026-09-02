<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">

  <meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
  >

  <meta name="theme-color" content="#003b71">

  <!-- Aplicación instalable -->
  <link rel="manifest" href="./manifest.webmanifest">

  <!-- Configuración para iPhone / iPad -->
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="default">
  <meta name="apple-mobile-web-app-title" content="HHEE">

  <title>Gestión de Horas Extras</title>

  <!-- Excel: lectura del SAP y exportación mensual -->
  <script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>

  <!-- PDF del informe de conciliación -->
  <script src="https://unpkg.com/jspdf@4.2.1/dist/jspdf.umd.min.js"></script>
  <script src="https://unpkg.com/jspdf-autotable@5.0.8/dist/jspdf.plugin.autotable.min.js"></script>

  <style>
    :root {
      --azul:#003b71;
      --azul2:#005b96;
      --azulClaro:#eaf4fb;
      --fondo:#f3f7fa;
      --fondoLogin:#062f5b;
      --fondoLogin2:#0a467f;
      --texto:#172033;
      --gris:#68778b;
      --borde:#dce5ec;
      --verde:#18865b;
      --verdeF:#eaf8f1;
      --rojo:#b52f38;
      --rojoF:#fff0f1;
      --naranjo:#8a5b00;
      --naranjoF:#fff4df;
      --morado:#7545a5;
      --moradoF:#f4edfb;
      --sombra:0 12px 32px rgba(0,59,113,.09);
    }

    *{box-sizing:border-box}
    html{scroll-behavior:smooth}
    body{margin:0;min-height:100vh;font-family:"Segoe UI",Arial,sans-serif;background:var(--fondo);color:var(--texto)}
    button,input,select,textarea{font:inherit}
    button{cursor:pointer}
    .hidden{display:none!important}

    #pantallaCarga{position:fixed;inset:0;z-index:9999;display:grid;place-items:center;background:var(--fondoLogin)}
    .spinner{width:44px;height:44px;border:4px solid #d9e6ee;border-top-color:var(--azul);border-radius:50%;animation:girar .8s linear infinite;margin:0 auto 12px}
    @keyframes girar{to{transform:rotate(360deg)}}

    #pantallaLogin{min-height:100vh;display:grid;grid-template-columns:1fr .82fr;background:linear-gradient(135deg,var(--fondoLogin),var(--fondoLogin2))}
    .loginMarca{display:flex;align-items:center;justify-content:center;padding:55px;color:#fff;background:linear-gradient(135deg,#062f5b 0%,#0a467f 55%,#0d5b9a 100%)}
    .loginMarca>div{max-width:580px}
    .logo{display:inline-block;padding:11px 17px;border:1px solid rgba(255,255,255,.42);border-radius:13px;background:rgba(255,255,255,.16);color:#ffffff;font-size:27px;font-weight:800;margin-bottom:26px;box-shadow:0 8px 20px rgba(0,0,0,.18)}
    .logo span{color:#ffd166}
    .loginMarca h1{font-size:clamp(38px,5vw,62px);line-height:1.05;margin:0 0 18px;color:#ffffff;text-shadow:0 3px 10px rgba(0,0,0,.25)}
    .loginMarca p{font-size:17px;line-height:1.6;color:rgba(255,255,255,.96);text-shadow:0 2px 8px rgba(0,0,0,.18)}
    .loginDerecha{display:flex;align-items:center;justify-content:center;padding:35px;background:var(--fondo);border-radius:38px 0 0 38px}
    .loginCard{width:min(430px,100%);padding:34px;background:#fff;border:1px solid var(--borde);border-radius:24px;box-shadow:var(--sombra)}
    .loginCard h2{margin:0 0 7px;color:var(--azul);font-size:28px}
    .sub{color:var(--gris);line-height:1.5;margin:0 0 24px}

    .campo{margin-bottom:16px}
    label{display:block;margin-bottom:7px;font-size:13px;font-weight:700;color:#334155}
    input,select,textarea{width:100%;padding:12px 13px;border:1px solid var(--borde);border-radius:10px;background:#fff;color:var(--texto);outline:none}
    input:focus,select:focus,textarea:focus{border-color:var(--azul2);box-shadow:0 0 0 3px rgba(0,91,150,.09)}
    input[readonly]{background:#f1f5f9;color:#475569}
    textarea{min-height:100px;resize:vertical}
    .passWrap{position:relative}
    .passWrap input{padding-right:48px}
    .passWrap button{position:absolute;right:5px;top:50%;transform:translateY(-50%);border:0;background:transparent;width:38px;height:36px}

    .btn{border:0;border-radius:10px;padding:12px 16px;font-weight:750}
    .btnPrimario{background:var(--azul);color:#fff}
    .btnPrimario:hover{background:var(--azul2)}
    .btnSecundario{border:1px solid var(--borde);background:#fff;color:var(--azul)}
    .btnPeligro{background:var(--rojo);color:#fff}
    .btnExito{background:var(--verde);color:#fff}
    .btn:disabled{opacity:.55;cursor:not-allowed}
    .w100{width:100%}
    .linkBtn{border:0;background:transparent;color:var(--azul);font-weight:700;padding:10px;width:100%}

    .mensaje{padding:12px 14px;border-radius:10px;margin:0 0 15px;line-height:1.45;font-size:13px}
    .mensajeError{background:var(--rojoF);color:var(--rojo)}
    .mensajeExito{background:var(--verdeF);color:var(--verde)}
    .mensajeInfo{background:var(--azulClaro);color:var(--azul)}
    .mensajeAviso{background:var(--naranjoF);color:var(--naranjo)}

    .cabecera{position:sticky;top:0;z-index:100;display:flex;align-items:center;justify-content:space-between;gap:15px;min-height:72px;padding:12px 25px;background:var(--azul);color:#fff;box-shadow:0 4px 14px rgba(0,0,0,.11)}
    .marcaHead{display:flex;align-items:center;gap:12px}
    .iconoHead{width:42px;height:42px;border-radius:12px;background:rgba(255,255,255,.13);display:grid;place-items:center;font-size:21px}
    .marcaHead h1{font-size:19px;margin:0}
    .marcaHead small{color:rgba(255,255,255,.72)}
    .usuarioHead{display:flex;align-items:center;gap:10px}
    .usuarioTexto{text-align:right}
    .usuarioTexto strong{display:block;font-size:14px}
    .usuarioTexto span{font-size:12px;color:rgba(255,255,255,.72)}
    .avatar{width:42px;height:42px;border-radius:50%;background:#fff;color:var(--azul);display:grid;place-items:center;font-weight:800}
    .btnSalir{border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.1);color:#fff;border-radius:9px;padding:9px 12px}

    /* ====================== NOTIFICACIONES ====================== */

    .notifBar{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:14px;
      margin-top:16px;
      padding:13px 14px;
      border:1px solid #cfe0eb;
      border-radius:12px;
      background:#fff;
    }

    .notifInfo{
      display:flex;
      align-items:center;
      gap:11px;
      min-width:0;
    }

    .notifIcon{
      width:40px;
      height:40px;
      border-radius:11px;
      background:var(--azulClaro);
      display:grid;
      place-items:center;
      font-size:20px;
      flex:0 0 auto;
    }

    .notifTexto strong{
      display:block;
      color:var(--azul);
      font-size:13px;
    }

    .notifTexto span{
      display:block;
      color:var(--gris);
      font-size:11px;
      margin-top:2px;
      line-height:1.35;
    }

    .btnNotif{
      border:1px solid #bdd5e5;
      background:var(--azul);
      color:#fff;
      border-radius:9px;
      padding:9px 12px;
      font-size:12px;
      font-weight:800;
      white-space:nowrap;
    }

    .btnNotif.activa{
      background:var(--verde);
      border-color:var(--verde);
    }

    .btnNotif.bloqueada{
      background:#64748b;
      border-color:#64748b;
    }

    .contenedor{width:min(1180px,calc(100% - 36px));margin:26px auto 60px}
    .card{background:#fff;border:1px solid var(--borde);border-radius:17px;box-shadow:0 7px 24px rgba(0,59,113,.05)}
    .perfil{padding:25px;background:linear-gradient(135deg,#fff,#edf6fb)}
    .perfil h2{margin:0 0 5px;color:var(--azul);font-size:27px}
    .perfil>p{margin:0;color:var(--gris)}
    .badges{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}
    .badge{padding:5px 9px;border-radius:999px;font-size:11px;font-weight:800}
    .bTrab{background:var(--verdeF);color:var(--verde)}
    .bSup{background:var(--azulClaro);color:var(--azul2)}
    .bAdm{background:var(--moradoF);color:var(--morado)}
    .perfilDatos{display:grid;grid-template-columns:repeat(4,1fr);gap:11px;margin-top:18px}
    .perfilDato{padding:12px 13px;border:1px solid var(--borde);border-radius:10px;background:rgba(255,255,255,.82)}
    .perfilDato span{display:block;color:var(--gris);font-size:10px;font-weight:800;text-transform:uppercase;margin-bottom:4px}

    .seccion{margin-top:30px}
    .seccion h2{font-size:21px;margin:0}
    .seccion>p{margin:5px 0 15px;color:var(--gris);font-size:13px}
    .gridModulos{display:grid;grid-template-columns:repeat(2,1fr);gap:15px}
    .modulo{display:flex;align-items:flex-start;gap:15px;min-height:120px;padding:19px;background:#fff;border:1px solid var(--borde);border-radius:15px;text-align:left;color:var(--texto);box-shadow:0 5px 17px rgba(0,59,113,.04)}
    .modulo:hover{transform:translateY(-2px);box-shadow:var(--sombra)}
    .modIcon{width:49px;height:49px;flex:0 0 auto;border-radius:13px;background:var(--azulClaro);display:grid;place-items:center;font-size:24px}
    .modulo h3{margin:2px 0 5px;color:var(--azul);font-size:16px}
    .modulo p{margin:0;color:var(--gris);font-size:13px;line-height:1.45}
    .flecha{margin-left:auto;color:#93a6b5;font-size:23px}

    /* ====================== NAVEGACIÓN ADMINISTRACIÓN ====================== */

    .adminTabs{
      position:sticky;
      top:84px;
      z-index:90;
      display:flex;
      align-items:center;
      gap:8px;
      margin-top:20px;
      padding:10px;
      overflow-x:auto;
      border:1px solid var(--borde);
      border-radius:14px;
      background:rgba(255,255,255,.97);
      box-shadow:0 7px 22px rgba(0,59,113,.06);
      scrollbar-width:thin;
    }

    .adminTab{
      flex:0 0 auto;
      border:1px solid transparent;
      border-radius:10px;
      padding:10px 13px;
      background:transparent;
      color:#52657a;
      font-size:12px;
      font-weight:800;
      white-space:nowrap;
    }

    .adminTab:hover{
      background:var(--azulClaro);
      color:var(--azul);
    }

    .adminTab.activa{
      background:var(--azul);
      color:#fff;
      box-shadow:0 4px 12px rgba(0,59,113,.18);
    }

    .adminModuloActivo #seccionTrabajador,
    .adminModuloActivo #seccionSupervisor,
    .adminModuloActivo #seccionAdministrador{
      display:none!important;
    }

    .adminModuloActivo #introAdmin{
      display:none!important;
    }

    .adminIntro{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:15px;
      margin-top:18px;
      padding:15px 17px;
      border:1px solid #cfe0eb;
      border-radius:12px;
      background:linear-gradient(135deg,#fff,#f3f9fd);
    }

    .adminIntro strong{display:block;color:var(--azul);font-size:14px}
    .adminIntro span{display:block;color:var(--gris);font-size:12px;margin-top:3px}

    .filtrosRango{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
    .periodoGrupo{display:grid;grid-template-columns:1fr 1fr;gap:8px}

    .tablaResumenTrabajador{min-width:1080px}
    .nombrePrincipal{font-weight:800;color:var(--azul)}
    .subDato{display:block;margin-top:2px;color:var(--gris);font-size:10px}

    .estadoRevision{background:var(--naranjoF);color:var(--naranjo)}
    .estadoSinSap{background:#fdecef;color:#a82332}

    .concCabecera{
      display:flex;
      align-items:flex-start;
      justify-content:space-between;
      gap:15px;
      margin-bottom:14px;
    }

    .concCabecera h3{margin:0 0 5px;color:var(--azul)}
    .concCabecera p{margin:0;color:var(--gris);font-size:12px;line-height:1.45}

    details.configAvanzada{
      margin-top:14px;
      border:1px solid var(--borde);
      border-radius:11px;
      background:#fff;
    }

    details.configAvanzada summary{
      cursor:pointer;
      padding:12px 14px;
      color:var(--azul);
      font-size:12px;
      font-weight:800;
    }

    details.configAvanzada .configContenido{padding:0 14px 14px}

    .notaReglaPago{
      margin:12px 0;
      padding:12px 14px;
      border-left:4px solid var(--naranjo);
      border-radius:8px;
      background:var(--naranjoF);
      color:#6d4b05;
      font-size:12px;
      line-height:1.5;
    }

    .bonificacionPago{
      display:inline-flex;
      align-items:center;
      gap:5px;
      margin-top:4px;
      padding:3px 7px;
      border-radius:999px;
      background:var(--azulClaro);
      color:var(--azul);
      font-size:10px;
      font-weight:800;
      white-space:nowrap;
    }

    .detallePagoLinea{
      display:block;
      margin-top:3px;
      color:var(--gris);
      font-size:10px;
      line-height:1.35;
    }

    #panelModulo{margin-top:18px;padding:24px}
    .panelHead{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;margin-bottom:20px}
    .panelHead h2{margin:0 0 5px;color:var(--azul)}
    .panelHead p{margin:0;color:var(--gris)}
    .bloque{padding:18px;border:1px solid var(--borde);border-radius:13px;background:#fbfdff;margin-bottom:16px}
    .bloque h3{margin:0 0 15px;color:var(--azul);font-size:17px}
    .grid2{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
    .grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
    .grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
    .full{grid-column:1/-1}
    .checkRow{display:flex;align-items:center;gap:9px}
    .checkRow input{width:18px;height:18px}
    .checkRow label{margin:0}
    .extension{padding:16px;margin-top:15px;border:1px solid #c5ddeb;border-radius:11px;background:var(--azulClaro)}
    .acciones{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end;margin-top:17px}
    .nota{font-size:11px;color:var(--gris);margin-top:5px;line-height:1.4}

    .resumenKpi{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0}
    .kpi{padding:14px;border:1px solid var(--borde);border-radius:11px;background:#fff}
    .kpi span{display:block;color:var(--gris);font-size:10px;font-weight:800;text-transform:uppercase;margin-bottom:6px}
    .kpi strong{font-size:21px;color:var(--azul)}

    .tablaWrap{overflow:auto;border:1px solid var(--borde);border-radius:11px;background:#fff}
    .tabla{width:100%;min-width:850px;border-collapse:collapse}
    .tabla th,.tabla td{padding:9px 10px;border-bottom:1px solid var(--borde);font-size:12px;text-align:left;vertical-align:top}
    .tabla th{background:#f6f9fb;color:#475569;position:sticky;top:0}
    .tabla tr:last-child td{border-bottom:0}

    .estado{display:inline-flex;padding:4px 8px;border-radius:999px;font-size:10px;font-weight:800}
    .pendiente{background:#fff3cd;color:#8a5b00}
    .aprobada{background:var(--verdeF);color:var(--verde)}
    .rechazada{background:var(--rojoF);color:var(--rojo)}
    .ok{background:var(--verdeF);color:var(--verde)}
    .diferencia{background:var(--rojoF);color:var(--rojo)}
    .miniAcciones{display:flex;gap:6px;flex-wrap:wrap}
    .miniAcciones .btn{padding:6px 8px;font-size:11px}
    .detalleTramos{margin-top:12px;padding-top:12px;border-top:1px solid var(--borde)}

    .modalFondo{position:fixed;inset:0;z-index:1000;background:rgba(15,23,42,.45);display:grid;place-items:center;padding:20px}
    .modal{width:min(760px,100%);max-height:90vh;overflow:auto;background:#fff;border-radius:17px;padding:22px;box-shadow:0 22px 70px rgba(0,0,0,.24)}
    .modal h3{margin:0 0 14px;color:var(--azul)}

    #toastContainer{position:fixed;right:18px;bottom:18px;z-index:2000;width:min(380px,calc(100% - 36px))}
    .toast{margin-top:8px;padding:13px 15px;border-radius:10px;color:#fff;box-shadow:0 10px 28px rgba(0,0,0,.16)}
    .toastInfo{background:var(--azul)}
    .toastExito{background:var(--verde)}
    .toastError{background:var(--rojo)}

    @media(max-width:820px){
      html,body{background:#f3f7fa!important}
      #pantallaLogin:not(.hidden){display:block;min-height:100vh;background:#f3f7fa}
      .loginMarca{
        min-height:515px!important;
        padding:54px 42px 82px!important;
        align-items:center!important;
        justify-content:flex-start!important;
        background:linear-gradient(160deg,#003b71 0%,#004f86 58%,#0067a3 100%)!important;
        color:#fff!important;
      }
      .loginMarca>div{max-width:100%!important;width:100%!important}
      .logo{
        color:#fff!important;
        -webkit-text-fill-color:#fff!important;
        background:rgba(255,255,255,.13)!important;
        border-color:rgba(255,255,255,.38)!important;
        box-shadow:0 8px 22px rgba(0,0,0,.18)!important;
      }
      .logo span{-webkit-text-fill-color:#ffd166!important;color:#ffd166!important}
      .loginMarca h1{
        font-size:42px!important;
        line-height:1.08!important;
        color:#fff!important;
        -webkit-text-fill-color:#fff!important;
        text-shadow:0 3px 12px rgba(0,0,0,.28)!important;
      }
      .loginMarca p{
        font-size:17px!important;
        line-height:1.55!important;
        color:#fff!important;
        -webkit-text-fill-color:#fff!important;
        opacity:1!important;
        text-shadow:0 2px 8px rgba(0,0,0,.2)!important;
      }
      .loginDerecha{
        margin-top:-34px!important;
        padding:0 15px 30px!important;
        border-radius:30px 30px 0 0!important;
        background:#f3f7fa!important;
      }
      .usuarioTexto,.marcaHead small{display:none}
      .notifBar{align-items:flex-start;flex-direction:column}
      .btnNotif{width:100%}
      .cabecera{padding:11px 14px}
      .contenedor{width:calc(100% - 22px);margin-top:17px}
      .perfilDatos,.gridModulos,.grid2,.grid3,.grid4,.filtrosRango{grid-template-columns:1fr}
      .resumenKpi{grid-template-columns:repeat(2,1fr)}
      .acciones{flex-direction:column-reverse}
      .acciones .btn{width:100%}
      .adminTabs{top:72px;margin-left:-2px;margin-right:-2px}
      .periodoGrupo{grid-template-columns:1fr 1fr}
      .adminIntro{align-items:flex-start;flex-direction:column}
    }
  </style>
</head>

<body>

  <div id="pantallaCarga">
    <div style="text-align:center;color:#003b71">
      <div class="spinner"></div>
      <strong>Cargando sistema...</strong>
    </div>
  </div>

  <main id="pantallaLogin" class="hidden">

    <section class="loginMarca">
      <div>
        <div class="logo">transelec<span>.</span></div>
        <h1>Gestión de<br>Horas Extras</h1>
        <p>Registro, revisión, aprobación y conciliación de horas extras desde computador o celular.</p>
      </div>
    </section>

    <section class="loginDerecha">

      <div class="loginCard">

        <h2>Iniciar sesión</h2>

        <p class="sub">
          Ingresa con la cuenta habilitada para el sistema de horas extras.
        </p>

        <div id="mensajeLogin" class="hidden"></div>

        <form id="formLogin">

          <div class="campo">
            <label>Correo electrónico</label>
            <input id="correo" type="email" autocomplete="username" required>
          </div>

          <div class="campo">
            <label>Contraseña</label>

            <div class="passWrap">
              <input id="password" type="password" autocomplete="current-password" required>
              <button id="verPassword" type="button">👁</button>
            </div>
          </div>

          <button id="btnLogin" class="btn btnPrimario w100" type="submit">
            Ingresar
          </button>

          <button id="btnRecuperar" class="linkBtn" type="button">
            Olvidé mi contraseña
          </button>

          <div style="margin:6px 0 4px;border-top:1px solid #e5edf3"></div>

          <button id="btnActivarCuenta" class="linkBtn" type="button">
            ¿Primera vez? Activar mi cuenta
          </button>

        </form>

        <div style="margin-top:20px;color:#7a8798;font-size:11px;text-align:center">
          Sistema de uso restringido. Solo usuarios previamente autorizados pueden acceder.
        </div>

      </div>

    </section>

  </main>


  <div id="pantallaApp" class="hidden">

    <header class="cabecera">

      <div class="marcaHead">
        <div class="iconoHead">⏱</div>

        <div>
          <h1>Gestión de Horas Extras</h1>
          <small>Portal HHEE</small>
        </div>
      </div>

      <div class="usuarioHead">

        <div class="usuarioTexto">
          <strong id="nombreHeader">Usuario</strong>
          <span id="areaHeader"></span>
        </div>

        <div id="avatarUsuario" class="avatar">U</div>

        <button id="btnSalir" class="btnSalir" type="button">
          Cerrar sesión
        </button>

      </div>

    </header>


    <main class="contenedor">

      <section class="card perfil">

        <h2>
          Hola, <span id="nombreBienvenida">Usuario</span>
        </h2>

        <p>
          Bienvenido al sistema de gestión de horas extras.
        </p>

        <div id="badgesUsuario" class="badges"></div>


        <!-- =====================================================
             NOTIFICACIONES
        ====================================================== -->

        <div id="barraNotificaciones" class="notifBar">

          <div class="notifInfo">

            <div class="notifIcon">
              🔔
            </div>

            <div class="notifTexto">
              <strong>Notificaciones del celular</strong>
              <span id="estadoNotificaciones">
                Comprobando compatibilidad...
              </span>
            </div>

          </div>

          <button
            id="btnNotificaciones"
            class="btnNotif"
            type="button">
            Activar notificaciones
          </button>

        </div>


        <div class="perfilDatos">

          <div class="perfilDato">
            <span>RUT</span>
            <strong id="perfilRut">-</strong>
          </div>

          <div class="perfilDato">
            <span>N° de personal</span>
            <strong id="perfilNumeroPersonal">-</strong>
          </div>

          <div class="perfilDato">
            <span>Área</span>
            <strong id="perfilArea">-</strong>
          </div>

          <div class="perfilDato">
            <span>Correo</span>
            <strong id="perfilCorreo">-</strong>
          </div>

        </div>

      </section>

      <nav id="barraAdminTabs" class="adminTabs hidden" aria-label="Módulos de Administración">

        <button class="adminTab" type="button" data-modulo="resumenAdmin">Resumen</button>
        <button class="adminTab" type="button" data-modulo="todasSolicitudes">Solicitudes</button>
        <button class="adminTab" type="button" data-modulo="horasPagar">Horas a pagar</button>
        <button class="adminTab" type="button" data-modulo="conciliacion">Conciliación SAP</button>
        <button class="adminTab" type="button" data-modulo="exportar">Exportar</button>
        <button class="adminTab" type="button" data-modulo="usuarios">Usuarios</button>

      </nav>

      <div id="introAdmin" class="adminIntro hidden">
        <div>
          <strong>Panel de Administración</strong>
          <span>Selecciona una pestaña. Solo se mostrará el módulo elegido, sin tener que bajar por toda la página.</span>
        </div>
      </div>


      <section id="seccionTrabajador" class="seccion hidden">

        <h2>Mis horas extras</h2>

        <p>
          Registro y seguimiento de tus solicitudes.
        </p>

        <div class="gridModulos">

          <button class="modulo" data-modulo="ingresoNormal">

            <div class="modIcon">➕</div>

            <div>
              <h3>Ingresar horas extras</h3>
              <p>Formulario normal de lunes a sábado.</p>
            </div>

            <div class="flecha">›</div>

          </button>


          <button class="modulo" data-modulo="domingoFeriado">

            <div class="modIcon">📅</div>

            <div>
              <h3>Domingos y feriados</h3>
              <p>Formulario especial para domingo o feriado.</p>
            </div>

            <div class="flecha">›</div>

          </button>


          <button class="modulo" data-modulo="misSolicitudes">

            <div class="modIcon">📋</div>

            <div>
              <h3>Mis solicitudes</h3>
              <p>Ver, modificar o eliminar solicitudes pendientes.</p>
            </div>

            <div class="flecha">›</div>

          </button>


          <button class="modulo" data-modulo="misHoras">

            <div class="modIcon">⏲</div>

            <div>
              <h3>Mis horas</h3>
              <p>Resumen de horas aprobadas al 50%, 75% y 100%.</p>
            </div>

            <div class="flecha">›</div>

          </button>

        </div>

      </section>


      <section id="seccionSupervisor" class="seccion hidden">

        <h2>Supervisión</h2>

        <p>
          Revisión de las solicitudes de tus trabajadores asignados.
        </p>

        <div class="gridModulos">

          <button class="modulo" data-modulo="pendientesSupervisor">

            <div class="modIcon">🔔</div>

            <div>
              <h3>Solicitudes pendientes</h3>
              <p>Aprobar o rechazar solicitudes.</p>
            </div>

            <div class="flecha">›</div>

          </button>


          <button class="modulo" data-modulo="historialSupervisor">

            <div class="modIcon">🗂</div>

            <div>
              <h3>Historial de revisión</h3>
              <p>Solicitudes que ya revisaste.</p>
            </div>

            <div class="flecha">›</div>

          </button>

        </div>

      </section>


      <section id="seccionAdministrador" class="seccion hidden">

        <h2>Administración</h2>
        <p>Usa las pestañas superiores para cambiar de módulo.</p>

      </section>


      <section id="panelModulo" class="card hidden">

        <div class="panelHead">

          <div>
            <h2 id="tituloModulo">Módulo</h2>
            <p id="descripcionModulo"></p>
          </div>

          <button id="cerrarPanel" class="btn btnSecundario" type="button">
            Cerrar
          </button>

        </div>

        <div id="contenidoModulo"></div>

      </section>

    </main>

  </div>


  <div id="modalContainer"></div>
  <div id="toastContainer"></div>


<script type="module">

  /* =========================================================
     FIREBASE
  ========================================================= */

  import {
    initializeApp
  } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-app.js";


  import {
    getAuth,
    signInWithEmailAndPassword,
    signOut,
    onAuthStateChanged,
    sendPasswordResetEmail,
    createUserWithEmailAndPassword,
    deleteUser
  } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-auth.js";


  import {
    getFirestore,
    collection,
    doc,
    getDoc,
    getDocs,
    addDoc,
    setDoc,
    updateDoc,
    deleteDoc,
    query,
    where,
    serverTimestamp,
    writeBatch
  } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-firestore.js";


  /* =========================================================
     FIREBASE CLOUD MESSAGING
  ========================================================= */

  import {
    getMessaging,
    isSupported as messagingIsSupported,
    onMessage,
    onRegistered,
    onUnregistered,
    register,
    unregister
  } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-messaging.js";


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


  const VAPID_KEY =
    "BE5FX2QbbkuLxvbaOfMI06EzOzdTcj8So10dfof-J99sxywBUYBxCimpXqlsNAdFCnB5osVnETWmKVSWlMMewvM";


  const app =
    initializeApp(firebaseConfig);

  const auth =
    getAuth(app);

  const db =
    getFirestore(app);


  /* =========================================================
     VARIABLES DE NOTIFICACIONES
  ========================================================= */

  let messaging =
    null;

  let swRegistration =
    null;

  let fidActual =
    "";

  let notificacionesActivas =
    false;

  let listenersMessagingConfigurados =
    false;

  let webPushSubscriptionActual =
    null;

  let webPushPublicKeyCache =
    "";

  let canalNotificaciones =
    "";


  /* =========================================================
     VARIABLES GENERALES
  ========================================================= */

  let usuarioActual = null;
  let perfilActual = null;

  let cacheSolicitudes =
    new Map();

  let cacheUsuarios =
    new Map();

  let cacheAutorizaciones =
    new Map();

  let activacionEnProceso =
    false;

  let sapFilas = [];
  let sapCabeceras = [];
  let sapArchivoNombre = "";

  let ultimaConciliacion =
    null;


  const MINUTO =
    60000;

  const DIA =
    86400000;


  const MESES = [
    "",
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre"
  ];


  /* =========================================================
     UTILIDADES
  ========================================================= */

  const $ =
    id =>
      document.getElementById(id);


  const normEstado =
    v =>
      String(v || "")
        .trim()
        .toLowerCase();


  const correoActual =
    () =>
      String(
        usuarioActual?.email || ""
      ).trim();


  const esAdmin =
    () =>
      perfilActual?.esAdministrador === true;


  const esSupervisor =
    () =>
      perfilActual?.esSupervisor === true;


  const puedeSolicitar =
    () =>
      perfilActual?.puedeSolicitarHE === true;


  function esc(v) {

    return String(v ?? "")
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;")
      .replace(/'/g,"&#039;");
  }


  function normalizarTexto(v) {

    return String(v ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g,"")
      .trim()
      .toLowerCase()
      .replace(/\s+/g," ");
  }


  function normalizarRut(v) {

    return String(v ?? "")
      .toUpperCase()
      .replace(/[^0-9K]/g,"");
  }


  function normalizarNumeroPersonal(v) {

    return String(v ?? "")
      .trim()
      .replace(/\s+/g, "");
  }


  function claveNumeroPersonal(v) {

    const texto =
      normalizarNumeroPersonal(v);


    if (/^\d+$/.test(texto)) {

      return texto.replace(/^0+(?=\d)/, "");
    }


    return texto.toUpperCase();
  }


  function normalizarCorreo(v) {

    return String(v ?? "")
      .trim()
      .toLowerCase();
  }


  function numero(v) {

    if (
      typeof v === "number"
    ) {

      return Number.isFinite(v)
        ? v
        : 0;
    }


    const s =
      String(v ?? "")
        .trim()
        .replace(/\s/g,"");


    if (!s) {
      return 0;
    }


    const limpio =
      s.includes(",")
      &&
      s.includes(".")
        ?
          s
            .replace(/\./g,"")
            .replace(",",".")
        :
          s
            .replace(",",".")
            .replace(/%/g,"");


    const n =
      Number(limpio);


    return Number.isFinite(n)
      ?
        n
      :
        0;
  }


  function recargoNumero(v) {

    let n =
      numero(v);


    if (
      n > 0
      &&
      n <= 1
    ) {

      n *= 100;
    }


    return Math.round(n);
  }


  function recargoSapNumero(v) {

    const n =
      Math.round(numero(v));


    if (n === 1214) return 50;
    if (n === 1215) return 75;
    if (n === 1216) return 100;


    return recargoNumero(v);
  }


  function red2(n) {

    return Math.round(
      (
        Number(n)
        ||
        0
      )
      *
      100
    )
    /
    100;
  }


  function vistaHoras(n) {

    return red2(n)
      .toFixed(2)
      .replace(".",",");
  }


  function minutosHoras(m) {

    return red2(
      (
        Number(m)
        ||
        0
      )
      /
      60
    );
  }


  /*
    =========================================================
    HORAS PAGABLES - REGLA DE MÍNIMO
    =========================================================

    Regla aplicada en Administración y Conciliación SAP:
    - Se conservan siempre las horas reales aprobadas.
    - El cálculo se agrupa por trabajador + fecha + recargo.
    - Si el total aprobado de ese grupo es mayor que 0 y menor
      que 3,00 h, se completa hasta 3,00 h.
    - La diferencia se informa por separado como BONIFICACIÓN.
    - La bonificación mantiene el mismo recargo del grupo.

    Ejemplo:
      1,00 h real al 50% + 2,00 h bonificación al 50%
      = 3,00 h a pagar al 50%.
  */

  const REGLA_PAGO_CONFIGURADA = true;
  const HORAS_MINIMAS_PAGO = 3;


  function detallePagoDesdeHoras(horasAprobadas) {

    const registradas =
      red2(
        Math.max(
          0,
          numero(horasAprobadas)
        )
      );


    const pagar =
      registradas > 0
      &&
      registradas < HORAS_MINIMAS_PAGO
        ? HORAS_MINIMAS_PAGO
        : registradas;


    const bonificacion =
      red2(
        Math.max(
          0,
          pagar - registradas
        )
      );


    return {
      registradas,
      bonificacion,
      pagar: red2(pagar)
    };
  }


  function horasPagablesDesdeHoras(horasAprobadas) {

    return detallePagoDesdeHoras(
      horasAprobadas
    ).pagar;
  }


  function bonificacionDesdeHoras(horasAprobadas) {

    return detallePagoDesdeHoras(
      horasAprobadas
    ).bonificacion;
  }


  function textoEstadoReglaPago() {

    return `Mínimo ${vistaHoras(HORAS_MINIMAS_PAGO)} h por trabajador, fecha y recargo. Si las horas aprobadas son menores, la diferencia se agrega como bonificación al mismo recargo.`;
  }


  function toast(
    msg,
    tipo = "info"
  ) {

    const d =
      document.createElement(
        "div"
      );


    d.className =
      `toast toast${
        tipo.charAt(0).toUpperCase()
        +
        tipo.slice(1)
      }`;


    d.textContent =
      msg;


    $("toastContainer")
      .appendChild(d);


    setTimeout(
      () =>
        d.remove(),
      3500
    );
  }


  function mostrarLogin() {

    $("pantallaCarga")
      .classList
      .add("hidden");

    $("pantallaApp")
      .classList
      .add("hidden");

    $("pantallaLogin")
      .classList
      .remove("hidden");
  }


  function mostrarCarga() {

    $("pantallaCarga")
      .classList
      .remove("hidden");

    $("pantallaLogin")
      .classList
      .add("hidden");

    $("pantallaApp")
      .classList
      .add("hidden");
  }


  function mostrarApp() {

    $("pantallaCarga")
      .classList
      .add("hidden");

    $("pantallaLogin")
      .classList
      .add("hidden");

    $("pantallaApp")
      .classList
      .remove("hidden");
  }


  function loginMsg(
    msg,
    tipo = "error"
  ) {

    $("mensajeLogin").className =
      `mensaje ${
        tipo === "error"
          ?
            "mensajeError"
          :
            "mensajeExito"
      }`;


    $("mensajeLogin")
      .textContent =
        msg;


    $("mensajeLogin")
      .classList
      .remove("hidden");
  }


  function modal(html) {

    $("modalContainer")
      .innerHTML =
        `
          <div
            class="modalFondo"
            id="modalFondo"
          >
            <div class="modal">
              ${html}
            </div>
          </div>
        `;
  }


  function cerrarModal() {

    $("modalContainer")
      .innerHTML =
        "";
  }


  window.cerrarModal =
    cerrarModal;


  /* =========================================================
     NOTIFICACIONES
     - Android / PC: Firebase Cloud Messaging
     - iPhone / iPad PWA: Web Push estándar
  ========================================================= */

  const WEB_PUSH_WORKER =
    "https://hhee-webpush.analivillegasbazan16.workers.dev";


  function idDispositivo() {

    const clave =
      "hheeDeviceId";


    let id =
      localStorage.getItem(
        clave
      );


    if (!id) {

      id =
        (
          crypto?.randomUUID?.()
          ||
          `${Date.now()}_${Math.random().toString(36).slice(2)}`
        )
        .replace(
          /[^a-zA-Z0-9]/g,
          ""
        );


      localStorage.setItem(
        clave,
        id
      );
    }


    return id;
  }


  function idRegistroFcm(uid) {

    return `${uid}_${idDispositivo()}`;
  }


  function idRegistroWebPush(uid) {

    return `${uid}_${idDispositivo()}`;
  }


  function esIOS() {

    const ua =
      navigator.userAgent
      ||
      "";


    const iOSClasico =
      /iPhone|iPad|iPod/i.test(
        ua
      );


    const iPadComoMac =
      navigator.platform ===
      "MacIntel"
      &&
      navigator.maxTouchPoints >
      1;


    return (
      iOSClasico
      ||
      iPadComoMac
    );
  }


  function esPwaInstalada() {

    return (
      window.matchMedia(
        "(display-mode: standalone)"
      ).matches
      ||
      window.navigator.standalone ===
      true
    );
  }


  function versionIOSActual() {

    if (!esIOS()) {
      return null;
    }


    const ua =
      navigator.userAgent
      ||
      "";


    const match =
      ua.match(
        /OS\s(\d+)[._](\d+)(?:[._](\d+))?/i
      );


    if (!match) {
      return null;
    }


    return {
      major:
        Number(match[1] || 0),

      minor:
        Number(match[2] || 0),

      patch:
        Number(match[3] || 0)
    };
  }


  function iosCompatibleConWebPush() {

    const version =
      versionIOSActual();


    if (!version) {
      return true;
    }


    return (
      version.major > 16
      ||
      (
        version.major === 16
        &&
        version.minor >= 4
      )
    );
  }


  function diagnosticoWebPushIOS() {

    return {
      pwa:
        esPwaInstalada(),

      secureContext:
        window.isSecureContext === true,

      serviceWorker:
        "serviceWorker" in navigator,

      notification:
        "Notification" in window,

      pushManagerGlobal:
        "PushManager" in window,

      pushManagerVentana:
        Boolean(
          window.pushManager
        ),

      version:
        versionIOSActual()
    };
  }


  function soportaWebPushEstandar() {

    /*
      En iPhone no bloqueamos Web Push por la presencia del
      constructor global Notification. La suscripción se obtiene
      desde ServiceWorkerRegistration.pushManager. El permiso se
      solicita al llamar subscribe() desde el botón del usuario.
    */

    return (
      window.isSecureContext === true
      &&
      "serviceWorker" in navigator
      &&
      iosCompatibleConWebPush()
    );
  }


  async function obtenerPushManagerWeb() {

    const registro =
      await prepararServiceWorker();


    const pushManager =
      registro?.pushManager
      ||
      window.pushManager
      ||
      null;


    if (!pushManager) {
      throw new Error(
        "PUSH_MANAGER_NO_DISPONIBLE"
      );
    }


    return {
      registro,
      pushManager
    };
  }


  function urlModuloNotificacion(
    modulo,
    solicitudId = ""
  ) {

    const u =
      new URL(
        window.location.href
      );


    u.search =
      "";

    u.hash =
      "";


    u.searchParams.set(
      "modulo",
      modulo
    );


    if (solicitudId) {

      u.searchParams.set(
        "solicitud",
        solicitudId
      );
    }


    return u.href;
  }


  function base64UrlAUint8Array(
    base64String
  ) {

    const padding =
      "=".repeat(
        (
          4
          -
          base64String.length % 4
        )
        %
        4
      );


    const base64 =
      (
        base64String
        +
        padding
      )
      .replace(/-/g, "+")
      .replace(/_/g, "/");


    const rawData =
      window.atob(
        base64
      );


    const outputArray =
      new Uint8Array(
        rawData.length
      );


    for (
      let i = 0;
      i < rawData.length;
      i++
    ) {

      outputArray[i] =
        rawData.charCodeAt(i);
    }


    return outputArray;
  }


  async function prepararServiceWorker() {

    if (swRegistration) {
      return swRegistration;
    }


    if (
      !("serviceWorker" in navigator)
    ) {

      throw new Error(
        "SERVICE_WORKER_NO_DISPONIBLE"
      );
    }


    swRegistration =
      await navigator
        .serviceWorker
        .register(
          "./firebase-messaging-sw.js",
          {
            scope:
              "./"
          }
        );


    await navigator
      .serviceWorker
      .ready;


    return swRegistration;
  }


  async function obtenerClavePublicaWebPush() {

    if (webPushPublicKeyCache) {
      return webPushPublicKeyCache;
    }


    const respuesta =
      await fetch(
        `${WEB_PUSH_WORKER}/vapid-public-key`,
        {
          method:
            "GET",

          cache:
            "no-store"
        }
      );


    if (!respuesta.ok) {

      throw new Error(
        "No fue posible obtener la clave pública Web Push."
      );
    }


    const datos =
      await respuesta.json();


    if (
      !datos?.ok
      ||
      !datos?.publicKey
    ) {

      throw new Error(
        "El servicio Web Push no devolvió una clave válida."
      );
    }


    webPushPublicKeyCache =
      String(
        datos.publicKey
      );


    return webPushPublicKeyCache;
  }


  function plataformaActual() {

    if (esIOS()) {
      return "iOS";
    }


    return (
      navigator.userAgentData?.platform
      ||
      navigator.platform
      ||
      ""
    );
  }


  async function guardarRegistroWebPush(
    subscription
  ) {

    if (
      !usuarioActual
      ||
      !subscription
    ) {
      return;
    }


    const json =
      subscription.toJSON();


    if (
      !json?.endpoint
      ||
      !json?.keys?.p256dh
      ||
      !json?.keys?.auth
    ) {

      throw new Error(
        "La suscripción Web Push está incompleta."
      );
    }


    const registroRef =
      doc(
        db,
        "webPushSubscriptions",
        idRegistroWebPush(
          usuarioActual.uid
        )
      );


    await setDoc(
      registroRef,
      {
        uid:
          usuarioActual.uid,

        correo:
          normalizarCorreo(
            usuarioActual.email
            ||
            perfilActual?.Correo
            ||
            ""
          ),

        dispositivoId:
          idDispositivo(),

        endpoint:
          json.endpoint,

        p256dh:
          json.keys.p256dh,

        auth:
          json.keys.auth,

        activo:
          true,

        actualizadoEn:
          serverTimestamp(),

        userAgent:
          navigator.userAgent
          ||
          "",

        plataforma:
          plataformaActual(),

        origen:
          "webpush"
      },
      {
        merge:
          true
      }
    );


    webPushSubscriptionActual =
      subscription;

    canalNotificaciones =
      "webpush";

    notificacionesActivas =
      true;


    actualizarEstadoNotificaciones();
  }


  async function marcarRegistroWebPushInactivo() {

    if (!usuarioActual) {
      return;
    }


    try {

      await setDoc(
        doc(
          db,
          "webPushSubscriptions",
          idRegistroWebPush(
            usuarioActual.uid
          )
        ),
        {
          uid:
            usuarioActual.uid,

          correo:
            normalizarCorreo(
              usuarioActual.email
              ||
              perfilActual?.Correo
              ||
              ""
            ),

          dispositivoId:
            idDispositivo(),

          activo:
            false,

          actualizadoEn:
            serverTimestamp()
        },
        {
          merge:
            true
        }
      );

    }
    catch (err) {

      console.warn(
        "[HHEE] No se pudo marcar Web Push como inactivo:",
        err
      );
    }
  }


  async function activarWebPushIOS() {

    if (!usuarioActual) {

      return toast(
        "Primero inicia sesión.",
        "error"
      );
    }


    if (!esPwaInstalada()) {

      actualizarEstadoNotificaciones();


      return toast(
        "En iPhone debes abrir HHEE desde el icono agregado a la pantalla de inicio.",
        "error"
      );
    }


    if (!soportaWebPushEstandar()) {

      const diag =
        diagnosticoWebPushIOS();


      console.log(
        "[HHEE] Diagnóstico Web Push iOS:",
        diag
      );


      if (!iosCompatibleConWebPush()) {

        return toast(
          "Web Push requiere iOS 16.4 o posterior.",
          "error"
        );
      }


      if (!diag.secureContext) {

        return toast(
          "HHEE debe abrirse mediante HTTPS para activar notificaciones.",
          "error"
        );
      }


      return toast(
        "El Service Worker no está disponible en esta instalación de HHEE.",
        "error"
      );
    }


    try {

      const {
        pushManager
      } =
        await obtenerPushManagerWeb();


      let subscription =
        await pushManager
          .getSubscription();


      if (!subscription) {

        const publicKey =
          await obtenerClavePublicaWebPush();


        /*
          En iOS el permiso se solicita desde subscribe().
          No exigimos Notification.requestPermission() como paso previo.
        */

        subscription =
          await pushManager
            .subscribe(
              {
                userVisibleOnly:
                  true,

                applicationServerKey:
                  base64UrlAUint8Array(
                    publicKey
                  )
              }
            );
      }


      await guardarRegistroWebPush(
        subscription
      );


      toast(
        "Notificaciones activadas en este iPhone.",
        "exito"
      );

    }
    catch (err) {

      console.error(
        "[HHEE] Error activando Web Push iOS:",
        err
      );


      const nombre =
        String(
          err?.name
          ||
          ""
        );


      const mensaje =
        String(
          err?.message
          ||
          err
          ||
          ""
        );


      if (
        nombre === "NotAllowedError"
        ||
        nombre === "SecurityError"
      ) {

        toast(
          "iOS no otorgó permiso para las notificaciones. Revisa Configuración > Notificaciones > Gestión de Horas Extras.",
          "error"
        );

      }
      else if (
        mensaje.includes(
          "PUSH_MANAGER_NO_DISPONIBLE"
        )
      ) {

        toast(
          "PushManager no está disponible. Actualiza iOS y vuelve a agregar HHEE a la pantalla de inicio desde Safari.",
          "error"
        );

      }
      else if (
        err?.code?.includes?.(
          "permission"
        )
      ) {

        toast(
          "Firestore rechazó guardar la suscripción del iPhone.",
          "error"
        );

      }
      else {

        toast(
          "No fue posible activar las notificaciones en este iPhone.",
          "error"
        );
      }
    }
    finally {

      actualizarEstadoNotificaciones();
    }
  }


  async function desactivarWebPushIOS() {

    if (
      !confirm(
        "¿Desactivar las notificaciones en este dispositivo?"
      )
    ) {
      return;
    }


    try {

      const {
        pushManager
      } =
        await obtenerPushManagerWeb();


      const subscription =
        webPushSubscriptionActual
        ||
        await pushManager
          .getSubscription();


      await marcarRegistroWebPushInactivo();


      if (subscription) {

        await subscription
          .unsubscribe();
      }


      webPushSubscriptionActual =
        null;

      canalNotificaciones =
        "";

      notificacionesActivas =
        false;


      actualizarEstadoNotificaciones();


      toast(
        "Notificaciones desactivadas en este iPhone.",
        "exito"
      );

    }
    catch (err) {

      console.error(
        "[HHEE] Error desactivando Web Push:",
        err
      );


      toast(
        "No fue posible desactivar las notificaciones.",
        "error"
      );
    }
  }


  function actualizarEstadoNotificaciones() {

    const estado =
      $("estadoNotificaciones");

    const btn =
      $("btnNotificaciones");


    if (
      !estado
      ||
      !btn
    ) {
      return;
    }


    if (esIOS()) {

      if (!esPwaInstalada()) {

        estado.textContent =
          "En iPhone, agrega HHEE a la pantalla de inicio y ábrela desde ese icono para activar notificaciones.";

        btn.textContent =
          "Abrir como app";

        btn.disabled =
          true;

        btn.className =
          "btnNotif bloqueada";

        return;
      }


      if (!iosCompatibleConWebPush()) {

        const v =
          versionIOSActual();


        estado.textContent =
          v
            ?
              `Este iPhone usa iOS ${v.major}.${v.minor}. Web Push requiere iOS 16.4 o posterior.`
            :
              "Web Push requiere iOS 16.4 o posterior.";

        btn.textContent =
          "No disponible";

        btn.disabled =
          true;

        btn.className =
          "btnNotif bloqueada";

        return;
      }


      if (
        !window.isSecureContext
        ||
        !("serviceWorker" in navigator)
      ) {

        estado.textContent =
          "El Service Worker de HHEE no está disponible en este iPhone.";

        btn.textContent =
          "No disponible";

        btn.disabled =
          true;

        btn.className =
          "btnNotif bloqueada";

        return;
      }


      if (
        notificacionesActivas
        &&
        webPushSubscriptionActual
      ) {

        estado.textContent =
          "Este iPhone está registrado para recibir avisos.";

        btn.textContent =
          "Desactivar";

        btn.disabled =
          false;

        btn.className =
          "btnNotif activa";

        return;
      }


      estado.textContent =
        "Activa los avisos de HHEE en este iPhone.";

      btn.textContent =
        "Activar notificaciones";

      btn.disabled =
        false;

      btn.className =
        "btnNotif";

      return;
    }


    if (
      !("Notification" in window)
      ||
      !("serviceWorker" in navigator)
    ) {

      estado.textContent =
        "Este navegador no admite notificaciones web.";

      btn.textContent =
        "No disponible";

      btn.disabled =
        true;

      btn.className =
        "btnNotif bloqueada";

      return;
    }


    if (
      Notification.permission ===
      "denied"
    ) {

      estado.textContent =
        "Las notificaciones están bloqueadas en el navegador.";

      btn.textContent =
        "Bloqueadas";

      btn.disabled =
        false;

      btn.className =
        "btnNotif bloqueada";

      return;
    }


    if (
      notificacionesActivas
      &&
      fidActual
    ) {

      estado.textContent =
        "Este dispositivo está registrado para recibir avisos.";

      btn.textContent =
        "Desactivar";

      btn.disabled =
        false;

      btn.className =
        "btnNotif activa";

      return;
    }


    if (
      Notification.permission ===
      "granted"
    ) {

      estado.textContent =
        "Permiso concedido. Falta sincronizar este dispositivo.";

      btn.textContent =
        "Activar notificaciones";

      btn.disabled =
        false;

      btn.className =
        "btnNotif";

      return;
    }


    estado.textContent =
      "Actívalas para recibir avisos de solicitudes y decisiones.";

    btn.textContent =
      "Activar notificaciones";

    btn.disabled =
      false;

    btn.className =
      "btnNotif";
  }


  async function guardarRegistroFcm(
    installationId
  ) {

    if (
      !usuarioActual
      ||
      !installationId
    ) {
      return;
    }


    const registroRef =
      doc(
        db,
        "fcmRegistrations",
        idRegistroFcm(
          usuarioActual.uid
        )
      );


    await setDoc(
      registroRef,
      {

        uid:
          usuarioActual.uid,

        correo:
          normalizarCorreo(
            usuarioActual.email
            ||
            perfilActual?.Correo
            ||
            ""
          ),

        installationId:
          installationId,

        dispositivoId:
          idDispositivo(),

        activo:
          true,

        actualizadoEn:
          serverTimestamp(),

        userAgent:
          navigator.userAgent
          ||
          "",

        plataforma:
          plataformaActual(),

        origen:
          "web"
      },
      {
        merge:
          true
      }
    );


    fidActual =
      installationId;

    canalNotificaciones =
      "fcm";

    notificacionesActivas =
      true;


    actualizarEstadoNotificaciones();
  }


  async function marcarRegistroFcmInactivo(
    installationId = fidActual
  ) {

    if (!usuarioActual) {
      return;
    }


    const registroRef =
      doc(
        db,
        "fcmRegistrations",
        idRegistroFcm(
          usuarioActual.uid
        )
      );


    try {

      await setDoc(
        registroRef,
        {

          uid:
            usuarioActual.uid,

          correo:
            normalizarCorreo(
              usuarioActual.email
              ||
              perfilActual?.Correo
              ||
              ""
            ),

          installationId:
            installationId
            ||
            "",

          dispositivoId:
            idDispositivo(),

          activo:
            false,

          actualizadoEn:
            serverTimestamp()
        },
        {
          merge:
            true
        }
      );

    }
    catch (err) {

      console.warn(
        "No se pudo marcar el registro FCM como inactivo:",
        err
      );
    }
  }


  async function prepararMessaging() {

    if (esIOS()) {
      return false;
    }


    if (
      messaging
      &&
      swRegistration
    ) {

      return true;
    }


    try {

      const compatible =
        await messagingIsSupported();


      if (!compatible) {

        actualizarEstadoNotificaciones();

        return false;
      }


      await prepararServiceWorker();


      messaging =
        getMessaging(app);


      if (
        !listenersMessagingConfigurados
      ) {

        listenersMessagingConfigurados =
          true;


        onRegistered(
          messaging,
          async installationId => {

            try {

              await guardarRegistroFcm(
                installationId
              );


              console.log(
                "[HHEE] FID registrado:",
                installationId
              );

            }
            catch (err) {

              console.error(
                "[HHEE] No se pudo guardar el FID:",
                err
              );


              toast(
                "El dispositivo se registró, pero Firestore rechazó guardar la notificación. Revisa las reglas.",
                "error"
              );
            }
          }
        );


        onUnregistered(
          messaging,
          async installationId => {

            try {

              await marcarRegistroFcmInactivo(
                installationId
              );

            }
            catch (err) {

              console.warn(
                "[HHEE] Error al desregistrar FCM:",
                err
              );
            }


            if (
              fidActual ===
              installationId
            ) {

              fidActual =
                "";
            }


            canalNotificaciones =
              "";

            notificacionesActivas =
              false;


            actualizarEstadoNotificaciones();
          }
        );


        onMessage(
          messaging,
          async payload => {

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
              new URL(
                "./",
                window.location.href
              ).href;


            toast(
              `${titulo}: ${cuerpo}`,
              "info"
            );


            if (
              Notification.permission ===
              "granted"
              &&
              swRegistration
            ) {

              try {

                await swRegistration
                  .showNotification(
                    titulo,
                    {

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
                    }
                  );

              }
              catch (err) {

                console.warn(
                  "[HHEE] No se pudo mostrar la notificación en primer plano:",
                  err
                );
              }
            }
          }
        );
      }


      return true;

    }
    catch (err) {

      console.error(
        "[HHEE] Error preparando Firebase Messaging:",
        err
      );


      actualizarEstadoNotificaciones();

      return false;
    }
  }


  async function activarNotificacionesFcm() {

    if (!usuarioActual) {

      return toast(
        "Primero inicia sesión.",
        "error"
      );
    }


    if (
      !("Notification" in window)
      ||
      !("serviceWorker" in navigator)
    ) {

      return toast(
        "Este navegador no admite notificaciones web.",
        "error"
      );
    }


    if (
      Notification.permission ===
      "denied"
    ) {

      actualizarEstadoNotificaciones();


      return toast(
        "Las notificaciones están bloqueadas. Habilítalas desde la configuración del navegador o del sitio.",
        "error"
      );
    }


    const permiso =
      Notification.permission ===
      "granted"
        ?
          "granted"
        :
          await Notification
            .requestPermission();


    if (
      permiso !==
      "granted"
    ) {

      actualizarEstadoNotificaciones();


      return toast(
        "No se otorgó permiso para mostrar notificaciones.",
        "error"
      );
    }


    const listo =
      await prepararMessaging();


    if (!listo) {

      return toast(
        "No fue posible iniciar Firebase Messaging en este dispositivo.",
        "error"
      );
    }


    try {

      $("btnNotificaciones")
        .disabled =
          true;


      $("btnNotificaciones")
        .textContent =
          "Activando...";


      await register(
        messaging,
        {

          vapidKey:
            VAPID_KEY,

          serviceWorkerRegistration:
            swRegistration
        }
      );


      toast(
        "Notificaciones activadas en este dispositivo.",
        "exito"
      );

    }
    catch (err) {

      console.error(
        "[HHEE] Error registrando FCM:",
        err
      );


      toast(
        "No fue posible registrar este dispositivo para notificaciones.",
        "error"
      );

    }
    finally {

      actualizarEstadoNotificaciones();
    }
  }


  async function desactivarNotificacionesFcm() {

    if (!messaging) {

      await prepararMessaging();
    }


    if (!messaging) {
      return;
    }


    if (
      !confirm(
        "¿Desactivar las notificaciones en este dispositivo?"
      )
    ) {

      return;
    }


    try {

      const fidAnterior =
        fidActual;


      await marcarRegistroFcmInactivo(
        fidAnterior
      );


      await unregister(
        messaging
      );


      fidActual =
        "";

      canalNotificaciones =
        "";

      notificacionesActivas =
        false;


      actualizarEstadoNotificaciones();


      toast(
        "Notificaciones desactivadas en este dispositivo.",
        "exito"
      );

    }
    catch (err) {

      console.error(
        "[HHEE] Error desactivando notificaciones:",
        err
      );


      toast(
        "No fue posible desactivar las notificaciones.",
        "error"
      );
    }
  }


  async function sincronizarWebPushSiCorresponde() {

    actualizarEstadoNotificaciones();


    if (
      !usuarioActual
      ||
      !esIOS()
      ||
      !esPwaInstalada()
      ||
      !soportaWebPushEstandar()
    ) {
      return;
    }


    try {

      /*
        Precargamos Service Worker y VAPID antes de que el usuario
        pulse el botón. Esto deja subscribe() preparado para ejecutarse
        directamente desde la interacción del usuario.
      */

      await Promise.all([
        prepararServiceWorker(),
        obtenerClavePublicaWebPush()
      ]);


      const {
        pushManager
      } =
        await obtenerPushManagerWeb();


      const subscription =
        await pushManager
          .getSubscription();


      if (subscription) {

        webPushSubscriptionActual =
          subscription;

        canalNotificaciones =
          "webpush";

        notificacionesActivas =
          true;


        await guardarRegistroWebPush(
          subscription
        );
      }

    }
    catch (err) {

      console.warn(
        "[HHEE] No se pudo preparar/sincronizar Web Push:",
        err
      );
    }


    actualizarEstadoNotificaciones();
  }


  async function sincronizarNotificacionesSiCorresponde() {

    actualizarEstadoNotificaciones();


    if (!usuarioActual) {
      return;
    }


    if (esIOS()) {

      await sincronizarWebPushSiCorresponde();

      return;
    }


    const listo =
      await prepararMessaging();


    if (!listo) {
      return;
    }


    if (
      Notification.permission ===
      "granted"
    ) {

      try {

        await register(
          messaging,
          {

            vapidKey:
              VAPID_KEY,

            serviceWorkerRegistration:
              swRegistration
          }
        );

      }
      catch (err) {

        console.warn(
          "[HHEE] No se pudo sincronizar el registro FCM:",
          err
        );
      }
    }
  }


  async function crearEventoNotificacion({
    tipo,
    solicitudId,
    destinatarioUid = "",
    destinatarioCorreo = "",
    titulo,
    cuerpo,
    url
  }) {

    if (!usuarioActual) {
      return;
    }


    try {

      await addDoc(
        collection(
          db,
          "notificationQueue"
        ),
        {

          tipo,

          solicitudId,

          destinatarioUid:
            destinatarioUid
            ||
            "",

          destinatarioCorreo:
            normalizarCorreo(
              destinatarioCorreo
              ||
              ""
            ),

          titulo,

          cuerpo,

          url,

          estado:
            "pendiente",

          creadoEn:
            serverTimestamp(),

          emisorUid:
            usuarioActual.uid,

          emisorCorreo:
            normalizarCorreo(
              usuarioActual.email
              ||
              ""
            )
        }
      );

    }
    catch (err) {

      console.warn(
        "[HHEE] La operación principal se guardó, pero no se pudo crear la cola de notificación:",
        err
      );
    }
  }


  async function cerrarSesionConNotificaciones() {

    try {

      if (esIOS()) {

        if (
          webPushSubscriptionActual
          ||
          notificacionesActivas
        ) {

          try {

            const {
              pushManager
            } =
              await obtenerPushManagerWeb();


            const subscription =
              webPushSubscriptionActual
              ||
              await pushManager
                .getSubscription();


            await marcarRegistroWebPushInactivo();


            if (subscription) {

              await subscription
                .unsubscribe();
            }

          }
          catch (err) {

            console.warn(
              "[HHEE] No se pudo cerrar Web Push:",
              err
            );
          }
        }

      }
      else if (
        messaging
        &&
        fidActual
      ) {

        const fidAnterior =
          fidActual;


        await marcarRegistroFcmInactivo(
          fidAnterior
        );


        try {

          await unregister(
            messaging
          );

        }
        catch (err) {

          console.warn(err);
        }
      }

    }
    finally {

      fidActual =
        "";

      webPushSubscriptionActual =
        null;

      canalNotificaciones =
        "";

      notificacionesActivas =
        false;


      await signOut(
        auth
      );
    }
  }


  $("btnNotificaciones")
    .addEventListener(
      "click",
      async () => {

        if (esIOS()) {

          if (
            notificacionesActivas
            &&
            webPushSubscriptionActual
          ) {

            await desactivarWebPushIOS();

          }
          else {

            await activarWebPushIOS();
          }


          return;
        }


        if (
          Notification.permission ===
          "denied"
        ) {

          actualizarEstadoNotificaciones();


          return toast(
            "Las notificaciones están bloqueadas. Habilítalas desde los permisos del sitio en tu navegador.",
            "error"
          );
        }


        if (
          notificacionesActivas
          &&
          fidActual
        ) {

          await desactivarNotificacionesFcm();

        }
        else {

          await activarNotificacionesFcm();
        }
      }
    );


  /* =========================================================
     FECHAS
  ========================================================= */

  function hoyIso() {

    const d =
      new Date();


    return `${d.getFullYear()}-${String(
      d.getMonth() + 1
    ).padStart(2,"0")}-${String(
      d.getDate()
    ).padStart(2,"0")}`;
  }


  function crearMomento(
    fecha,
    hora
  ) {

    const f =
      String(fecha)
        .split("-")
        .map(Number);


    const h =
      String(hora)
        .split(":")
        .map(Number);


    return Date.UTC(
      f[0],
      f[1] - 1,
      f[2],
      h[0],
      h[1]
    );
  }


  function inicioDia(fecha) {

    const f =
      String(fecha)
        .split("-")
        .map(Number);


    return Date.UTC(
      f[0],
      f[1] - 1,
      f[2]
    );
  }


  function isoDesdeMs(ms) {

    const d =
      new Date(ms);


    return `${d.getUTCFullYear()}-${String(
      d.getUTCMonth() + 1
    ).padStart(2,"0")}-${String(
      d.getUTCDate()
    ).padStart(2,"0")}`;
  }


  function horaDesdeMs(ms) {

    const d =
      new Date(ms);


    return `${String(
      d.getUTCHours()
    ).padStart(2,"0")}:${String(
      d.getUTCMinutes()
    ).padStart(2,"0")}`;
  }


  function sumarDias(
    fecha,
    n
  ) {

    return isoDesdeMs(
      inicioDia(fecha)
      +
      n * DIA
    );
  }


  function fechaCL(iso) {

    if (!iso) {
      return "";
    }


    const p =
      String(iso)
        .slice(0,10)
        .split("-");


    return p.length === 3
      ?
        `${p[2]}-${p[1]}-${p[0]}`
      :
        String(iso);
  }


  function timestampCL(v) {

    if (!v) {
      return "";
    }


    try {

      const d =
        typeof v.toDate ===
        "function"
          ?
            v.toDate()
          :
            new Date(v);


      if (
        Number.isNaN(
          d.getTime()
        )
      ) {

        return "";
      }


      return `${String(
        d.getDate()
      ).padStart(2,"0")}-${String(
        d.getMonth() + 1
      ).padStart(2,"0")}-${d.getFullYear()} ${String(
        d.getHours()
      ).padStart(2,"0")}:${String(
        d.getMinutes()
      ).padStart(2,"0")}`;

    }
    catch {

      return "";
    }
  }


  function nombreCompleto(o) {

    return `${o?.Nombre || ""} ${o?.Apellido || ""}`
      .replace(/\s+/g," ")
      .trim();
  }


  /* =========================================================
     PERFIL
  ========================================================= */

  async function obtenerPerfil(user) {

    const s =
      await getDoc(
        doc(
          db,
          "usuarios",
          user.uid
        )
      );


    if (!s.exists()) {

      throw new Error(
        "PERFIL_NO_EXISTE"
      );
    }


    const p =
      s.data();


    if (
      p.Activo !==
      true
    ) {

      throw new Error(
        "USUARIO_INACTIVO"
      );
    }


    return p;
  }


  function pintarPerfil(
    p,
    user
  ) {

    const nom =
      nombreCompleto(p)
      ||
      user.email
      ||
      "Usuario";


    $("nombreHeader")
      .textContent =
        nom;


    $("areaHeader")
      .textContent =
        p.Area
        ||
        "";


    $("nombreBienvenida")
      .textContent =
        p.Nombre
        ||
        nom;


    $("perfilRut")
      .textContent =
        p.Rut
        ||
        "-";


    $("perfilNumeroPersonal")
      .textContent =
        p.NumeroPersonal
        ||
        "-";


    $("perfilArea")
      .textContent =
        p.Area
        ||
        "-";


    $("perfilCorreo")
      .textContent =
        p.Correo
        ||
        user.email
        ||
        "-";


    $("avatarUsuario")
      .textContent =
        (
          (
            p.Nombre
            ||
            "U"
          ).charAt(0)
          +
          (
            p.Apellido
            ||
            ""
          ).charAt(0)
        )
        .toUpperCase();


    $("seccionTrabajador")
      .classList
      .toggle(
        "hidden",
        !puedeSolicitar()
      );


    $("seccionSupervisor")
      .classList
      .toggle(
        "hidden",
        !esSupervisor()
      );


    $("seccionAdministrador")
      .classList
      .add(
        "hidden"
      );


    $("barraAdminTabs")
      .classList
      .toggle(
        "hidden",
        !esAdmin()
      );


    $("introAdmin")
      .classList
      .toggle(
        "hidden",
        !esAdmin()
      );


    $("badgesUsuario")
      .innerHTML =
        `
          ${
            puedeSolicitar()
              ?
                '<span class="badge bTrab">Trabajador habilitado</span>'
              :
                ""
          }

          ${
            esSupervisor()
              ?
                '<span class="badge bSup">Supervisor</span>'
              :
                ""
          }

          ${
            esAdmin()
              ?
                '<span class="badge bAdm">Administrador</span>'
              :
                ""
          }
        `;
  }


  /* =========================================================
     LOGIN
  ========================================================= */

  $("formLogin")
    .addEventListener(
      "submit",
      async e => {

        e.preventDefault();


        $("mensajeLogin")
          .classList
          .add("hidden");


        $("btnLogin")
          .disabled =
            true;


        $("btnLogin")
          .textContent =
            "Ingresando...";


        try {

          await signInWithEmailAndPassword(
            auth,
            $("correo")
              .value
              .trim(),
            $("password")
              .value
          );

        }
        catch (err) {

          loginMsg(
            err.code ===
            "auth/invalid-credential"
              ?
                "Correo o contraseña incorrectos."
              :
                "No fue posible iniciar sesión."
          );

        }
        finally {

          $("btnLogin")
            .disabled =
              false;


          $("btnLogin")
            .textContent =
              "Ingresar";
        }
      }
    );


  $("verPassword")
    .addEventListener(
      "click",
      () => {

        $("password")
          .type =
            $("password")
              .type ===
            "password"
              ?
                "text"
              :
                "password";
      }
    );


  $("btnRecuperar")
    .addEventListener(
      "click",
      async () => {

        const c =
          $("correo")
            .value
            .trim();


        if (!c) {

          return loginMsg(
            "Escribe primero tu correo."
          );
        }


        try {

          await sendPasswordResetEmail(
            auth,
            c
          );


          loginMsg(
            "Si la cuenta está registrada, recibirás un correo de recuperación.",
            "exito"
          );

        }
        catch {

          loginMsg(
            "No fue posible enviar el correo de recuperación."
          );
        }
      }
    );


  /* =========================================================
     ACTIVAR CUENTA
  ========================================================= */

  function abrirActivacionCuenta() {

    modal(
      `
        <h3>
          Activar mi cuenta
        </h3>

        <div class="mensaje mensajeInfo">
          Tu correo debe haber sido autorizado previamente
          por Administración. No necesitas conocer ningún UID.
        </div>

        <form id="formActivacion">

          <div class="campo">

            <label>
              Correo corporativo *
            </label>

            <input
              id="actCorreo"
              type="email"
              autocomplete="username"
              value="${esc(
                normalizarCorreo(
                  $("correo").value
                )
              )}"
              required
            >

          </div>


          <div class="grid2">

            <div>

              <label>
                Crear contraseña *
              </label>

              <input
                id="actPassword"
                type="password"
                autocomplete="new-password"
                minlength="6"
                required
              >

              <div class="nota">
                Mínimo 6 caracteres.
              </div>

            </div>


            <div>

              <label>
                Confirmar contraseña *
              </label>

              <input
                id="actPassword2"
                type="password"
                autocomplete="new-password"
                minlength="6"
                required
              >

            </div>

          </div>


          <div class="acciones">

            <button
              type="button"
              class="btn btnSecundario"
              onclick="window.cerrarModal()"
            >
              Cancelar
            </button>

            <button
              id="btnConfirmarActivacion"
              class="btn btnPrimario"
              type="submit"
            >
              Activar cuenta
            </button>

          </div>

        </form>
      `
    );


    $("formActivacion")
      .addEventListener(
        "submit",
        activarCuenta
      );
  }


  $("btnActivarCuenta")
    .addEventListener(
      "click",
      abrirActivacionCuenta
    );


  async function activarCuenta(e) {

    e.preventDefault();


    const correo =
      normalizarCorreo(
        $("actCorreo").value
      );


    const password =
      $("actPassword").value;


    const password2 =
      $("actPassword2").value;


    if (!correo) {

      return toast(
        "Ingresa tu correo.",
        "error"
      );
    }


    if (
      password.length <
      6
    ) {

      return toast(
        "La contraseña debe tener al menos 6 caracteres.",
        "error"
      );
    }


    if (
      password !==
      password2
    ) {

      return toast(
        "Las contraseñas no coinciden.",
        "error"
      );
    }


    const btn =
      $("btnConfirmarActivacion");


    btn.disabled =
      true;


    btn.textContent =
      "Activando...";


    let creado =
      null;


    try {

      activacionEnProceso =
        true;


      const cred =
        await createUserWithEmailAndPassword(
          auth,
          correo,
          password
        );


      creado =
        cred.user;


      const preRef =
        doc(
          db,
          "preautorizados",
          correo
        );


      const preSnap =
        await getDoc(
          preRef
        );


      if (
        !preSnap.exists()
      ) {

        throw new Error(
          "NO_AUTORIZADO"
        );
      }


      const pre =
        preSnap.data();


      if (
        pre.Activo !==
        true
      ) {

        throw new Error(
          "AUTORIZACION_INACTIVA"
        );
      }


      if (
        pre.activado ===
        true
        &&
        pre.uid
        &&
        pre.uid !==
        creado.uid
      ) {

        throw new Error(
          "AUTORIZACION_USADA"
        );
      }


      const perfil = {

        Nombre:
          pre.Nombre
          ||
          "",

        Apellido:
          pre.Apellido
          ||
          "",

        Rut:
          pre.Rut
          ||
          "",

        NumeroPersonal:
          normalizarNumeroPersonal(
            pre.NumeroPersonal
            ||
            ""
          ),

        Area:
          pre.Area
          ||
          "",

        Correo:
          correo,

        supervisorcorreo:
          normalizarCorreo(
            pre.supervisorcorreo
            ||
            ""
          ),

        Activo:
          pre.Activo ===
          true,

        puedeSolicitarHE:
          pre.puedeSolicitarHE ===
          true,

        esSupervisor:
          pre.esSupervisor ===
          true,

        esAdministrador:
          pre.esAdministrador ===
          true,

        fechaActivacion:
          serverTimestamp(),

        origenActivacion:
          "preautorizacion"
      };


      const batch =
        writeBatch(db);


      batch.set(
        doc(
          db,
          "usuarios",
          creado.uid
        ),
        perfil
      );


      batch.update(
        preRef,
        {

          activado:
            true,

          uid:
            creado.uid,

          fechaActivacion:
            serverTimestamp()
        }
      );


      await batch.commit();


      perfilActual =
        await obtenerPerfil(
          creado
        );


      usuarioActual =
        creado;


      activacionEnProceso =
        false;


      cerrarModal();


      pintarPerfil(
        perfilActual,
        creado
      );


      mostrarApp();


      await sincronizarNotificacionesSiCorresponde();


      toast(
        "Cuenta activada correctamente. Ya puedes usar el sistema.",
        "exito"
      );

    }
    catch (err) {

      console.error(err);


      if (creado) {

        try {

          await deleteUser(
            creado
          );

        }
        catch (e2) {

          console.error(
            "No se pudo eliminar la cuenta incompleta",
            e2
          );
        }
      }


      activacionEnProceso =
        false;


      try {

        await signOut(
          auth
        );

      }
      catch {}


      mostrarLogin();


      const msg =
        err.code ===
        "auth/email-already-in-use"
          ?
            "Ese correo ya tiene una cuenta. Usa Ingresar u Olvidé mi contraseña."
          :
        err.code ===
        "auth/weak-password"
          ?
            "La contraseña es demasiado débil."
          :
        err.code ===
        "auth/invalid-email"
          ?
            "El correo no es válido."
          :
        err.message ===
        "NO_AUTORIZADO"
          ?
            "Tu correo todavía no ha sido autorizado por Administración."
          :
        err.message ===
        "AUTORIZACION_INACTIVA"
          ?
            "Tu autorización está desactivada. Contacta a Administración."
          :
        err.message ===
        "AUTORIZACION_USADA"
          ?
            "Esta autorización ya fue utilizada. Contacta a Administración."
          :
        err.code?.includes(
          "permission"
        )
          ?
            "No fue posible validar tu autorización por las reglas de Firestore. Publica las reglas nuevas."
          :
            "No fue posible activar la cuenta.";


      loginMsg(
        msg
      );

    }
    finally {

      if (btn) {

        btn.disabled =
          false;


        btn.textContent =
          "Activar cuenta";
      }
    }
  }


  $("btnSalir")
    .addEventListener(
      "click",
      cerrarSesionConNotificaciones
    );


  /* =========================================================
     CÁLCULO HHEE
  ========================================================= */

  function interseccion(
    lista,
    ini,
    fin,
    rIni,
    rFin,
    recargo,
    origen
  ) {

    const a =
      Math.max(
        ini,
        rIni
      );


    const b =
      Math.min(
        fin,
        rFin
      );


    if (
      b <= a
    ) {
      return;
    }


    lista.push({

      inicio:
        a,

      fin:
        b,

      minutos:
        Math.round(
          (
            b - a
          )
          /
          MINUTO
        ),

      recargo,

      origen
    });
  }


  function calcularNormal(
    ini,
    fin,
    origen
  ) {

    if (
      fin <= ini
    ) {

      throw new Error(
        "El término debe ser posterior al inicio."
      );
    }


    const t = [];


    const primer =
      Math.floor(
        ini / DIA
      )
      *
      DIA;


    const ultimo =
      Math.floor(
        (
          fin - 1
        )
        /
        DIA
      )
      *
      DIA;


    for (
      let dia = primer;
      dia <= ultimo;
      dia += DIA
    ) {

      const dow =
        new Date(
          dia
        )
        .getUTCDay();


      /* 00:00 a 08:00 = 100% */

      interseccion(
        t,
        ini,
        fin,
        dia,
        dia + 8 * 60 * MINUTO,
        100,
        origen
      );


      /* Lunes a jueves 17:30 a 24:00 = 50% */

      if (
        dow >= 1
        &&
        dow <= 4
      ) {

        interseccion(
          t,
          ini,
          fin,
          dia
          +
          (
            17 * 60
            +
            30
          )
          *
          MINUTO,
          dia + DIA,
          50,
          origen
        );
      }


      /* Viernes 14:30 a 24:00 = 50% */

      if (
        dow === 5
      ) {

        interseccion(
          t,
          ini,
          fin,
          dia
          +
          (
            14 * 60
            +
            30
          )
          *
          MINUTO,
          dia + DIA,
          50,
          origen
        );
      }


      /* Sábado 08:00 a 24:00 = 50% */

      if (
        dow === 6
      ) {

        interseccion(
          t,
          ini,
          fin,
          dia
          +
          8 * 60 * MINUTO,
          dia + DIA,
          50,
          origen
        );
      }
    }


    return validarCobertura(
      ini,
      fin,
      t
    );
  }


  function calcularEspecial(
    ini,
    fin,
    origen,
    fechaEspecial,
    tipoDia
  ) {

    if (
      fin <= ini
    ) {

      throw new Error(
        "El término debe ser posterior al inicio."
      );
    }


    if (
      tipoDia ===
      "Domingo"
      &&
      new Date(
        inicioDia(
          fechaEspecial
        )
      )
      .getUTCDay() !==
      0
    ) {

      throw new Error(
        "La fecha seleccionada no corresponde a domingo."
      );
    }


    const t = [];


    const primer =
      Math.floor(
        ini / DIA
      )
      *
      DIA;


    const ultimo =
      Math.floor(
        (
          fin - 1
        )
        /
        DIA
      )
      *
      DIA;


    /*
      00:00 a 08:00 = 100%
    */

    for (
      let dia = primer;
      dia <= ultimo;
      dia += DIA
    ) {

      interseccion(
        t,
        ini,
        fin,
        dia,
        dia + 8 * 60 * MINUTO,
        100,
        origen
      );
    }


    /*
      Domingo / feriado
      08:00 a 24:00 = 75%
    */

    const esp =
      inicioDia(
        fechaEspecial
      );


    interseccion(
      t,
      ini,
      fin,
      esp + 8 * 60 * MINUTO,
      esp + DIA,
      75,
      origen
    );


    return validarCobertura(
      ini,
      fin,
      t
    );
  }


  function validarCobertura(
    ini,
    fin,
    tramos
  ) {

    tramos.sort(
      (
        a,
        b
      ) =>
        a.inicio - b.inicio
    );


    const periodo =
      Math.round(
        (
          fin - ini
        )
        /
        MINUTO
      );


    const clas =
      tramos.reduce(
        (
          s,
          x
        ) =>
          s + x.minutos,
        0
      );


    const fuera =
      Math.max(
        0,
        periodo - clas
      );


    const suma =
      r =>
        tramos
          .filter(
            x =>
              x.recargo === r
          )
          .reduce(
            (
              s,
              x
            ) =>
              s + x.minutos,
            0
          );


    return {

      tramos,

      minutos50:
        suma(50),

      minutos75:
        suma(75),

      minutos100:
        suma(100),

      minutosTotales:
        clas,

      minutosFuera:
        fuera,

      minutosPeriodo:
        periodo
    };
  }


  function calcularSolicitudDesdeCampos(
    tipo
  ) {

    const fIni =
      $("fFechaInicio").value;

    const fFin =
      $("fFechaTermino").value;

    const hIni =
      $("fHoraInicio").value;

    const hFin =
      $("fHoraTermino").value;


    if (
      !fIni
      ||
      !fFin
      ||
      !hIni
      ||
      !hFin
    ) {

      return null;
    }


    const ini =
      crearMomento(
        fIni,
        hIni
      );


    const fin =
      crearMomento(
        fFin,
        hFin
      );


    let principal;


    if (
      tipo ===
      "Normal"
    ) {

      principal =
        calcularNormal(
          ini,
          fin,
          "Jornada principal"
        );

    }
    else {

      principal =
        calcularEspecial(
          ini,
          fin,
          "Jornada principal",
          fIni,
          $("fDiaTrabajo").value
        );
    }


    let ext =
      null;

    let fechaExtIni =
      "";

    let fechaExtFin =
      "";

    let iniExt =
      null;

    let finExt =
      null;


    if (
      $("fTieneExtension").checked
    ) {

      const mh =
        $("fMotivoExtension").value;

      const hi =
        $("fHoraInicioExtension").value;

      const hf =
        $("fHoraTerminoExtension").value;


      if (
        !mh
        ||
        !hi
        ||
        !hf
      ) {

        throw new Error(
          "Completa motivo y horas de la extensión."
        );
      }


      fechaExtIni =
        fFin;


      iniExt =
        crearMomento(
          fechaExtIni,
          hi
        );


      if (
        iniExt < fin
      ) {

        fechaExtIni =
          sumarDias(
            fechaExtIni,
            1
          );


        iniExt =
          crearMomento(
            fechaExtIni,
            hi
          );
      }


      fechaExtFin =
        fechaExtIni;


      finExt =
        crearMomento(
          fechaExtFin,
          hf
        );


      if (
        finExt <=
        iniExt
      ) {

        fechaExtFin =
          sumarDias(
            fechaExtFin,
            1
          );


        finExt =
          crearMomento(
            fechaExtFin,
            hf
          );
      }


      if (
        tipo ===
        "Normal"
      ) {

        ext =
          calcularNormal(
            iniExt,
            finExt,
            "Extensión"
          );

      }
      else {

        ext =
          calcularEspecial(
            iniExt,
            finExt,
            "Extensión",
            fIni,
            $("fDiaTrabajo").value
          );
      }
    }


    const tramos = [
      ...principal.tramos,
      ...(
        ext?.tramos
        ||
        []
      )
    ];


    const minutos50 =
      principal.minutos50
      +
      (
        ext?.minutos50
        ||
        0
      );


    const minutos75 =
      principal.minutos75
      +
      (
        ext?.minutos75
        ||
        0
      );


    const minutos100 =
      principal.minutos100
      +
      (
        ext?.minutos100
        ||
        0
      );


    const minutosFuera =
      principal.minutosFuera
      +
      (
        ext?.minutosFuera
        ||
        0
      );


    return {

      tramos,

      minutos50,

      minutos75,

      minutos100,

      minutosTotales:
        minutos50
        +
        minutos75
        +
        minutos100,

      minutosFuera,

      fechaExtIni,

      fechaExtFin
    };
  }


  function renderCalculo(tipo) {

    const caja =
      $("vistaCalculo");


    if (!caja) {
      return;
    }


    try {

      const c =
        calcularSolicitudDesdeCampos(
          tipo
        );


      if (!c) {

        caja.innerHTML =
          `
            <div class="mensaje mensajeInfo">
              Completa fecha y horario para ver el cálculo.
            </div>
          `;

        return;
      }


      const aviso =
        c.minutosFuera > 0
          ?
            `
              <div class="mensaje mensajeError">
                Hay ${c.minutosFuera} minutos fuera de las reglas definidas.
                Ajusta el horario antes de enviar.
              </div>
            `
          :
            "";


      caja.innerHTML =
        `
          <div class="resumenKpi">

            <div class="kpi">
              <span>50%</span>
              <strong>${vistaHoras(minutosHoras(c.minutos50))} h</strong>
            </div>

            <div class="kpi">
              <span>75%</span>
              <strong>${vistaHoras(minutosHoras(c.minutos75))} h</strong>
            </div>

            <div class="kpi">
              <span>100%</span>
              <strong>${vistaHoras(minutosHoras(c.minutos100))} h</strong>
            </div>

            <div class="kpi">
              <span>Total</span>
              <strong>${vistaHoras(minutosHoras(c.minutosTotales))} h</strong>
            </div>

          </div>

          ${aviso}
        `;

    }
    catch (err) {

      caja.innerHTML =
        `
          <div class="mensaje mensajeError">
            ${esc(err.message)}
          </div>
        `;
    }
  }


  /* =========================================================
     FORMULARIO SOLICITUD
  ========================================================= */

  function formularioSolicitud(
    tipo = "Normal",
    data = null,
    id = null
  ) {

    const especial =
      tipo ===
      "DomingoFeriado";


    return `
      <form id="formSolicitud">

        <div class="bloque">

          <h3>
            Datos del trabajador
          </h3>

          <div class="grid3">

            <div>
              <label>Nombre</label>
              <input
                value="${esc(perfilActual.Nombre || "")}"
                readonly
              >
            </div>

            <div>
              <label>Apellido</label>
              <input
                value="${esc(perfilActual.Apellido || "")}"
                readonly
              >
            </div>

            <div>
              <label>RUT</label>
              <input
                value="${esc(perfilActual.Rut || "")}"
                readonly
              >
            </div>

            <div>
              <label>N° de personal</label>
              <input
                value="${esc(perfilActual.NumeroPersonal || "")}"
                readonly
              >
            </div>

            <div>
              <label>Área</label>
              <input
                value="${esc(perfilActual.Area || "")}"
                readonly
              >
            </div>

            <div>
              <label>Correo</label>
              <input
                value="${esc(perfilActual.Correo || usuarioActual.email || "")}"
                readonly
              >
            </div>

            <div>
              <label>Supervisor</label>
              <input
                value="${esc(perfilActual.supervisorcorreo || "")}"
                readonly
              >
            </div>

          </div>

        </div>


        <div class="bloque">

          <h3>
            ${
              especial
                ?
                  "Trabajo domingo o feriado"
                :
                  "Jornada de horas extras"
            }
          </h3>

          <div class="grid2">

            ${
              especial
                ?
                  `
                    <div>

                      <label>
                        Día de trabajo *
                      </label>

                      <select
                        id="fDiaTrabajo"
                        required
                      >

                        <option value="">
                          Seleccione
                        </option>

                        <option
                          ${
                            data?.diaTrabajo ===
                            "Domingo"
                              ?
                                "selected"
                              :
                                ""
                          }
                        >
                          Domingo
                        </option>

                        <option
                          ${
                            data?.diaTrabajo ===
                            "Feriado"
                              ?
                                "selected"
                              :
                                ""
                          }
                        >
                          Feriado
                        </option>

                      </select>

                    </div>
                  `
                :
                  ""
            }


            <div>

              <label>
                Fecha inicio *
              </label>

              <input
                id="fFechaInicio"
                type="date"
                value="${esc(data?.fechaInicio || hoyIso())}"
                required
              >

            </div>


            <div>

              <label>
                Fecha término *
              </label>

              <input
                id="fFechaTermino"
                type="date"
                value="${esc(data?.fechaTermino || hoyIso())}"
                required
              >

            </div>


            <div>

              <label>
                Hora inicio *
              </label>

              <input
                id="fHoraInicio"
                type="time"
                value="${esc(data?.horaInicio || "")}"
                required
              >

            </div>


            <div>

              <label>
                Hora término *
              </label>

              <input
                id="fHoraTermino"
                type="time"
                value="${esc(data?.horaTermino || "")}"
                required
              >

            </div>


            <div>

              <label>
                Tipo / motivo de trabajo *
              </label>

              <select
                id="fMotivoTrabajo"
                required
              >

                <option value="">
                  Seleccione
                </option>

                ${
                  [
                    "Emergencia",
                    "Planificado",
                    "Terceros"
                  ]
                  .map(
                    x =>
                      `
                        <option
                          ${
                            data?.motivoTrabajo ===
                            x
                              ?
                                "selected"
                              :
                                ""
                          }
                        >
                          ${x}
                        </option>
                      `
                  )
                  .join("")
                }

              </select>

            </div>


            <div>

              <label>
                Orden Interna CAPEX
              </label>

              <input
                id="fCapex"
                value="${esc(data?.ordenInternaCAPEX || "")}"
              >

            </div>


            <div class="full">

              <label>
                Justificación *
              </label>

              <textarea
                id="fJustificacion"
                required
              >${esc(data?.justificacion || "")}</textarea>

            </div>

          </div>

        </div>


        <div class="bloque">

          <h3>
            Extensión de faena
          </h3>


          <div class="checkRow">

            <input
              id="fTieneExtension"
              type="checkbox"
              ${
                data?.tieneExtension
                  ?
                    "checked"
                  :
                    ""
              }
            >

            <label for="fTieneExtension">
              La faena de trabajo se extendió
            </label>

          </div>


          <div
            id="fExtension"
            class="extension ${
              data?.tieneExtension
                ?
                  ""
                :
                  "hidden"
            }"
          >

            <div class="grid3">

              <div>

                <label>
                  Motivo extensión *
                </label>

                <select id="fMotivoExtension">

                  <option value="">
                    Seleccione
                  </option>

                  ${
                    [
                      "Programado",
                      "Emergencia"
                    ]
                    .map(
                      x =>
                        `
                          <option
                            ${
                              data?.motivoExtension ===
                              x
                                ?
                                  "selected"
                                :
                                  ""
                            }
                          >
                            ${x}
                          </option>
                        `
                    )
                    .join("")
                  }

                </select>

              </div>


              <div>

                <label>
                  Hora inicio extensión *
                </label>

                <input
                  id="fHoraInicioExtension"
                  type="time"
                  value="${esc(data?.horaInicioExtension || "")}"
                >

              </div>


              <div>

                <label>
                  Hora término extensión *
                </label>

                <input
                  id="fHoraTerminoExtension"
                  type="time"
                  value="${esc(data?.horaTerminoExtension || "")}"
                >

              </div>

            </div>

          </div>

        </div>


        <div class="bloque">

          <h3>
            Cálculo automático
          </h3>

          <div id="vistaCalculo"></div>

        </div>


        <div class="acciones">

          <button
            class="btn btnSecundario"
            type="button"
            id="fCancelar"
          >
            Cancelar
          </button>

          <button
            class="btn btnPrimario"
            type="submit"
            id="fGuardar"
          >
            ${
              id
                ?
                  "Guardar cambios"
                :
                  "Enviar solicitud"
            }
          </button>

        </div>

      </form>
    `;
  }


  function configurarFormularioSolicitud(
    tipo,
    id = null
  ) {

    const eventos = [
      "fFechaInicio",
      "fFechaTermino",
      "fHoraInicio",
      "fHoraTermino",
      "fDiaTrabajo",
      "fHoraInicioExtension",
      "fHoraTerminoExtension"
    ];


    eventos.forEach(
      k =>
        $(k)
          ?.addEventListener(
            "change",
            () =>
              renderCalculo(
                tipo
              )
          )
    );


    $("fTieneExtension")
      .addEventListener(
        "change",
        e => {

          $("fExtension")
            .classList
            .toggle(
              "hidden",
              !e.target.checked
            );


          renderCalculo(
            tipo
          );
        }
      );


    $("fMotivoExtension")
      .addEventListener(
        "change",
        () =>
          renderCalculo(
            tipo
          )
      );


    $("fCancelar")
      .addEventListener(
        "click",
        () => {

          if (id) {

            abrirModulo(
              "misSolicitudes"
            );

          }
          else {

            cerrarPanel();
          }
        }
      );


    $("formSolicitud")
      .addEventListener(
        "submit",
        e =>
          guardarSolicitud(
            e,
            tipo,
            id
          )
      );


    renderCalculo(
      tipo
    );
  }


  /* =========================================================
     GUARDAR SOLICITUD
  ========================================================= */

  async function guardarSolicitud(
    e,
    tipo,
    id = null
  ) {

    e.preventDefault();


    if (
      !puedeSolicitar()
    ) {

      return toast(
        "Este usuario no está habilitado para solicitar HHEE.",
        "error"
      );
    }


    if (
      !perfilActual.supervisorcorreo
    ) {

      return toast(
        "Tu perfil no tiene supervisor asignado.",
        "error"
      );
    }


    let c;


    try {

      c =
        calcularSolicitudDesdeCampos(
          tipo
        );

    }
    catch (err) {

      return toast(
        err.message,
        "error"
      );
    }


    if (
      !c
      ||
      c.minutosTotales <=
      0
    ) {

      return toast(
        "No hay horas extras válidas en el horario ingresado.",
        "error"
      );
    }


    if (
      c.minutosFuera >
      0
    ) {

      return toast(
        "Hay tiempo fuera de las reglas definidas. Ajusta el horario.",
        "error"
      );
    }


    const btn =
      $("fGuardar");


    btn.disabled =
      true;


    btn.textContent =
      "Guardando...";


    try {

      const tramos =
        c.tramos.map(
          t => ({

            origen:
              t.origen,

            fecha:
              isoDesdeMs(
                t.inicio
              ),

            horaInicio:
              horaDesdeMs(
                t.inicio
              ),

            horaTermino:
              horaDesdeMs(
                t.fin
              ),

            minutos:
              t.minutos,

            horas:
              minutosHoras(
                t.minutos
              ),

            recargo:
              t.recargo
          })
        );


      const datos = {

        tipoFormulario:
          tipo,

        diaTrabajo:
          tipo ===
          "DomingoFeriado"
            ?
              $("fDiaTrabajo").value
            :
              "",

        trabajadorUid:
          usuarioActual.uid,

        correoTrabajador:
          usuarioActual.email,

        Nombre:
          perfilActual.Nombre
          ||
          "",

        Apellido:
          perfilActual.Apellido
          ||
          "",

        Rut:
          perfilActual.Rut
          ||
          "",

        NumeroPersonal:
          normalizarNumeroPersonal(
            perfilActual.NumeroPersonal
            ||
            ""
          ),

        Area:
          perfilActual.Area
          ||
          "",

        supervisorCorreo:
          perfilActual.supervisorcorreo,

        estado:
          "pendiente",

        fechaInicio:
          $("fFechaInicio").value,

        fechaTermino:
          $("fFechaTermino").value,

        horaInicio:
          $("fHoraInicio").value,

        horaTermino:
          $("fHoraTermino").value,

        motivoTrabajo:
          $("fMotivoTrabajo").value,

        justificacion:
          $("fJustificacion")
            .value
            .trim(),

        ordenInternaCAPEX:
          $("fCapex")
            .value
            .trim(),

        tieneExtension:
          $("fTieneExtension").checked,

        motivoExtension:
          $("fTieneExtension").checked
            ?
              $("fMotivoExtension").value
            :
              "",

        fechaInicioExtension:
          $("fTieneExtension").checked
            ?
              c.fechaExtIni
            :
              "",

        fechaTerminoExtension:
          $("fTieneExtension").checked
            ?
              c.fechaExtFin
            :
              "",

        horaInicioExtension:
          $("fTieneExtension").checked
            ?
              $("fHoraInicioExtension").value
            :
              "",

        horaTerminoExtension:
          $("fTieneExtension").checked
            ?
              $("fHoraTerminoExtension").value
            :
              "",

        minutos50:
          c.minutos50,

        minutos75:
          c.minutos75,

        minutos100:
          c.minutos100,

        minutosTotales:
          c.minutosTotales,

        horas50:
          minutosHoras(
            c.minutos50
          ),

        horas75:
          minutosHoras(
            c.minutos75
          ),

        horas100:
          minutosHoras(
            c.minutos100
          ),

        horasTotales:
          minutosHoras(
            c.minutosTotales
          ),

        tramosCalculados:
          tramos
      };


      if (id) {

        const original =
          cacheSolicitudes.get(
            id
          );


        if (
          !original
          ||
          original.trabajadorUid !==
          usuarioActual.uid
          ||
          normEstado(
            original.estado
          ) !==
          "pendiente"
        ) {

          throw new Error(
            "Esta solicitud ya no se puede modificar."
          );
        }


        await updateDoc(
          doc(
            db,
            "solicitudes",
            id
          ),
          {

            ...datos,

            fechaActualizacion:
              serverTimestamp()
          }
        );


        toast(
          "Solicitud actualizada.",
          "exito"
        );


        abrirModulo(
          "misSolicitudes"
        );

      }
      else {

        /*
          NUEVA SOLICITUD
        */

        const nuevaRef =
          await addDoc(
            collection(
              db,
              "solicitudes"
            ),
            {

              ...datos,

              fechaRegistro:
                serverTimestamp(),

              fechaDecision:
                null,

              decisionPor:
                "",

              comentarioSupervisor:
                ""
            }
          );


        /*
          NUEVA NOTIFICACIÓN AL SUPERVISOR
        */

        await crearEventoNotificacion({

          tipo:
            "nueva_solicitud",

          solicitudId:
            nuevaRef.id,

          destinatarioCorreo:
            perfilActual.supervisorcorreo,

          titulo:
            "Nueva solicitud de horas extras",

          cuerpo:
            `${nombreCompleto(perfilActual)} envió una solicitud para ${fechaCL(datos.fechaInicio)}.`,

          url:
            urlModuloNotificacion(
              "pendientesSupervisor",
              nuevaRef.id
            )
        });


        toast(
          "Solicitud enviada y pendiente de revisión.",
          "exito"
        );


        abrirModulo(
          "misSolicitudes"
        );
      }

    }
    catch (err) {

      console.error(err);


      toast(
        err.code?.includes(
          "permission"
        )
          ?
            "Firestore rechazó la operación por permisos."
          :
            err.message
            ||
            "No fue posible guardar.",
        "error"
      );

    }
    finally {

      btn.disabled =
        false;


      btn.textContent =
        id
          ?
            "Guardar cambios"
          :
            "Enviar solicitud";
    }
  }


  /* =========================================================
     SOLICITUDES
  ========================================================= */

  async function cargarSolicitudesPropias() {

    const q =
      query(
        collection(
          db,
          "solicitudes"
        ),
        where(
          "trabajadorUid",
          "==",
          usuarioActual.uid
        )
      );


    const s =
      await getDocs(q);


    const arr =
      s.docs.map(
        d => ({
          id:
            d.id,
          ...d.data()
        })
      );


    arr.sort(
      (
        a,
        b
      ) =>
        `${b.fechaInicio || ""}${b.horaInicio || ""}`
          .localeCompare(
            `${a.fechaInicio || ""}${a.horaInicio || ""}`
          )
    );


    arr.forEach(
      x =>
        cacheSolicitudes.set(
          x.id,
          x
        )
    );


    return arr;
  }


  async function cargarSolicitudesSupervisor() {

    const q =
      query(
        collection(
          db,
          "solicitudes"
        ),
        where(
          "supervisorCorreo",
          "==",
          correoActual()
        )
      );


    const s =
      await getDocs(q);


    const arr =
      s.docs.map(
        d => ({
          id:
            d.id,
          ...d.data()
        })
      );


    arr.sort(
      (
        a,
        b
      ) =>
        `${b.fechaInicio || ""}${b.horaInicio || ""}`
          .localeCompare(
            `${a.fechaInicio || ""}${a.horaInicio || ""}`
          )
    );


    arr.forEach(
      x =>
        cacheSolicitudes.set(
          x.id,
          x
        )
    );


    return arr;
  }


  async function cargarTodasSolicitudes() {

    const s =
      await getDocs(
        collection(
          db,
          "solicitudes"
        )
      );


    const arr =
      s.docs.map(
        d => ({
          id:
            d.id,
          ...d.data()
        })
      );


    arr.sort(
      (
        a,
        b
      ) =>
        `${b.fechaInicio || ""}${b.horaInicio || ""}`
          .localeCompare(
            `${a.fechaInicio || ""}${a.horaInicio || ""}`
          )
    );


    arr.forEach(
      x =>
        cacheSolicitudes.set(
          x.id,
          x
        )
    );


    return arr;
  }


  function estadoHtml(e) {

    const n =
      normEstado(e);


    return `
      <span class="estado ${
        n === "aprobada"
          ?
            "aprobada"
          :
        n === "rechazada"
          ?
            "rechazada"
          :
            "pendiente"
      }">
        ${
          n === "aprobada"
            ?
              "Aprobada"
            :
          n === "rechazada"
            ?
              "Rechazada"
            :
              "Pendiente"
        }
      </span>
    `;
  }


  function tablaSolicitudes(
    lista,
    modo
  ) {

    if (
      !lista.length
    ) {

      return `
        <div class="mensaje mensajeInfo">
          No hay solicitudes para mostrar.
        </div>
      `;
    }


    return `
      <div class="tablaWrap">

        <table class="tabla">

          <thead>

            <tr>
              <th>Trabajador</th>
              <th>Fecha</th>
              <th>Horario</th>
              <th>Tipo</th>
              <th>50%</th>
              <th>75%</th>
              <th>100%</th>
              <th>Estado</th>
              <th>Acciones</th>
            </tr>

          </thead>

          <tbody>

            ${
              lista.map(
                s =>
                  `
                    <tr>

                      <td>
                        ${esc(nombreCompleto(s))}
                        <br>
                        <small>
                          ${esc(s.NumeroPersonal ? `N° ${s.NumeroPersonal} · ` : "")}${esc(s.Rut || "")}
                        </small>
                      </td>

                      <td>
                        ${fechaCL(s.fechaInicio)}
                      </td>

                      <td>
                        ${esc(s.horaInicio || "")}
                        -
                        ${esc(s.horaTermino || "")}
                      </td>

                      <td>
                        ${esc(s.tipoFormulario || "")}
                      </td>

                      <td>
                        ${vistaHoras(s.horas50 || 0)}
                      </td>

                      <td>
                        ${vistaHoras(s.horas75 || 0)}
                      </td>

                      <td>
                        ${vistaHoras(s.horas100 || 0)}
                      </td>

                      <td>
                        ${estadoHtml(s.estado)}
                      </td>

                      <td>

                        <div class="miniAcciones">

                          <button
                            class="btn btnSecundario"
                            onclick="window.verSolicitud('${s.id}')"
                          >
                            Ver
                          </button>

                          ${
                            modo ===
                            "trabajador"
                            &&
                            normEstado(
                              s.estado
                            ) ===
                            "pendiente"
                              ?
                                `
                                  <button
                                    class="btn btnSecundario"
                                    onclick="window.editarSolicitud('${s.id}')"
                                  >
                                    Editar
                                  </button>

                                  <button
                                    class="btn btnPeligro"
                                    onclick="window.eliminarSolicitud('${s.id}')"
                                  >
                                    Eliminar
                                  </button>
                                `
                              :
                                ""
                          }

                          ${
                            modo ===
                            "supervisor"
                            &&
                            normEstado(
                              s.estado
                            ) ===
                            "pendiente"
                              ?
                                `
                                  <button
                                    class="btn btnExito"
                                    onclick="window.decisionSupervisor('${s.id}','aprobada')"
                                  >
                                    Aprobar
                                  </button>

                                  <button
                                    class="btn btnPeligro"
                                    onclick="window.decisionSupervisor('${s.id}','rechazada')"
                                  >
                                    Rechazar
                                  </button>
                                `
                              :
                                ""
                          }

                        </div>

                      </td>

                    </tr>
                  `
              )
              .join("")
            }

          </tbody>

        </table>

      </div>
    `;
  }


  window.verSolicitud =
    id => {

      const s =
        cacheSolicitudes.get(
          id
        );


      if (!s) {
        return;
      }


      const puedeRevisarEstaSolicitud =
        esSupervisor()
        &&
        normEstado(
          s.estado
        ) ===
        "pendiente"
        &&
        normalizarCorreo(
          s.supervisorCorreo
          ||
          ""
        ) ===
        normalizarCorreo(
          correoActual()
        );


      const esSolicitudPropia =
        s.trabajadorUid ===
        usuarioActual?.uid;


      const tr =
        (
          s.tramosCalculados
          ||
          []
        )
        .map(
          x =>
            `
              <tr>

                <td>
                  ${fechaCL(x.fecha)}
                </td>

                <td>
                  ${esc(x.horaInicio)}
                  -
                  ${esc(x.horaTermino)}
                </td>

                <td>
                  ${x.recargo}%
                </td>

                <td>
                  ${vistaHoras(x.horas || 0)}
                </td>

                <td>
                  ${esc(x.origen || "")}
                </td>

              </tr>
            `
        )
        .join("");


      modal(
        `
          <h3>
            Detalle de solicitud
          </h3>

          <div class="grid2">

            <div>
              <b>Trabajador</b>
              <br>
              ${esc(nombreCompleto(s))}
            </div>

            <div>
              <b>RUT</b>
              <br>
              ${esc(s.Rut || "")}
            </div>

            <div>
              <b>N° de personal</b>
              <br>
              ${esc(s.NumeroPersonal || "-")}
            </div>

            <div>
              <b>Área</b>
              <br>
              ${esc(s.Area || "")}
            </div>

            <div>
              <b>Estado</b>
              <br>
              ${estadoHtml(s.estado)}
            </div>

            <div>
              <b>Fecha</b>
              <br>
              ${fechaCL(s.fechaInicio)}
              a
              ${fechaCL(s.fechaTermino)}
            </div>

            <div>
              <b>Horario</b>
              <br>
              ${esc(s.horaInicio || "")}
              -
              ${esc(s.horaTermino || "")}
            </div>

            <div>
              <b>Motivo</b>
              <br>
              ${esc(s.motivoTrabajo || "")}
            </div>

            <div>
              <b>Supervisor</b>
              <br>
              ${esc(s.supervisorCorreo || "")}
            </div>

            <div class="full">
              <b>Justificación</b>
              <br>
              ${esc(s.justificacion || "")}
            </div>

            <div class="full">
              <b>Comentario supervisor</b>
              <br>
              ${esc(s.comentarioSupervisor || "-")}
            </div>

          </div>


          <div class="detalleTramos">

            <h4>
              Tramos calculados
            </h4>

            <div class="tablaWrap">

              <table class="tabla">

                <thead>

                  <tr>
                    <th>Fecha</th>
                    <th>Horario</th>
                    <th>Recargo</th>
                    <th>Horas</th>
                    <th>Origen</th>
                  </tr>

                </thead>

                <tbody>

                  ${
                    tr
                    ||
                    `
                      <tr>
                        <td colspan="5">
                          Sin detalle
                        </td>
                      </tr>
                    `
                  }

                </tbody>

              </table>

            </div>

          </div>


          <div class="acciones">

            ${
              esSolicitudPropia
              &&
              normEstado(
                s.estado
              ) ===
              "pendiente"
                ?
                  `
                    <button
                      class="btn btnSecundario"
                      onclick="window.cerrarModal();window.editarSolicitud('${s.id}')"
                    >
                      Editar
                    </button>
                  `
                :
                  ""
            }

            ${
              puedeRevisarEstaSolicitud
                ?
                  `
                    <button
                      class="btn btnExito"
                      onclick="window.decisionSupervisor('${s.id}','aprobada')"
                    >
                      Aprobar
                    </button>

                    <button
                      class="btn btnPeligro"
                      onclick="window.decisionSupervisor('${s.id}','rechazada')"
                    >
                      Rechazar
                    </button>
                  `
                :
                  ""
            }

            <button
              class="btn btnSecundario"
              onclick="window.cerrarModal()"
            >
              Cerrar
            </button>

          </div>
        `
      );
    };


  window.editarSolicitud =
    id => {

      const s =
        cacheSolicitudes.get(
          id
        );


      if (
        !s
        ||
        s.trabajadorUid !==
        usuarioActual.uid
        ||
        normEstado(
          s.estado
        ) !==
        "pendiente"
      ) {

        return toast(
          "Solo puedes editar tus solicitudes pendientes.",
          "error"
        );
      }


      $("tituloModulo")
        .textContent =
          "Editar solicitud";


      $("descripcionModulo")
        .textContent =
          "Modifica la información antes de que el supervisor la revise.";


      $("contenidoModulo")
        .innerHTML =
          formularioSolicitud(
            s.tipoFormulario
            ||
            "Normal",
            s,
            id
          );


      configurarFormularioSolicitud(
        s.tipoFormulario
        ||
        "Normal",
        id
      );


      $("panelModulo")
        .scrollIntoView({
          behavior:
            "smooth"
        });
    };


  window.eliminarSolicitud =
    async id => {

      const s =
        cacheSolicitudes.get(
          id
        );


      if (
        !s
        ||
        s.trabajadorUid !==
        usuarioActual.uid
        ||
        normEstado(
          s.estado
        ) !==
        "pendiente"
      ) {

        return toast(
          "Solo puedes eliminar solicitudes pendientes.",
          "error"
        );
      }


      if (
        !confirm(
          "¿Eliminar esta solicitud pendiente?"
        )
      ) {

        return;
      }


      try {

        await deleteDoc(
          doc(
            db,
            "solicitudes",
            id
          )
        );


        toast(
          "Solicitud eliminada.",
          "exito"
        );


        abrirModulo(
          "misSolicitudes"
        );

      }
      catch (err) {

        console.error(err);


        toast(
          "No fue posible eliminar la solicitud.",
          "error"
        );
      }
    };


  /* =========================================================
     DECISIÓN SUPERVISOR
  ========================================================= */

  window.decisionSupervisor =
    (
      id,
      estado
    ) => {

      const s =
        cacheSolicitudes.get(
          id
        );


      if (
        !s
        ||
        normEstado(
          s.estado
        ) !==
        "pendiente"
      ) {

        return toast(
          "La solicitud ya fue revisada.",
          "error"
        );
      }


      const requiere =
        estado ===
        "rechazada";


      modal(
        `
          <h3>
            ${
              estado ===
              "aprobada"
                ?
                  "Aprobar"
                :
                  "Rechazar"
            }
            solicitud
          </h3>

          <p>
            ${esc(nombreCompleto(s))}
            ·
            ${fechaCL(s.fechaInicio)}
            ·
            ${esc(s.horaInicio)}
            -
            ${esc(s.horaTermino)}
          </p>

          <div class="campo">

            <label>
              Comentario ${
                requiere
                  ?
                    "*"
                  :
                    "(opcional)"
              }
            </label>

            <textarea id="comentarioDecision"></textarea>

          </div>

          <div class="acciones">

            <button
              class="btn btnSecundario"
              onclick="window.cerrarModal()"
            >
              Cancelar
            </button>

            <button
              class="btn ${
                estado ===
                "aprobada"
                  ?
                    "btnExito"
                  :
                    "btnPeligro"
              }"
              onclick="window.confirmarDecision('${id}','${estado}')"
            >
              Confirmar
            </button>

          </div>
        `
      );
    };


  window.confirmarDecision =
    async (
      id,
      estado
    ) => {

      const comentario =
        $("comentarioDecision")
          .value
          .trim();


      if (
        estado ===
        "rechazada"
        &&
        !comentario
      ) {

        return toast(
          "Indica el motivo del rechazo.",
          "error"
        );
      }


      const s =
        cacheSolicitudes.get(
          id
        );


      if (!s) {

        return toast(
          "No se encontró la solicitud.",
          "error"
        );
      }


      try {

        await updateDoc(
          doc(
            db,
            "solicitudes",
            id
          ),
          {

            estado,

            fechaDecision:
              serverTimestamp(),

            decisionPor:
              correoActual(),

            comentarioSupervisor:
              comentario
          }
        );


        /*
          NOTIFICACIÓN AL TRABAJADOR
        */

        await crearEventoNotificacion({

          tipo:
            estado ===
            "aprobada"
              ?
                "solicitud_aprobada"
              :
                "solicitud_rechazada",

          solicitudId:
            id,

          destinatarioUid:
            s.trabajadorUid
            ||
            "",

          destinatarioCorreo:
            s.correoTrabajador
            ||
            "",

          titulo:
            estado ===
            "aprobada"
              ?
                "Solicitud de HHEE aprobada"
              :
                "Solicitud de HHEE rechazada",

          cuerpo:
            estado ===
            "aprobada"
              ?
                `Tu solicitud del ${fechaCL(s.fechaInicio)} fue aprobada.`
              :
                `Tu solicitud del ${fechaCL(s.fechaInicio)} fue rechazada.${
                  comentario
                    ?
                      ` Motivo: ${comentario}`
                    :
                      ""
                }`,

          url:
            urlModuloNotificacion(
              "misSolicitudes",
              id
            )
        });


        cerrarModal();


        toast(
          `Solicitud ${estado}.`,
          "exito"
        );


        abrirModulo(
          "pendientesSupervisor"
        );

      }
      catch (err) {

        console.error(err);


        toast(
          "No fue posible registrar la decisión.",
          "error"
        );
      }
    };


  async function renderMisSolicitudes() {

    const arr =
      await cargarSolicitudesPropias();


    $("contenidoModulo")
      .innerHTML =
        `
          <div class="resumenKpi">

            <div class="kpi">
              <span>Total</span>
              <strong>${arr.length}</strong>
            </div>

            <div class="kpi">
              <span>Pendientes</span>
              <strong>${
                arr.filter(
                  x =>
                    normEstado(
                      x.estado
                    ) ===
                    "pendiente"
                ).length
              }</strong>
            </div>

            <div class="kpi">
              <span>Aprobadas</span>
              <strong>${
                arr.filter(
                  x =>
                    normEstado(
                      x.estado
                    ) ===
                    "aprobada"
                ).length
              }</strong>
            </div>

            <div class="kpi">
              <span>Rechazadas</span>
              <strong>${
                arr.filter(
                  x =>
                    normEstado(
                      x.estado
                    ) ===
                    "rechazada"
                ).length
              }</strong>
            </div>

          </div>

          ${tablaSolicitudes(arr,"trabajador")}
        `;
  }


  async function renderMisHoras() {

    const arr =
      (
        await cargarSolicitudesPropias()
      )
      .filter(
        x =>
          normEstado(
            x.estado
          ) ===
          "aprobada"
      );


    const h50 =
      arr.reduce(
        (
          s,
          x
        ) =>
          s + numero(x.horas50),
        0
      );


    const h75 =
      arr.reduce(
        (
          s,
          x
        ) =>
          s + numero(x.horas75),
        0
      );


    const h100 =
      arr.reduce(
        (
          s,
          x
        ) =>
          s + numero(x.horas100),
        0
      );


    $("contenidoModulo")
      .innerHTML =
        `
          <div class="resumenKpi">

            <div class="kpi">
              <span>50%</span>
              <strong>${vistaHoras(h50)} h</strong>
            </div>

            <div class="kpi">
              <span>75%</span>
              <strong>${vistaHoras(h75)} h</strong>
            </div>

            <div class="kpi">
              <span>100%</span>
              <strong>${vistaHoras(h100)} h</strong>
            </div>

            <div class="kpi">
              <span>Total</span>
              <strong>${vistaHoras(h50 + h75 + h100)} h</strong>
            </div>

          </div>

          ${tablaSolicitudes(arr,"trabajador")}
        `;
  }


  async function renderSupervisor(
    pendientes
  ) {

    const arr =
      await cargarSolicitudesSupervisor();


    const lista =
      arr.filter(
        x =>
          pendientes
            ?
              normEstado(
                x.estado
              ) ===
              "pendiente"
            :
              normEstado(
                x.estado
              ) !==
              "pendiente"
      );


    $("contenidoModulo")
      .innerHTML =
        tablaSolicitudes(
          lista,
          pendientes
            ?
              "supervisor"
            :
              "historial"
        );
  }


  /* =========================================================
     ADMIN - TODAS SOLICITUDES
  ========================================================= */

  async function renderTodasSolicitudes() {

    const arr =
      await cargarTodasSolicitudes();


    const areas = [
      ...new Set(
        arr
          .map(
            x =>
              x.Area
          )
          .filter(Boolean)
      )
    ]
    .sort();


    $("contenidoModulo")
      .innerHTML =
        `
          <div class="bloque">

            <div class="grid3">

              <div>

                <label>
                  Estado
                </label>

                <select id="admEstado">

                  <option value="">
                    Todos
                  </option>

                  <option value="pendiente">
                    Pendiente
                  </option>

                  <option value="aprobada">
                    Aprobada
                  </option>

                  <option value="rechazada">
                    Rechazada
                  </option>

                </select>

              </div>


              <div>

                <label>
                  Área
                </label>

                <select id="admArea">

                  <option value="">
                    Todas
                  </option>

                  ${
                    areas
                      .map(
                        a =>
                          `<option>${esc(a)}</option>`
                      )
                      .join("")
                  }

                </select>

              </div>


              <div>

                <label>
                  Mes
                </label>

                <input
                  id="admMes"
                  type="month"
                >

              </div>

            </div>

          </div>

          <div id="admTabla"></div>
        `;


    const aplicar =
      () => {

        const e =
          $("admEstado").value;

        const a =
          $("admArea").value;

        const m =
          $("admMes").value;


        const f =
          arr.filter(
            x =>
              (
                !e
                ||
                normEstado(
                  x.estado
                ) ===
                e
              )
              &&
              (
                !a
                ||
                x.Area ===
                a
              )
              &&
              (
                !m
                ||
                String(
                  x.fechaInicio
                  ||
                  ""
                )
                .startsWith(
                  m
                )
              )
          );


        $("admTabla")
          .innerHTML =
            tablaSolicitudes(
              f,
              "admin"
            );
      };


    [
      "admEstado",
      "admArea",
      "admMes"
    ]
    .forEach(
      k =>
        $(k)
          .addEventListener(
            "change",
            aplicar
          )
    );


    aplicar();
  }


  /* =========================================================
     ADMIN - RESUMEN
  ========================================================= */

  async function renderResumenAdmin() {

    const arr =
      await cargarTodasSolicitudes();


    const pendientes =
      arr.filter(
        x =>
          normEstado(x.estado) === "pendiente"
      );


    const aprobadas =
      arr.filter(
        x =>
          normEstado(x.estado) === "aprobada"
      );


    const rechazadas =
      arr.filter(
        x =>
          normEstado(x.estado) === "rechazada"
      );


    const horasAprobadas =
      aprobadas.reduce(
        (suma, x) =>
          suma + numero(x.horasTotales),
        0
      );


    const ultimas =
      arr
        .slice(0, 10);


    $("contenidoModulo")
      .innerHTML =
        `
          <div class="resumenKpi">

            <div class="kpi">
              <span>Pendientes</span>
              <strong>${pendientes.length}</strong>
            </div>

            <div class="kpi">
              <span>Aprobadas</span>
              <strong>${aprobadas.length}</strong>
            </div>

            <div class="kpi">
              <span>Rechazadas</span>
              <strong>${rechazadas.length}</strong>
            </div>

            <div class="kpi">
              <span>Horas aprobadas</span>
              <strong>${vistaHoras(horasAprobadas)} h</strong>
            </div>

          </div>


          <div class="bloque">

            <h3>Últimas solicitudes</h3>

            ${tablaSolicitudes(ultimas, "admin")}

          </div>
        `;
  }


  /* =========================================================
     ADMIN - HORAS A PAGAR
  ========================================================= */

  function numeroPersonalDeSolicitud(s) {

    const directo =
      normalizarNumeroPersonal(
        s?.NumeroPersonal
        ||
        ""
      );


    if (directo) {
      return directo;
    }


    const usuario =
      s?.trabajadorUid
        ?
          cacheUsuarios.get(
            s.trabajadorUid
          )
        :
          null;


    return normalizarNumeroPersonal(
      usuario?.NumeroPersonal
      ||
      ""
    );
  }


  function filasAprobadasPorFechaRecargo(
    solicitudes,
    rango
  ) {

    const mapa =
      new Map();


    solicitudes.forEach(
      s => {

        if (
          normEstado(s.estado) !==
          "aprobada"
        ) {
          return;
        }


        const numeroPersonal =
          numeroPersonalDeSolicitud(s);


        const idInterno =
          numeroPersonal
          ||
          s.trabajadorUid
          ||
          normalizarRut(s.Rut)
          ||
          normalizarTexto(nombreCompleto(s));


        (s.tramosCalculados || [])
          .forEach(
            t => {

              const fecha =
                String(t.fecha || "");


              if (
                fecha < rango.ini
                ||
                fecha >= rango.fin
              ) {
                return;
              }


              const recargo =
                recargoNumero(t.recargo);


              if (
                ![50, 75, 100].includes(recargo)
              ) {
                return;
              }


              const clave =
                `${idInterno}|${fecha}|${recargo}`;


              if (!mapa.has(clave)) {

                mapa.set(
                  clave,
                  {
                    clave,
                    idInterno,
                    trabajadorUid: s.trabajadorUid || "",
                    numeroPersonal,
                    nombre: nombreCompleto(s) || "Trabajador",
                    area: s.Area || "",
                    rut: s.Rut || "",
                    fecha,
                    recargo,
                    horasRegistradas: 0,
                    bonificacion: 0,
                    horasPagar: 0,
                    solicitudes: new Set()
                  }
                );
              }


              const fila =
                mapa.get(clave);


              fila.horasRegistradas +=
                numero(t.horas);


              if (s.id) {
                fila.solicitudes.add(s.id);
              }
            }
          );
      }
    );


    return [...mapa.values()]
      .map(
        fila => {

          const pago =
            detallePagoDesdeHoras(
              fila.horasRegistradas
            );


          return {
            ...fila,
            horasRegistradas: pago.registradas,
            bonificacion: pago.bonificacion,
            horasPagar: pago.pagar,
            solicitudes: [...fila.solicitudes]
          };
        }
      )
      .sort(
        (a, b) =>
          a.nombre.localeCompare(b.nombre, "es")
          ||
          a.fecha.localeCompare(b.fecha)
          ||
          a.recargo - b.recargo
      );
  }


  function resumenTrabajadoresPago(filas) {

    const mapa =
      new Map();


    filas.forEach(
      f => {

        const clave = f.idInterno;


        if (!mapa.has(clave)) {

          mapa.set(
            clave,
            {
              idInterno: clave,
              numeroPersonal: f.numeroPersonal || "",
              nombre: f.nombre,
              area: f.area,
              horasRegistradas: 0,
              bonificacion: 0,
              h50: 0,
              h75: 0,
              h100: 0,
              totalPagar: 0
            }
          );
        }


        const x = mapa.get(clave);

        x.horasRegistradas += f.horasRegistradas;
        x.bonificacion += f.bonificacion;
        x.totalPagar += f.horasPagar;


        if (f.recargo === 50) x.h50 += f.horasPagar;
        if (f.recargo === 75) x.h75 += f.horasPagar;
        if (f.recargo === 100) x.h100 += f.horasPagar;
      }
    );


    return [...mapa.values()]
      .map(
        x => ({
          ...x,
          horasRegistradas: red2(x.horasRegistradas),
          bonificacion: red2(x.bonificacion),
          h50: red2(x.h50),
          h75: red2(x.h75),
          h100: red2(x.h100),
          totalPagar: red2(x.totalPagar)
        })
      )
      .sort(
        (a, b) =>
          a.nombre.localeCompare(b.nombre, "es")
      );
  }


  async function renderHorasPagar() {

    await cargarUsuarios();


    const hoy = new Date();


    $("contenidoModulo").innerHTML =
      `
        <div class="bloque">

          <h3>Horas a pagar</h3>

          <div class="mensaje mensajeInfo">
            Solo se consideran solicitudes <b>aprobadas por el supervisor</b>.
            El cálculo se separa automáticamente por trabajador, fecha y recargo.
          </div>

          <div class="notaReglaPago">
            <b>Regla de horas mínimas a pagar:</b> ${esc(textoEstadoReglaPago())}
            <br>
            <b>Ejemplo:</b> 1,00 h real al 50% + 2,00 h de bonificación al 50% = 3,00 h a pagar al 50%.
          </div>

          <div class="filtrosRango">

            <div>
              <label>Período desde</label>
              <div class="periodoGrupo">
                <select id="pagDesdeMes">${opcionesMeses(hoy.getMonth() + 1)}</select>
                <select id="pagDesdeAnio">${opcionesAnios(hoy.getFullYear())}</select>
              </div>
            </div>

            <div>
              <label>Período hasta</label>
              <div class="periodoGrupo">
                <select id="pagHastaMes">${opcionesMeses(hoy.getMonth() + 1)}</select>
                <select id="pagHastaAnio">${opcionesAnios(hoy.getFullYear())}</select>
              </div>
            </div>

            <div>
              <label>Buscar trabajador</label>
              <input id="pagBuscar" placeholder="Nombre o N° personal">
            </div>

            <div>
              <label>Área</label>
              <input id="pagArea" placeholder="Todas">
            </div>

          </div>

          <div class="acciones">
            <button class="btn btnPrimario" id="btnPagCargar">Cargar</button>
            <button class="btn btnSecundario" id="btnPagExcel" disabled>Descargar Excel</button>
          </div>

        </div>

        <div id="pagResultado"></div>
      `;


    let filasBase = [];
    let rangoActual = null;


    const filtradas =
      () => {

        const buscar =
          normalizarTexto(
            $("pagBuscar").value
          );


        const area =
          normalizarTexto(
            $("pagArea").value
          );


        return filasBase.filter(
          f =>
            (
              !buscar
              ||
              normalizarTexto(f.nombre).includes(buscar)
              ||
              normalizarTexto(f.numeroPersonal).includes(buscar)
            )
            &&
            (
              !area
              ||
              normalizarTexto(f.area) === area
            )
        );
      };


    function pintar() {

      const filas = filtradas();
      const resumen = resumenTrabajadoresPago(filas);


      const total50 = resumen.reduce((a, x) => a + x.h50, 0);
      const total75 = resumen.reduce((a, x) => a + x.h75, 0);
      const total100 = resumen.reduce((a, x) => a + x.h100, 0);
      const totalBonificacion = resumen.reduce((a, x) => a + x.bonificacion, 0);
      const total = resumen.reduce((a, x) => a + x.totalPagar, 0);


      $("btnPagExcel").disabled =
        !filas.length;


      $("pagResultado").innerHTML =
        `
          ${
            rangoActual
              ? `<div class="mensaje mensajeInfo"><b>Período:</b> ${esc(etiquetaRangoPeriodo(rangoActual))}</div>`
              : ""
          }

          <div class="resumenKpi">
            <div class="kpi"><span>50% a pagar</span><strong>${vistaHoras(total50)} h</strong></div>
            <div class="kpi"><span>75% a pagar</span><strong>${vistaHoras(total75)} h</strong></div>
            <div class="kpi"><span>100% a pagar</span><strong>${vistaHoras(total100)} h</strong></div>
            <div class="kpi"><span>Bonificación mínima</span><strong>${vistaHoras(totalBonificacion)} h</strong></div>
            <div class="kpi"><span>Total a pagar</span><strong>${vistaHoras(total)} h</strong></div>
          </div>

          <div class="bloque">
            <h3>Resumen por trabajador</h3>

            ${
              resumen.length
                ?
                  `
                    <div class="tablaWrap">
                      <table class="tabla tablaResumenTrabajador">
                        <thead>
                          <tr>
                            <th>N° personal</th>
                            <th>Trabajador</th>
                            <th>Área</th>
                            <th>Registradas</th>
                            <th>Bonificación</th>
                            <th>50%</th>
                            <th>75%</th>
                            <th>100%</th>
                            <th>Total a pagar</th>
                          </tr>
                        </thead>
                        <tbody>
                          ${
                            resumen.map(
                              x =>
                                `
                                  <tr>
                                    <td>${esc(x.numeroPersonal || "Sin N°")}</td>
                                    <td class="nombrePrincipal">${esc(x.nombre)}</td>
                                    <td>${esc(x.area || "-")}</td>
                                    <td>${vistaHoras(x.horasRegistradas)} h</td>
                                    <td>${vistaHoras(x.bonificacion)} h</td>
                                    <td>${vistaHoras(x.h50)} h</td>
                                    <td>${vistaHoras(x.h75)} h</td>
                                    <td>${vistaHoras(x.h100)} h</td>
                                    <td><b>${vistaHoras(x.totalPagar)} h</b></td>
                                  </tr>
                                `
                            ).join("")
                          }
                        </tbody>
                      </table>
                    </div>
                  `
                :
                  `<div class="mensaje mensajeInfo">No hay horas aprobadas para mostrar.</div>`
            }
          </div>

          <div class="bloque">
            <h3>Detalle por fecha y recargo</h3>

            ${
              filas.length
                ?
                  `
                    <div class="tablaWrap">
                      <table class="tabla">
                        <thead>
                          <tr>
                            <th>N° personal</th>
                            <th>Trabajador</th>
                            <th>Fecha</th>
                            <th>Recargo</th>
                            <th>Horas registradas</th>
                            <th>Bonificación por mínimo</th>
                            <th>Horas a pagar</th>
                          </tr>
                        </thead>
                        <tbody>
                          ${
                            filas.map(
                              f =>
                                `
                                  <tr>
                                    <td>${esc(f.numeroPersonal || "Sin N°")}</td>
                                    <td>${esc(f.nombre)}</td>
                                    <td>${fechaCL(f.fecha)}</td>
                                    <td>${f.recargo}%</td>
                                    <td>${vistaHoras(f.horasRegistradas)} h</td>
                                    <td>
                                      ${vistaHoras(f.bonificacion)} h
                                      ${
                                        f.bonificacion > 0
                                          ? `<span class="bonificacionPago">Bonificación al ${f.recargo}%</span>`
                                          : `<span class="detallePagoLinea">No aplica mínimo adicional</span>`
                                      }
                                    </td>
                                    <td>
                                      <b>${vistaHoras(f.horasPagar)} h</b>
                                      ${
                                        f.bonificacion > 0
                                          ? `<span class="detallePagoLinea">${vistaHoras(f.horasRegistradas)} h reales + ${vistaHoras(f.bonificacion)} h bonificación</span>`
                                          : ""
                                      }
                                    </td>
                                  </tr>
                                `
                            ).join("")
                          }
                        </tbody>
                      </table>
                    </div>
                  `
                :
                  `<div class="mensaje mensajeInfo">Sin detalle.</div>`
            }
          </div>
        `;
    }


    $("btnPagCargar")
      .addEventListener(
        "click",
        async () => {

          try {

            rangoActual =
              rangoPeriodos(
                Number($("pagDesdeAnio").value),
                Number($("pagDesdeMes").value),
                Number($("pagHastaAnio").value),
                Number($("pagHastaMes").value)
              );


            const solicitudes =
              await solicitudesRango(
                rangoActual.yDesde,
                rangoActual.mDesde,
                rangoActual.yHasta,
                rangoActual.mHasta,
                true
              );


            filasBase =
              filasAprobadasPorFechaRecargo(
                solicitudes,
                rangoActual
              );


            pintar();

          }
          catch (err) {

            console.error(err);
            toast(err.message || "No fue posible calcular las horas a pagar.", "error");
          }
        }
      );


    $("pagBuscar").addEventListener("input", pintar);
    $("pagArea").addEventListener("input", pintar);


    $("btnPagExcel")
      .addEventListener(
        "click",
        () => {

          const filas = filtradas();


          if (!filas.length || !rangoActual) return;


          const resumen =
            resumenTrabajadoresPago(filas)
              .map(
                x => ({
                  "N° de personal": x.numeroPersonal,
                  "Trabajador": x.nombre,
                  "Área": x.area,
                  "Horas registradas": x.horasRegistradas,
                  "Bonificación por mínimo": x.bonificacion,
                  "Horas 50% a pagar": x.h50,
                  "Horas 75% a pagar": x.h75,
                  "Horas 100% a pagar": x.h100,
                  "Total horas a pagar": x.totalPagar
                })
              );


          const detalle =
            filas.map(
              f => ({
                "N° de personal": f.numeroPersonal,
                "Trabajador": f.nombre,
                "Área": f.area,
                "Fecha": fechaCL(f.fecha),
                "Recargo": `${f.recargo}%`,
                "Horas registradas": f.horasRegistradas,
                "Bonificación por mínimo": f.bonificacion,
                "Horas a pagar": f.horasPagar,
                "Explicación pago": f.bonificacion > 0
                  ? `${vistaHoras(f.horasRegistradas)} h reales + ${vistaHoras(f.bonificacion)} h bonificación al ${f.recargo}% = ${vistaHoras(f.horasPagar)} h a pagar`
                  : `${vistaHoras(f.horasRegistradas)} h reales = ${vistaHoras(f.horasPagar)} h a pagar`
              })
            );


          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), "Resumen pago");
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detalle), "Detalle pago");


          XLSX.writeFile(
            wb,
            `Horas_a_pagar_${rangoActual.yDesde}-${String(rangoActual.mDesde).padStart(2,"0")}_a_${rangoActual.yHasta}-${String(rangoActual.mHasta).padStart(2,"0")}.xlsx`,
            { compression: true }
          );
        }
      );
  }


  /* =========================================================
     USUARIOS / PREAUTORIZACIONES
  ========================================================= */

  async function cargarUsuarios() {

    const s =
      await getDocs(
        collection(
          db,
          "usuarios"
        )
      );


    const arr =
      s.docs.map(
        d => ({
          id:
            d.id,
          ...d.data()
        })
      );


    arr.sort(
      (
        a,
        b
      ) =>
        nombreCompleto(a)
          .localeCompare(
            nombreCompleto(b),
            "es"
          )
    );


    cacheUsuarios.clear();


    arr.forEach(
      x =>
        cacheUsuarios.set(
          x.id,
          x
        )
    );


    return arr;
  }


  async function cargarAutorizaciones() {

    const s =
      await getDocs(
        collection(
          db,
          "preautorizados"
        )
      );


    const arr =
      s.docs.map(
        d => ({
          id:
            d.id,
          ...d.data()
        })
      );


    arr.sort(
      (
        a,
        b
      ) =>
        nombreCompleto(a)
          .localeCompare(
            nombreCompleto(b),
            "es"
          )
    );


    cacheAutorizaciones.clear();


    arr.forEach(
      x =>
        cacheAutorizaciones.set(
          x.id,
          x
        )
    );


    return arr;
  }


  async function numeroPersonalDisponible(
    numeroPersonal,
    {
      excluirUsuarioId = "",
      excluirAutorizacionId = ""
    } = {}
  ) {

    const numeroBuscado =
      claveNumeroPersonal(
        numeroPersonal
      );


    if (!numeroBuscado) {
      return true;
    }


    const [usuariosSnap, autorizacionesSnap] =
      await Promise.all([
        getDocs(collection(db, "usuarios")),
        getDocs(collection(db, "preautorizados"))
      ]);


    const repetidoUsuario =
      usuariosSnap.docs.some(
        d =>
          d.id !== excluirUsuarioId
          &&
          claveNumeroPersonal(
            d.data()?.NumeroPersonal
          ) === numeroBuscado
      );


    const repetidoAutorizacion =
      autorizacionesSnap.docs.some(
        d =>
          d.id !== excluirAutorizacionId
          &&
          claveNumeroPersonal(
            d.data()?.NumeroPersonal
          ) === numeroBuscado
      );


    return !(
      repetidoUsuario
      ||
      repetidoAutorizacion
    );
  }


  async function renderUsuarios() {

    const [
      usuarios,
      autorizaciones
    ] =
      await Promise.all([
        cargarUsuarios(),
        cargarAutorizaciones()
      ]);


    const pendientes =
      autorizaciones.filter(
        a =>
          a.activado !==
          true
      );


    const activadas =
      autorizaciones.filter(
        a =>
          a.activado ===
          true
      );


    $("contenidoModulo")
      .innerHTML =
        `
          <div class="mensaje mensajeInfo">

            <b>Nuevo flujo:</b>

            Administración autoriza a la persona aquí.

            El trabajador entra a la pantalla de inicio y pulsa

            <b>
              ¿Primera vez? Activar mi cuenta
            </b>.

            Firebase genera el UID automáticamente.

          </div>


          <div
            class="acciones"
            style="justify-content:flex-start"
          >

            <button
              class="btn btnPrimario"
              id="nuevaAutorizacion"
            >
              Autorizar usuario
            </button>

          </div>


          <div
            class="bloque"
            style="margin-top:16px"
          >

            <h3>
              Autorizaciones pendientes
            </h3>

            ${
              pendientes.length
                ?
                  `
                    <div class="tablaWrap">

                      <table class="tabla">

                        <thead>

                          <tr>
                            <th>N° personal</th>
                            <th>Nombre</th>
                            <th>Correo</th>
                            <th>Área</th>
                            <th>Solicita HHEE</th>
                            <th>Supervisor</th>
                            <th>Administrador</th>
                            <th>Acciones</th>
                          </tr>

                        </thead>

                        <tbody>

                          ${
                            pendientes
                              .map(
                                a =>
                                  `
                                    <tr>

                                      <td>
                                        ${esc(a.NumeroPersonal || "-")}
                                      </td>

                                      <td>
                                        ${esc(nombreCompleto(a))}
                                        <br>
                                        <small>
                                          ${esc(a.Rut || "")}
                                        </small>
                                      </td>

                                      <td>
                                        ${esc(a.Correo || a.id)}
                                      </td>

                                      <td>
                                        ${esc(a.Area || "")}
                                      </td>

                                      <td>
                                        ${
                                          a.puedeSolicitarHE ===
                                          true
                                            ?
                                              "Sí"
                                            :
                                              "No"
                                        }
                                      </td>

                                      <td>
                                        ${
                                          a.esSupervisor ===
                                          true
                                            ?
                                              "Sí"
                                            :
                                              "No"
                                        }
                                      </td>

                                      <td>
                                        ${
                                          a.esAdministrador ===
                                          true
                                            ?
                                              "Sí"
                                            :
                                              "No"
                                        }
                                      </td>

                                      <td>

                                        <div class="miniAcciones">

                                          <button
                                            class="btn btnSecundario"
                                            onclick="window.editarAutorizacion('${encodeURIComponent(a.id)}')"
                                          >
                                            Editar
                                          </button>

                                          <button
                                            class="btn btnPeligro"
                                            onclick="window.eliminarAutorizacion('${encodeURIComponent(a.id)}')"
                                          >
                                            Eliminar
                                          </button>

                                        </div>

                                      </td>

                                    </tr>
                                  `
                              )
                              .join("")
                          }

                        </tbody>

                      </table>

                    </div>
                  `
                :
                  `
                    <div class="mensaje mensajeInfo">
                      No hay autorizaciones pendientes.
                    </div>
                  `
            }

          </div>


          <div class="bloque">

            <h3>
              Usuarios activos
            </h3>

            ${
              usuarios.length
                ?
                  `
                    <div class="tablaWrap">

                      <table class="tabla">

                        <thead>

                          <tr>
                            <th>N° personal</th>
                            <th>Nombre</th>
                            <th>Correo</th>
                            <th>Área</th>
                            <th>Activo</th>
                            <th>Solicita HHEE</th>
                            <th>Supervisor</th>
                            <th>Administrador</th>
                            <th>Acciones</th>
                          </tr>

                        </thead>

                        <tbody>

                          ${
                            usuarios
                              .map(
                                u =>
                                  `
                                    <tr>

                                      <td>
                                        ${esc(u.NumeroPersonal || "-")}
                                      </td>

                                      <td>
                                        ${esc(nombreCompleto(u))}
                                        <br>
                                        <small>
                                          ${esc(u.Rut || "")}
                                        </small>
                                      </td>

                                      <td>
                                        ${esc(u.Correo || "")}
                                      </td>

                                      <td>
                                        ${esc(u.Area || "")}
                                      </td>

                                      <td>
                                        ${
                                          u.Activo ===
                                          true
                                            ?
                                              "Sí"
                                            :
                                              "No"
                                        }
                                      </td>

                                      <td>
                                        ${
                                          u.puedeSolicitarHE ===
                                          true
                                            ?
                                              "Sí"
                                            :
                                              "No"
                                        }
                                      </td>

                                      <td>
                                        ${
                                          u.esSupervisor ===
                                          true
                                            ?
                                              "Sí"
                                            :
                                              "No"
                                        }
                                      </td>

                                      <td>
                                        ${
                                          u.esAdministrador ===
                                          true
                                            ?
                                              "Sí"
                                            :
                                              "No"
                                        }
                                      </td>

                                      <td>

                                        <button
                                          class="btn btnSecundario"
                                          onclick="window.editarPerfilActivo('${u.id}')"
                                        >
                                          Editar
                                        </button>

                                      </td>

                                    </tr>
                                  `
                              )
                              .join("")
                          }

                        </tbody>

                      </table>

                    </div>
                  `
                :
                  `
                    <div class="mensaje mensajeInfo">
                      Todavía no hay usuarios activos.
                    </div>
                  `
            }

          </div>


          ${
            activadas.length
              ?
                `
                  <div class="mensaje mensajeInfo">

                    Autorizaciones ya utilizadas:

                    <b>
                      ${activadas.length}
                    </b>.

                    Sus permisos actuales se administran desde la tabla

                    <b>
                      Usuarios activos
                    </b>.

                  </div>
                `
              :
                ""
          }
        `;


    $("nuevaAutorizacion")
      .addEventListener(
        "click",
        () =>
          formAutorizacion(
            null
          )
      );
  }


  function formAutorizacion(a) {

    const nuevo =
      !a;


    modal(
      `
        <h3>
          ${
            nuevo
              ?
                "Autorizar usuario"
              :
                "Editar autorización"
          }
        </h3>

        <form id="formAutorizacion">

          <div class="grid2">

            <div>

              <label>
                Nombre *
              </label>

              <input
                id="aNombre"
                value="${esc(a?.Nombre || "")}"
                required
              >

            </div>


            <div>

              <label>
                Apellido *
              </label>

              <input
                id="aApellido"
                value="${esc(a?.Apellido || "")}"
                required
              >

            </div>


            <div>

              <label>
                RUT
              </label>

              <input
                id="aRut"
                value="${esc(a?.Rut || "")}"
              >

            </div>


            <div>

              <label>
                N° de personal
              </label>

              <input
                id="aNumeroPersonal"
                inputmode="numeric"
                value="${esc(a?.NumeroPersonal || "")}"
                placeholder="Ej.: 009981"
              >

              <div class="nota">
                Identificador utilizado por SAP. Para quienes solicitan HHEE es obligatorio y no puede repetirse.
              </div>

            </div>


            <div>

              <label>
                Área
              </label>

              <input
                id="aArea"
                value="${esc(a?.Area || "")}"
              >

            </div>


            <div>

              <label>
                Correo *
              </label>

              <input
                id="aCorreo"
                type="email"
                value="${esc(a?.Correo || a?.id || "")}"
                ${
                  nuevo
                    ?
                      ""
                    :
                      "readonly"
                }
                required
              >

              <div class="nota">
                Este correo será la llave de activación.
                No se necesita UID.
              </div>

            </div>


            <div>

              <label>
                Correo supervisor
              </label>

              <input
                id="aSupervisorCorreo"
                type="email"
                value="${esc(a?.supervisorcorreo || "")}"
              >

            </div>

          </div>


          <div
            class="grid4"
            style="margin-top:14px"
          >

            <div class="checkRow">

              <input
                id="aActivo"
                type="checkbox"
                ${
                  a?.Activo !==
                  false
                    ?
                      "checked"
                    :
                      ""
                }
              >

              <label>
                Activo
              </label>

            </div>


            <div class="checkRow">

              <input
                id="aSolicita"
                type="checkbox"
                ${
                  a?.puedeSolicitarHE
                    ?
                      "checked"
                    :
                      ""
                }
              >

              <label>
                Solicita HHEE
              </label>

            </div>


            <div class="checkRow">

              <input
                id="aSupervisor"
                type="checkbox"
                ${
                  a?.esSupervisor
                    ?
                      "checked"
                    :
                      ""
                }
              >

              <label>
                Supervisor
              </label>

            </div>


            <div class="checkRow">

              <input
                id="aAdmin"
                type="checkbox"
                ${
                  a?.esAdministrador
                    ?
                      "checked"
                    :
                      ""
                }
              >

              <label>
                Administrador
              </label>

            </div>

          </div>


          ${
            a?.activado ===
            true
              ?
                `
                  <div
                    class="mensaje mensajeAviso"
                    style="margin-top:14px"
                  >
                    Esta autorización ya fue activada.
                    Los cambios también se copiarán al perfil activo.
                  </div>
                `
              :
                ""
          }


          <div class="acciones">

            <button
              type="button"
              class="btn btnSecundario"
              onclick="window.cerrarModal()"
            >
              Cancelar
            </button>

            <button
              class="btn btnPrimario"
              type="submit"
            >
              Guardar
            </button>

          </div>

        </form>
      `
    );


    $("formAutorizacion")
      .addEventListener(
        "submit",
        async e => {

          e.preventDefault();


          const correo =
            normalizarCorreo(
              $("aCorreo").value
            );


          if (!correo) {

            return toast(
              "Ingresa un correo válido.",
              "error"
            );
          }


          const numeroPersonal =
            normalizarNumeroPersonal(
              $("aNumeroPersonal").value
            );


          if (
            $("aSolicita").checked
            &&
            !numeroPersonal
          ) {

            return toast(
              "El N° de personal es obligatorio para quienes solicitan HHEE.",
              "error"
            );
          }


          const datos = {

            Nombre:
              $("aNombre")
                .value
                .trim(),

            Apellido:
              $("aApellido")
                .value
                .trim(),

            Rut:
              $("aRut")
                .value
                .trim(),

            NumeroPersonal:
              numeroPersonal,

            Area:
              $("aArea")
                .value
                .trim(),

            Correo:
              correo,

            supervisorcorreo:
              normalizarCorreo(
                $("aSupervisorCorreo").value
              ),

            Activo:
              $("aActivo").checked,

            puedeSolicitarHE:
              $("aSolicita").checked,

            esSupervisor:
              $("aSupervisor").checked,

            esAdministrador:
              $("aAdmin").checked
          };


          try {

            if (numeroPersonal) {

              const disponible =
                await numeroPersonalDisponible(
                  numeroPersonal,
                  {
                    excluirUsuarioId:
                      a?.uid
                      ||
                      "",

                    excluirAutorizacionId:
                      a?.id
                      ||
                      ""
                  }
                );


              if (!disponible) {

                return toast(
                  `El N° de personal ${numeroPersonal} ya está asignado a otro usuario.`,
                  "error"
                );
              }
            }


            if (nuevo) {

              const existente =
                await getDoc(
                  doc(
                    db,
                    "preautorizados",
                    correo
                  )
                );


              if (
                existente.exists()
              ) {

                return toast(
                  "Ese correo ya tiene una autorización.",
                  "error"
                );
              }


              await setDoc(
                doc(
                  db,
                  "preautorizados",
                  correo
                ),
                {

                  ...datos,

                  activado:
                    false,

                  uid:
                    "",

                  fechaAutorizacion:
                    serverTimestamp(),

                  fechaActualizacion:
                    serverTimestamp(),

                  autorizadoPor:
                    correoActual()
                }
              );

            }
            else {

              const original =
                cacheAutorizaciones.get(
                  a.id
                )
                ||
                a;


              await setDoc(
                doc(
                  db,
                  "preautorizados",
                  a.id
                ),
                {

                  ...datos,

                  fechaActualizacion:
                    serverTimestamp()
                },
                {
                  merge:
                    true
                }
              );


              if (
                original.activado ===
                true
                &&
                original.uid
              ) {

                if (
                  original.uid ===
                  usuarioActual.uid
                  &&
                  !datos.esAdministrador
                ) {

                  return toast(
                    "No puedes quitarte a ti misma el permiso de administrador.",
                    "error"
                  );
                }


                await setDoc(
                  doc(
                    db,
                    "usuarios",
                    original.uid
                  ),
                  {

                    ...datos,

                    fechaActualizacion:
                      serverTimestamp()
                  },
                  {
                    merge:
                      true
                  }
                );
              }
            }


            cerrarModal();


            toast(
              nuevo
                ?
                  "Usuario autorizado. Ya puede activar su cuenta."
                :
                  "Autorización actualizada.",
              "exito"
            );


            abrirModulo(
              "usuarios"
            );

          }
          catch (err) {

            console.error(err);


            toast(
              err.code?.includes(
                "permission"
              )
                ?
                  "Firestore rechazó la operación por permisos."
                :
                  "No fue posible guardar la autorización.",
              "error"
            );
          }
        }
      );
  }


  window.editarAutorizacion =
    encoded => {

      const id =
        decodeURIComponent(
          encoded
        );


      const a =
        cacheAutorizaciones.get(
          id
        );


      if (a) {

        formAutorizacion(
          a
        );
      }
    };


  window.eliminarAutorizacion =
    async encoded => {

      const id =
        decodeURIComponent(
          encoded
        );


      const a =
        cacheAutorizaciones.get(
          id
        );


      if (!a) {
        return;
      }


      if (
        a.activado ===
        true
      ) {

        return toast(
          "Una autorización ya activada no se elimina desde aquí. Desactiva el perfil activo.",
          "error"
        );
      }


      if (
        !confirm(
          `¿Eliminar la autorización de ${a.Correo || id}?`
        )
      ) {

        return;
      }


      try {

        await deleteDoc(
          doc(
            db,
            "preautorizados",
            id
          )
        );


        toast(
          "Autorización eliminada.",
          "exito"
        );


        abrirModulo(
          "usuarios"
        );

      }
      catch (err) {

        console.error(err);


        toast(
          "No fue posible eliminar la autorización.",
          "error"
        );
      }
    };


  function formPerfilActivo(u) {

    modal(
      `
        <h3>
          Editar usuario activo
        </h3>

        <form id="formPerfilActivo">

          <div class="grid2">

            <div>
              <label>Nombre *</label>
              <input id="pNombre" value="${esc(u?.Nombre || "")}" required>
            </div>

            <div>
              <label>Apellido *</label>
              <input id="pApellido" value="${esc(u?.Apellido || "")}" required>
            </div>

            <div>
              <label>RUT</label>
              <input id="pRut" value="${esc(u?.Rut || "")}">
            </div>

            <div>
              <label>N° de personal</label>
              <input id="pNumeroPersonal" inputmode="numeric" value="${esc(u?.NumeroPersonal || "")}" placeholder="Ej.: 009981">
              <div class="nota">Identificador SAP. Obligatorio si el usuario solicita HHEE.</div>
            </div>

            <div>
              <label>Área</label>
              <input id="pArea" value="${esc(u?.Area || "")}">
            </div>

            <div>
              <label>Correo</label>
              <input id="pCorreo" type="email" value="${esc(u?.Correo || "")}" readonly>
            </div>

            <div>
              <label>Correo supervisor</label>
              <input id="pSupervisorCorreo" type="email" value="${esc(u?.supervisorcorreo || "")}">
            </div>

          </div>


          <div class="grid4" style="margin-top:14px">

            <div class="checkRow">
              <input id="pActivo" type="checkbox" ${u?.Activo !== false ? "checked" : ""}>
              <label>Activo</label>
            </div>

            <div class="checkRow">
              <input id="pSolicita" type="checkbox" ${u?.puedeSolicitarHE ? "checked" : ""}>
              <label>Solicita HHEE</label>
            </div>

            <div class="checkRow">
              <input id="pSupervisor" type="checkbox" ${u?.esSupervisor ? "checked" : ""}>
              <label>Supervisor</label>
            </div>

            <div class="checkRow">
              <input id="pAdmin" type="checkbox" ${u?.esAdministrador ? "checked" : ""}>
              <label>Administrador</label>
            </div>

          </div>


          <div class="acciones">

            <button
              type="button"
              class="btn btnSecundario"
              onclick="window.cerrarModal()"
            >
              Cancelar
            </button>

            <button
              class="btn btnPrimario"
              type="submit"
            >
              Guardar
            </button>

          </div>

        </form>
      `
    );


    $("formPerfilActivo")
      .addEventListener(
        "submit",
        async e => {

          e.preventDefault();


          if (
            u.id ===
            usuarioActual.uid
            &&
            !$("pAdmin").checked
          ) {

            return toast(
              "No puedes quitarte a ti misma el permiso de administrador.",
              "error"
            );
          }


          const numeroPersonal =
            normalizarNumeroPersonal(
              $("pNumeroPersonal").value
            );


          if (
            $("pSolicita").checked
            &&
            !numeroPersonal
          ) {

            return toast(
              "El N° de personal es obligatorio para quienes solicitan HHEE.",
              "error"
            );
          }


          const disponible =
            await numeroPersonalDisponible(
              numeroPersonal,
              {
                excluirUsuarioId:
                  u.id,

                excluirAutorizacionId:
                  normalizarCorreo(
                    u.Correo
                    ||
                    ""
                  )
              }
            );


          if (
            numeroPersonal
            &&
            !disponible
          ) {

            return toast(
              `El N° de personal ${numeroPersonal} ya está asignado a otro usuario.`,
              "error"
            );
          }


          const datos = {

            Nombre:
              $("pNombre")
                .value
                .trim(),

            Apellido:
              $("pApellido")
                .value
                .trim(),

            Rut:
              $("pRut")
                .value
                .trim(),

            NumeroPersonal:
              numeroPersonal,

            Area:
              $("pArea")
                .value
                .trim(),

            Correo:
              normalizarCorreo(
                $("pCorreo").value
              ),

            supervisorcorreo:
              normalizarCorreo(
                $("pSupervisorCorreo").value
              ),

            Activo:
              $("pActivo").checked,

            puedeSolicitarHE:
              $("pSolicita").checked,

            esSupervisor:
              $("pSupervisor").checked,

            esAdministrador:
              $("pAdmin").checked,

            fechaActualizacion:
              serverTimestamp()
          };


          try {

            await setDoc(
              doc(
                db,
                "usuarios",
                u.id
              ),
              datos,
              {
                merge:
                  true
              }
            );


            const preId =
              normalizarCorreo(
                u.Correo
                ||
                ""
              );


            if (preId) {

              const pre =
                await getDoc(
                  doc(
                    db,
                    "preautorizados",
                    preId
                  )
                );


              if (
                pre.exists()
              ) {

                await setDoc(
                  doc(
                    db,
                    "preautorizados",
                    preId
                  ),
                  {

                    ...datos,

                    fechaActualizacion:
                      serverTimestamp()
                  },
                  {
                    merge:
                      true
                  }
                );
              }
            }


            cerrarModal();


            toast(
              "Perfil actualizado.",
              "exito"
            );


            abrirModulo(
              "usuarios"
            );

          }
          catch (err) {

            console.error(err);


            toast(
              "No fue posible actualizar el perfil.",
              "error"
            );
          }
        }
      );
  }


  window.editarPerfilActivo =
    id => {

      const u =
        cacheUsuarios.get(
          id
        );


      if (u) {

        formPerfilActivo(
          u
        );
      }
    };


  /* =========================================================
     EXPORTACIÓN
  ========================================================= */

  function selectPeriodo(prefix) {

    const hoy =
      new Date();


    return `
      <div>

        <label>
          Año
        </label>

        <select id="${prefix}Anio">

          ${
            Array.from(
              {
                length:
                  8
              },
              (
                _,
                i
              ) =>
                hoy.getFullYear()
                -
                5
                +
                i
            )
            .map(
              y =>
                `
                  <option
                    value="${y}"
                    ${
                      y ===
                      hoy.getFullYear()
                        ?
                          "selected"
                        :
                          ""
                    }
                  >
                    ${y}
                  </option>
                `
            )
            .join("")
          }

        </select>

      </div>


      <div>

        <label>
          Mes
        </label>

        <select id="${prefix}Mes">

          ${
            MESES
              .slice(1)
              .map(
                (
                  m,
                  i
                ) =>
                  `
                    <option
                      value="${i + 1}"
                      ${
                        i + 1 ===
                        hoy.getMonth() + 1
                          ?
                            "selected"
                          :
                            ""
                      }
                    >
                      ${m}
                    </option>
                  `
              )
              .join("")
          }

        </select>

      </div>
    `;
  }


  function rangoMes(
    y,
    m
  ) {

    const ini =
      `${y}-${String(m).padStart(2,"0")}-01`;


    const ny =
      m === 12
        ?
          y + 1
        :
          y;


    const nm =
      m === 12
        ?
          1
        :
          m + 1;


    return {

      ini,

      fin:
        `${ny}-${String(nm).padStart(2,"0")}-01`
    };
  }


  async function solicitudesPeriodo(
    y,
    m,
    buffer = false
  ) {

    const r =
      rangoMes(
        y,
        m
      );


    let ini =
      r.ini;

    let fin =
      r.fin;


    if (buffer) {

      ini =
        sumarDias(
          ini,
          -1
        );


      fin =
        sumarDias(
          fin,
          1
        );
    }


    const q =
      query(
        collection(
          db,
          "solicitudes"
        ),
        where(
          "fechaInicio",
          ">=",
          ini
        ),
        where(
          "fechaInicio",
          "<",
          fin
        )
      );


    const s =
      await getDocs(q);


    const arr =
      s.docs.map(
        d => ({
          id:
            d.id,
          ...d.data()
        })
      );


    arr.forEach(
      x =>
        cacheSolicitudes.set(
          x.id,
          x
        )
    );


    return arr;
  }


  function opcionesAnios(
    seleccionado = new Date().getFullYear()
  ) {

    const hoy =
      new Date();


    return Array.from(
      {
        length: 10
      },
      (_, i) =>
        hoy.getFullYear()
        -
        6
        +
        i
    )
    .map(
      y =>
        `
          <option
            value="${y}"
            ${
              y === Number(seleccionado)
                ?
                  "selected"
                :
                  ""
            }
          >
            ${y}
          </option>
        `
    )
    .join("");
  }


  function opcionesMeses(
    seleccionado = new Date().getMonth() + 1
  ) {

    return MESES
      .slice(1)
      .map(
        (m, i) =>
          `
            <option
              value="${i + 1}"
              ${
                i + 1 === Number(seleccionado)
                  ?
                    "selected"
                  :
                    ""
              }
            >
              ${m}
            </option>
          `
      )
      .join("");
  }


  function rangoPeriodos(
    yDesde,
    mDesde,
    yHasta,
    mHasta
  ) {

    const desdeIndice =
      Number(yDesde) * 12
      +
      Number(mDesde)
      -
      1;


    const hastaIndice =
      Number(yHasta) * 12
      +
      Number(mHasta)
      -
      1;


    if (hastaIndice < desdeIndice) {
      throw new Error(
        "El período hasta no puede ser anterior al período desde."
      );
    }


    const ini =
      `${Number(yDesde)}-${String(Number(mDesde)).padStart(2,"0")}-01`;


    const siguienteMes =
      Number(mHasta) === 12
        ?
          1
        :
          Number(mHasta) + 1;


    const siguienteAnio =
      Number(mHasta) === 12
        ?
          Number(yHasta) + 1
        :
          Number(yHasta);


    const fin =
      `${siguienteAnio}-${String(siguienteMes).padStart(2,"0")}-01`;


    return {
      ini,
      fin,
      yDesde: Number(yDesde),
      mDesde: Number(mDesde),
      yHasta: Number(yHasta),
      mHasta: Number(mHasta)
    };
  }


  function etiquetaRangoPeriodo(r) {

    if (!r) {
      return "";
    }


    if (
      r.yDesde === r.yHasta
      &&
      r.mDesde === r.mHasta
    ) {

      return `${MESES[r.mDesde]} ${r.yDesde}`;
    }


    return `${MESES[r.mDesde]} ${r.yDesde} a ${MESES[r.mHasta]} ${r.yHasta}`;
  }


  async function solicitudesRango(
    yDesde,
    mDesde,
    yHasta,
    mHasta,
    buffer = false
  ) {

    const r =
      rangoPeriodos(
        yDesde,
        mDesde,
        yHasta,
        mHasta
      );


    let ini = r.ini;
    let fin = r.fin;


    if (buffer) {
      ini = sumarDias(ini, -1);
      fin = sumarDias(fin, 1);
    }


    const q =
      query(
        collection(db, "solicitudes"),
        where("fechaInicio", ">=", ini),
        where("fechaInicio", "<", fin)
      );


    const snap =
      await getDocs(q);


    const arr =
      snap.docs.map(
        d => ({
          id: d.id,
          ...d.data()
        })
      );


    arr.forEach(
      x =>
        cacheSolicitudes.set(
          x.id,
          x
        )
    );


    return arr;
  }


  async function renderExportar() {

    await cargarUsuarios();

    const hoy = new Date();


    $("contenidoModulo")
      .innerHTML =
        `
          <div class="bloque">

            <h3>
              Respaldo por rango en Excel
            </h3>

            <div class="mensaje mensajeInfo">
              Puedes descargar varios meses en un solo archivo.
              Por ejemplo: marzo 2026 a julio 2026.
            </div>

            <div class="filtrosRango">

              <div>
                <label>Período desde</label>
                <div class="periodoGrupo">
                  <select id="expDesdeMes">
                    ${opcionesMeses(hoy.getMonth() + 1)}
                  </select>
                  <select id="expDesdeAnio">
                    ${opcionesAnios(hoy.getFullYear())}
                  </select>
                </div>
              </div>

              <div>
                <label>Período hasta</label>
                <div class="periodoGrupo">
                  <select id="expHastaMes">
                    ${opcionesMeses(hoy.getMonth() + 1)}
                  </select>
                  <select id="expHastaAnio">
                    ${opcionesAnios(hoy.getFullYear())}
                  </select>
                </div>
              </div>

              <div>
                <label>Estado</label>
                <select id="expEstado">
                  <option value="">Todos</option>
                  <option value="pendiente">Pendiente</option>
                  <option value="aprobada">Aprobada</option>
                  <option value="rechazada">Rechazada</option>
                </select>
              </div>

              <div>
                <label>Área</label>
                <input id="expArea" placeholder="Todas">
              </div>

            </div>


            <div class="acciones">

              <button
                class="btn btnPrimario"
                id="btnExpCargar"
              >
                Cargar
              </button>

              <button
                class="btn btnSecundario"
                id="btnExpExcel"
                disabled
              >
                Descargar Excel
              </button>

            </div>

          </div>


          <div id="expResultado"></div>
        `;


    let base = [];
    let rangoActual = null;


    const obtenerRangoSeleccionado =
      () =>
        rangoPeriodos(
          Number($("expDesdeAnio").value),
          Number($("expDesdeMes").value),
          Number($("expHastaAnio").value),
          Number($("expHastaMes").value)
        );


    const filtradas =
      () =>
        base.filter(
          x =>
            (
              !$("expEstado").value
              ||
              normEstado(x.estado) === $("expEstado").value
            )
            &&
            (
              !$("expArea").value.trim()
              ||
              normalizarTexto(x.Area) ===
              normalizarTexto($("expArea").value)
            )
        );


    function renderVistaExport() {

      const f = filtradas();


      $("btnExpExcel").disabled =
        !f.length;


      $("expResultado").innerHTML =
        `
          ${
            rangoActual
              ?
                `<div class="mensaje mensajeInfo"><b>Período cargado:</b> ${esc(etiquetaRangoPeriodo(rangoActual))}</div>`
              :
                ""
          }

          <div class="resumenKpi">

            <div class="kpi">
              <span>Total</span>
              <strong>${f.length}</strong>
            </div>

            <div class="kpi">
              <span>Pendientes</span>
              <strong>${f.filter(x => normEstado(x.estado) === "pendiente").length}</strong>
            </div>

            <div class="kpi">
              <span>Aprobadas</span>
              <strong>${f.filter(x => normEstado(x.estado) === "aprobada").length}</strong>
            </div>

            <div class="kpi">
              <span>Rechazadas</span>
              <strong>${f.filter(x => normEstado(x.estado) === "rechazada").length}</strong>
            </div>

          </div>

          ${tablaSolicitudes(f, "admin")}
        `;
    }


    $("btnExpCargar")
      .addEventListener(
        "click",
        async () => {

          try {

            rangoActual =
              obtenerRangoSeleccionado();


            base =
              await solicitudesRango(
                rangoActual.yDesde,
                rangoActual.mDesde,
                rangoActual.yHasta,
                rangoActual.mHasta
              );


            renderVistaExport();

          }
          catch (err) {

            console.error(err);


            toast(
              err.message || "No fue posible cargar el período.",
              "error"
            );
          }
        }
      );


    $("expEstado")
      .addEventListener(
        "change",
        renderVistaExport
      );


    $("expArea")
      .addEventListener(
        "input",
        renderVistaExport
      );


    $("btnExpExcel")
      .addEventListener(
        "click",
        () => {

          if (!rangoActual) {
            return;
          }


          exportarExcelRango(
            filtradas(),
            rangoActual
          );
        }
      );
  }


  function exportarExcelRango(
    lista,
    rango
  ) {

    if (!window.XLSX) {

      return toast(
        "No se cargó SheetJS.",
        "error"
      );
    }


    const sol =
      lista.map(
        (s, i) => ({

          "N°": i + 1,
          "N° de personal": numeroPersonalDeSolicitud(s),
          "Nombre": s.Nombre || "",
          "Apellidos": s.Apellido || "",
          "RUT": s.Rut || "",
          "Área": s.Area || "",
          "Correo trabajador": s.correoTrabajador || "",
          "Supervisor": s.supervisorCorreo || "",
          "Tipo formulario": s.tipoFormulario || "",
          "Día trabajo": s.diaTrabajo || "",
          "Fecha inicio": fechaCL(s.fechaInicio),
          "Fecha término": fechaCL(s.fechaTermino),
          "Hora inicio": s.horaInicio || "",
          "Hora término": s.horaTermino || "",
          "Motivo trabajo": s.motivoTrabajo || "",
          "Justificación": s.justificacion || "",
          "CAPEX": s.ordenInternaCAPEX || "",
          "Estado": s.estado || "",
          "Comentario supervisor": s.comentarioSupervisor || "",
          "Fecha registro": timestampCL(s.fechaRegistro),
          "Fecha revisión": timestampCL(s.fechaDecision),
          "Extensión": s.tieneExtension ? "Sí" : "No",
          "Motivo extensión": s.motivoExtension || "",
          "Hora inicio extensión": s.horaInicioExtension || "",
          "Hora término extensión": s.horaTerminoExtension || "",
          "Horas 50%": numero(s.horas50),
          "Horas 75%": numero(s.horas75),
          "Horas 100%": numero(s.horas100),
          "Horas totales": numero(s.horasTotales)
        })
      );


    const base = [];
    let n = 1;


    lista.forEach(
      s =>
        (s.tramosCalculados || [])
          .forEach(
            t => {

              const fecha = String(t.fecha || "");
              const anio = Number(fecha.slice(0, 4)) || "";
              const mes = Number(fecha.slice(5, 7)) || "";


              base.push({
                "Número": n++,
                "N° de personal": numeroPersonalDeSolicitud(s),
                "Nombre": s.Nombre || "",
                "Apellidos": s.Apellido || "",
                "RUT": s.Rut || "",
                "Fecha inicio": fechaCL(t.fecha),
                "Hr inicio": t.horaInicio || "",
                "Hr término": t.horaTermino === "00:00" ? "23:59" : t.horaTermino || "",
                "Horas trabajadas": numero(t.horas),
                "% hrs extras": numero(t.recargo) / 100,
                "Motivo del trabajo": String(t.origen || "").toLowerCase().includes("extensión")
                  ? `Extensión de trabajo ${s.motivoExtension || ""}`
                  : s.motivoTrabajo || "",
                "Estado": s.estado || "",
                "Comentarios": s.justificacion || "",
                "MES": mes,
                "Año": anio,
                "día": Number(fecha.slice(8, 10)) || "",
                "CAPEX": s.ordenInternaCAPEX || "",
                "Fecha Registro": timestampCL(s.fechaRegistro),
                "Fecha revisión": timestampCL(s.fechaDecision),
                "Extensión de jornada": s.tieneExtension ? "Sí" : "No",
                "Hora inicio extensión": s.horaInicioExtension || "",
                "Hora término extensión": s.horaTerminoExtension || "",
                "Nombre completo": nombreCompleto(s),
                "Área": s.Area || ""
              });
            }
          )
    );


    const wb =
      XLSX.utils.book_new();


    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet([
        {
          Concepto: "Período",
          Valor: etiquetaRangoPeriodo(rango)
        },
        {
          Concepto: "Total solicitudes",
          Valor: lista.length
        },
        {
          Concepto: "Generado",
          Valor: new Date().toLocaleString("es-CL")
        }
      ]),
      "Resumen"
    );


    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(sol),
      "Solicitudes"
    );


    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(base),
      "Base HHEE"
    );


    const desde =
      `${rango.yDesde}-${String(rango.mDesde).padStart(2,"0")}`;


    const hasta =
      `${rango.yHasta}-${String(rango.mHasta).padStart(2,"0")}`;


    XLSX.writeFile(
      wb,
      `HHEE_${desde}_a_${hasta}.xlsx`,
      {
        compression: true
      }
    );


    toast(
      "Excel generado.",
      "exito"
    );
  }


  /* =========================================================
     CONCILIACIÓN SAP
  ========================================================= */

  function guessHeader(
    headers,
    terms
  ) {

    const norm =
      headers.map(
        h => ({
          h,
          n:
            normalizarTexto(h)
        })
      );


    for (
      const term of terms
    ) {

      const x =
        norm.find(
          o =>
            o.n ===
            term
            ||
            o.n.includes(
              term
            )
        );


      if (x) {
        return x.h;
      }
    }


    return headers[0]
      ||
      "";
  }


  async function leerSap(file) {

    const wb =
      XLSX.read(
        await file.arrayBuffer(),
        {
          type:
            "array",
          cellDates:
            true
        }
      );


    const ws =
      wb.Sheets[
        wb.SheetNames[0]
      ];


    sapFilas =
      XLSX.utils.sheet_to_json(
        ws,
        {
          defval:
            "",
          raw:
            true
        }
      );


    sapCabeceras =
      sapFilas.length
        ?
          Object.keys(
            sapFilas[0]
          )
        :
          [];


    return sapFilas.length;
  }


  function optionsHeaders(sel) {

    return sapCabeceras
      .map(
        h =>
          `
            <option
              value="${esc(h)}"
              ${
                h ===
                sel
                  ?
                    "selected"
                  :
                    ""
              }
            >
              ${esc(h)}
            </option>
          `
      )
      .join("");
  }


  function renderMapeoSap() {

    const hId =
      guessHeader(
        sapCabeceras,
        [
          "numero de personal",
          "n° de personal",
          "nº de personal",
          "numero personal",
          "personal",
          "pernr",
          "empleado"
        ]
      );


    const hFecha =
      guessHeader(
        sapCabeceras,
        [
          "fecha",
          "fecha trabajo",
          "dia"
        ]
      );


    const hRec =
      guessHeader(
        sapCabeceras,
        [
          "recargo",
          "porcentaje",
          "% hrs extras",
          "tipo hora",
          "excepcion",
          "excepción"
        ]
      );


    const hHoras =
      guessHeader(
        sapCabeceras,
        [
          "horas",
          "cantidad",
          "hhee",
          "horas extras"
        ]
      );


    $("sapMapeo")
      .innerHTML =
        `
          <div class="bloque">

            <div class="concCabecera">
              <div>
                <h3>Comparación a realizar</h3>
                <p>
                  La aplicación cruzará <b>N° de personal + fecha + recargo</b>.
                  El nombre siempre se mostrará desde BD.HE para que Administración pueda identificar a la persona.
                </p>
              </div>
            </div>

            <div class="mensaje mensajeInfo">
              <b>Archivo SAP:</b> ${esc(sapArchivoNombre || "Archivo cargado")} ·
              <b>${sapFilas.length}</b> fila(s).
            </div>

            <details class="configAvanzada">
              <summary>Configuración de columnas SAP (solo revisar si cambió el formato)</summary>
              <div class="configContenido">
                <div class="grid4">

                  <div>
                    <label>N° de personal SAP</label>
                    <select id="sapColId">${optionsHeaders(hId)}</select>
                  </div>

                  <div>
                    <label>Fecha</label>
                    <select id="sapColFecha">${optionsHeaders(hFecha)}</select>
                  </div>

                  <div>
                    <label>Recargo</label>
                    <select id="sapColRecargo">${optionsHeaders(hRec)}</select>
                  </div>

                  <div>
                    <label>Horas</label>
                    <select id="sapColHoras">${optionsHeaders(hHoras)}</select>
                  </div>

                </div>
              </div>
            </details>

            <div class="acciones">
              <button class="btn btnPrimario" id="btnComparar">
                Conciliar BD.HE vs SAP
              </button>
            </div>

          </div>
        `;


    $("btnComparar")
      .addEventListener(
        "click",
        ejecutarConciliacion
      );
  }


  function fechaSAP(v) {

    if (
      v instanceof Date
      &&
      !Number.isNaN(v.getTime())
    ) {

      return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
    }


    if (
      typeof v === "number"
      &&
      window.XLSX
    ) {

      const p =
        XLSX.SSF.parse_date_code(v);


      if (p) {
        return `${p.y}-${String(p.m).padStart(2,"0")}-${String(p.d).padStart(2,"0")}`;
      }
    }


    const texto =
      String(v ?? "").trim();


    let m;


    if ((m = texto.match(/^(\d{4})[-\/]([01]?\d)[-\/]([0-3]?\d)/))) {
      return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
    }


    if ((m = texto.match(/^([0-3]?\d)[-\/]([01]?\d)[-\/](\d{4})/))) {
      return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
    }


    return "";
  }


  function claveConc(
    numeroPersonal,
    fecha,
    recargo
  ) {

    return `${numeroPersonal}|${fecha}|${recargo}`;
  }


  function resumenConciliacionPorTrabajador(resultados) {

    const mapa = new Map();


    resultados.forEach(
      fila => {

        const claveTrabajador =
          fila.claveTrabajador
          ||
          fila.numeroPersonal
          ||
          fila.nombre;


        if (!mapa.has(claveTrabajador)) {

          mapa.set(
            claveTrabajador,
            {
              claveTrabajador,
              numeroPersonal: fila.numeroPersonal || "",
              nombre: fila.nombre || "Sin nombre",
              area: fila.area || "",
              registradas: 0,
              bonificacion: 0,
              pagar: 0,
              sap: 0,
              diferencia: 0,
              incidencias: 0,
              estado: "Conforme",
              filas: []
            }
          );
        }


        const x = mapa.get(claveTrabajador);

        x.registradas += numero(fila.registradas);
        x.bonificacion += numero(fila.bonificacion);
        x.pagar += numero(fila.pagar);
        x.sap += numero(fila.sap);
        x.diferencia += numero(fila.diferencia);
        x.filas.push(fila);


        if (fila.estado !== "Conforme") {
          x.incidencias += 1;
          x.estado = "Revisar";
        }
      }
    );


    return [...mapa.values()]
      .map(
        x => ({
          ...x,
          registradas: red2(x.registradas),
          bonificacion: red2(x.bonificacion),
          pagar: red2(x.pagar),
          sap: red2(x.sap),
          diferencia: red2(x.diferencia)
        })
      )
      .sort(
        (a, b) =>
          a.nombre.localeCompare(b.nombre, "es")
      );
  }


  function textoResultadoTrabajador(x) {

    const problemas =
      x.filas
        .filter(f => f.estado !== "Conforme")
        .slice(0, 3)
        .map(
          f =>
            `${fechaCL(f.fecha)} ${f.recargo}%: ${vistaHoras(f.registradas)} h reales${f.bonificacion > 0 ? ` + ${vistaHoras(f.bonificacion)} h bonificación` : ""} = ${vistaHoras(f.pagar)} h a pagar / SAP ${vistaHoras(f.sap)} h`
        );


    if (!problemas.length) {
      return "Sin diferencias.";
    }


    const extra =
      x.incidencias > problemas.length
        ? `; y ${x.incidencias - problemas.length} incidencia(s) adicional(es).`
        : ".";


    return `${problemas.join("; ")}${extra}`;
  }


  async function ejecutarConciliacion() {

    const y = Number($("conAnio").value);
    const m = Number($("conMes").value);

    const cId = $("sapColId").value;
    const cF = $("sapColFecha").value;
    const cR = $("sapColRecargo").value;
    const cH = $("sapColHoras").value;


    if (!sapFilas.length) {
      return toast("Primero carga el Excel SAP.", "error");
    }


    try {

      await cargarUsuarios();


      const solicitudes =
        (
          await solicitudesPeriodo(y, m, true)
        )
        .filter(
          x =>
            normEstado(x.estado) === "aprobada"
        );


      const rango = rangoMes(y, m);


      const filasBd =
        filasAprobadasPorFechaRecargo(
          solicitudes,
          rango
        );


      const bd = new Map();
      const sap = new Map();
      const meta = new Map();


      filasBd.forEach(
        fila => {

          const numeroPersonal =
            normalizarNumeroPersonal(
              fila.numeroPersonal
            );


          const idCruce =
            claveNumeroPersonal(numeroPersonal)
            ||
            `__SINNUM__${fila.idInterno}`;


          const k =
            claveConc(
              idCruce,
              fila.fecha,
              fila.recargo
            );


          bd.set(
            k,
            {
              registradas: fila.horasRegistradas,
              bonificacion: fila.bonificacion,
              pagar: fila.horasPagar,
              solicitudes: fila.solicitudes || []
            }
          );


          meta.set(
            idCruce,
            {
              numeroPersonal,
              nombre: fila.nombre,
              area: fila.area,
              claveTrabajador: fila.idInterno
            }
          );
        }
      );


      const usuariosPorNumero = new Map();

      cacheUsuarios.forEach(
        u => {
          const np = claveNumeroPersonal(u.NumeroPersonal);
          if (np) usuariosPorNumero.set(np, u);
        }
      );


      const fechasSapValidas = [];


      sapFilas.forEach(
        row => {

          const numeroPersonalOriginal =
            normalizarNumeroPersonal(
              row[cId]
            );

          const numeroPersonal =
            claveNumeroPersonal(
              numeroPersonalOriginal
            );


          const fecha = fechaSAP(row[cF]);
          const recargo = recargoSapNumero(row[cR]);
          const horas = numero(row[cH]);


          if (fecha) {
            fechasSapValidas.push(fecha);
          }


          if (
            !numeroPersonal
            ||
            !fecha
            ||
            ![50, 75, 100].includes(recargo)
          ) {
            return;
          }


          if (
            fecha < rango.ini
            ||
            fecha >= rango.fin
          ) {
            return;
          }


          const k =
            claveConc(
              numeroPersonal,
              fecha,
              recargo
            );


          sap.set(
            k,
            red2(
              (sap.get(k) || 0)
              +
              horas
            )
          );


          if (!meta.has(numeroPersonal)) {

            const usuario =
              usuariosPorNumero.get(numeroPersonal);


            meta.set(
              numeroPersonal,
              {
                numeroPersonal: usuario?.NumeroPersonal || numeroPersonalOriginal || numeroPersonal,
                nombre: usuario ? nombreCompleto(usuario) : `N° personal ${numeroPersonalOriginal || numeroPersonal}`,
                area: usuario?.Area || "",
                claveTrabajador: usuario?.id || numeroPersonal
              }
            );
          }
        }
      );


      const filasSapPeriodo =
        [...sap.values()].length;


      if (
        filasSapPeriodo === 0
        &&
        fechasSapValidas.length
      ) {

        const ordenadas =
          [...fechasSapValidas].sort();


        throw new Error(
          `El Excel SAP no contiene registros válidos para ${MESES[m]} ${y}. `
          +
          `Las fechas detectadas van desde ${fechaCL(ordenadas[0])} hasta ${fechaCL(ordenadas[ordenadas.length - 1])}. Revisa el período o el archivo cargado.`
        );
      }


      const keys =
        [...new Set([
          ...bd.keys(),
          ...sap.keys()
        ])]
        .sort();


      const resultados =
        keys.map(
          k => {

            const partes = k.split("|");
            const idCruce = partes[0];
            const fecha = partes[1];
            const recargo = Number(partes[2]);

            const b = bd.get(k) || { registradas: 0, bonificacion: 0, pagar: 0, solicitudes: [] };
            const hs = red2(sap.get(k) || 0);

            const mta =
              meta.get(idCruce)
              ||
              {
                numeroPersonal: idCruce.startsWith("__SINNUM__") ? "" : idCruce,
                nombre: idCruce.startsWith("__SINNUM__") ? "Trabajador sin N° personal" : `N° personal ${idCruce}`,
                area: "",
                claveTrabajador: idCruce
              };


            const registradas = red2(b.registradas || 0);
            const bonificacion = red2(b.bonificacion || 0);
            const pagar = red2(b.pagar || 0);
            const diferencia = red2(hs - pagar);

            let estado = "Conforme";


            if (idCruce.startsWith("__SINNUM__")) {
              estado = "Sin N° personal";
            }
            else if (pagar === 0 && hs > 0) {
              estado = "Solo SAP";
            }
            else if (pagar > 0 && hs === 0) {
              estado = "N° no encontrado / sin horas SAP";
            }
            else if (Math.abs(diferencia) >= 0.01) {
              estado = "Revisar";
            }


            return {
              clave: k,
              claveTrabajador: mta.claveTrabajador || idCruce,
              numeroPersonal: mta.numeroPersonal || "",
              nombre: mta.nombre || "Sin nombre",
              area: mta.area || "",
              fecha,
              recargo,
              registradas,
              bonificacion,
              pagar,
              sap: hs,
              diferencia,
              estado,
              solicitudes: b.solicitudes || []
            };
          }
        );


      const resumenTrabajadores =
        resumenConciliacionPorTrabajador(
          resultados
        );


      ultimaConciliacion = {
        anio: y,
        mes: m,
        resultados,
        resumenTrabajadores,
        archivoSap: sapArchivoNombre || "",
        generado: new Date().toISOString(),
        generadoPor: nombreCompleto(perfilActual),
        reglaPagoConfigurada: REGLA_PAGO_CONFIGURADA,
        horasMinimasPago: HORAS_MINIMAS_PAGO
      };


      try {
        sessionStorage.setItem(
          "ultimaConciliacionHHEE",
          JSON.stringify(ultimaConciliacion)
        );
      }
      catch {}


      renderResultadoConciliacion();

    }
    catch (err) {

      console.error(err);

      toast(
        err.message || "No fue posible realizar la conciliación.",
        "error"
      );
    }
  }


  window.verDetalleConciliacion =
    encoded => {

      if (!ultimaConciliacion) return;


      const clave =
        decodeURIComponent(encoded);


      const trabajador =
        (ultimaConciliacion.resumenTrabajadores || [])
          .find(
            x =>
              x.claveTrabajador === clave
          );


      if (!trabajador) return;


      const filas =
        (ultimaConciliacion.resultados || [])
          .filter(
            x =>
              x.claveTrabajador === clave
          );


      modal(
        `
          <h3>${esc(trabajador.nombre)}</h3>

          <div class="mensaje mensajeInfo">
            <b>N° de personal:</b> ${esc(trabajador.numeroPersonal || "Sin registrar")} ·
            <b>Área:</b> ${esc(trabajador.area || "-")}
          </div>

          <div class="resumenKpi">
            <div class="kpi"><span>Registradas</span><strong>${vistaHoras(trabajador.registradas)} h</strong></div>
            <div class="kpi"><span>Bonificación</span><strong>${vistaHoras(trabajador.bonificacion)} h</strong></div>
            <div class="kpi"><span>Horas a pagar</span><strong>${vistaHoras(trabajador.pagar)} h</strong></div>
            <div class="kpi"><span>SAP</span><strong>${vistaHoras(trabajador.sap)} h</strong></div>
            <div class="kpi"><span>Diferencia</span><strong>${vistaHoras(trabajador.diferencia)} h</strong></div>
          </div>

          <div class="tablaWrap">
            <table class="tabla">
              <thead>
                <tr>
                  <th>Fecha</th>
                  <th>Recargo</th>
                  <th>Registradas</th>
                  <th>Bonificación</th>
                  <th>Horas a pagar</th>
                  <th>SAP</th>
                  <th>Diferencia</th>
                  <th>Estado</th>
                </tr>
              </thead>
              <tbody>
                ${
                  filas.map(
                    f =>
                      `
                        <tr>
                          <td>${fechaCL(f.fecha)}</td>
                          <td>${f.recargo}%</td>
                          <td>${vistaHoras(f.registradas)} h</td>
                          <td>
                            ${vistaHoras(f.bonificacion)} h
                            ${f.bonificacion > 0 ? `<span class="bonificacionPago">Mínimo ${vistaHoras(HORAS_MINIMAS_PAGO)} h</span>` : ""}
                          </td>
                          <td><b>${vistaHoras(f.pagar)} h</b></td>
                          <td>${vistaHoras(f.sap)} h</td>
                          <td>${vistaHoras(f.diferencia)} h</td>
                          <td>
                            <span class="estado ${f.estado === "Conforme" ? "ok" : "diferencia"}">${esc(f.estado)}</span>
                          </td>
                        </tr>
                      `
                  ).join("")
                }
              </tbody>
            </table>
          </div>

          <div class="acciones">
            <button class="btn btnSecundario" onclick="window.cerrarModal()">Cerrar</button>
          </div>
        `
      );
    };


  function tablaResumenConciliacion(lista) {

    if (!lista.length) {
      return `<div class="mensaje mensajeInfo">No hay trabajadores en esta sección.</div>`;
    }


    return `
      <div class="tablaWrap">
        <table class="tabla tablaResumenTrabajador">
          <thead>
            <tr>
              <th>N° personal</th>
              <th>Trabajador</th>
              <th>Horas registradas</th>
              <th>Bonificación</th>
              <th>Horas a pagar</th>
              <th>SAP</th>
              <th>Diferencia</th>
              <th>Incidencias</th>
              <th>Estado</th>
              <th>Resultado</th>
              <th>Detalle</th>
            </tr>
          </thead>
          <tbody>
            ${
              lista.map(
                x =>
                  `
                    <tr>
                      <td>${esc(x.numeroPersonal || "Sin N°")}</td>
                      <td>
                        <span class="nombrePrincipal">${esc(x.nombre)}</span>
                        <span class="subDato">${esc(x.area || "")}</span>
                      </td>
                      <td>${vistaHoras(x.registradas)} h</td>
                      <td>${vistaHoras(x.bonificacion)} h</td>
                      <td><b>${vistaHoras(x.pagar)} h</b></td>
                      <td>${vistaHoras(x.sap)} h</td>
                      <td>${vistaHoras(x.diferencia)} h</td>
                      <td>${x.incidencias}</td>
                      <td>
                        <span class="estado ${x.estado === "Conforme" ? "ok" : "estadoRevision"}">${esc(x.estado)}</span>
                      </td>
                      <td>${esc(textoResultadoTrabajador(x))}</td>
                      <td>
                        <button
                          class="btn btnSecundario"
                          onclick="window.verDetalleConciliacion('${encodeURIComponent(x.claveTrabajador)}')"
                        >
                          Ver detalle
                        </button>
                      </td>
                    </tr>
                  `
              ).join("")
            }
          </tbody>
        </table>
      </div>
    `;
  }


  function renderResultadoConciliacion() {

    if (!ultimaConciliacion) return;


    const resumen =
      ultimaConciliacion.resumenTrabajadores
      ||
      resumenConciliacionPorTrabajador(
        ultimaConciliacion.resultados || []
      );


    ultimaConciliacion.resumenTrabajadores = resumen;


    const conformes =
      resumen.filter(x => x.estado === "Conforme");


    const revisar =
      resumen.filter(x => x.estado !== "Conforme");


    const totalBonificacion =
      resumen.reduce((a, x) => a + numero(x.bonificacion), 0);


    const totalPagar =
      resumen.reduce((a, x) => a + numero(x.pagar), 0);


    const totalSap =
      resumen.reduce((a, x) => a + numero(x.sap), 0);


    $("conResultado").innerHTML =
      `
        <div class="notaReglaPago">
          <b>Regla de horas mínimas:</b> ${esc(textoEstadoReglaPago())}
          <br>
          La conciliación compara <b>Horas a pagar</b> contra SAP. Las horas reales y la bonificación se mantienen visibles por separado.
        </div>

        <div class="resumenKpi">
          <div class="kpi"><span>Trabajadores</span><strong>${resumen.length}</strong></div>
          <div class="kpi"><span>Conformes</span><strong>${conformes.length}</strong></div>
          <div class="kpi"><span>Por revisar</span><strong>${revisar.length}</strong></div>
          <div class="kpi"><span>Bonificación</span><strong>${vistaHoras(totalBonificacion)} h</strong></div>
          <div class="kpi"><span>Horas a pagar</span><strong>${vistaHoras(totalPagar)} h</strong></div>
          <div class="kpi"><span>Horas SAP</span><strong>${vistaHoras(totalSap)} h</strong></div>
        </div>

        <div class="acciones" style="justify-content:flex-start">
          <button class="btn btnPrimario" id="btnPdfCon">Generar PDF</button>
        </div>

        <div class="bloque">
          <h3>1. Trabajadores conformes</h3>
          <div class="nota">Registros donde las horas a pagar coinciden con SAP por N° de personal, fecha y recargo.</div>
          <div style="margin-top:12px">${tablaResumenConciliacion(conformes)}</div>
        </div>

        <div class="bloque">
          <h3>2. Trabajadores con diferencias</h3>
          <div class="nota">Casos que deben revisarse antes de autorizar el pago.</div>
          <div style="margin-top:12px">${tablaResumenConciliacion(revisar)}</div>
        </div>
      `;


    $("btnPdfCon")
      .addEventListener(
        "click",
        generarPdfConciliacion
      );
  }


  async function renderConciliacion() {

    $("contenidoModulo").innerHTML =
      `
        <div class="bloque">

          <div class="concCabecera">
            <div>
              <h3>Conciliación BD.HE vs SAP</h3>
              <p>
                Se consideran únicamente solicitudes aprobadas.
                El cruce se realiza por <b>N° de personal + fecha + recargo</b> y compara SAP contra las <b>Horas a pagar</b>.
                Cuando un grupo aprobado es menor a ${vistaHoras(HORAS_MINIMAS_PAGO)} h, la diferencia se muestra como <b>bonificación por mínimo</b> al mismo recargo.
              </p>
            </div>
          </div>

          <div class="grid3">
            ${selectPeriodo("con")}

            <div>
              <label>Excel SAP *</label>
              <input id="sapArchivo" type="file" accept=".xlsx,.xls,.csv">
            </div>
          </div>

          <div id="sapEstado" class="mensaje mensajeInfo" style="margin-top:14px">
            Selecciona el mismo período que corresponde al Excel SAP y carga el archivo.
          </div>

        </div>

        <div id="sapMapeo"></div>
        <div id="conResultado"></div>
      `;


    $("sapArchivo")
      .addEventListener(
        "change",
        async e => {

          const f = e.target.files?.[0];
          if (!f) return;


          try {

            sapArchivoNombre = f.name || "";
            const n = await leerSap(f);

            $("sapEstado").className = "mensaje mensajeExito";
            $("sapEstado").textContent = `Archivo leído correctamente: ${n} filas. Verifica el período y luego concilia.`;

            renderMapeoSap();

          }
          catch (err) {

            console.error(err);
            sapArchivoNombre = "";
            $("sapEstado").className = "mensaje mensajeError";
            $("sapEstado").textContent = "No se pudo leer el Excel SAP.";
          }
        }
      );


    if (ultimaConciliacion) {
      renderResultadoConciliacion();
    }
  }


  /* =========================================================
     PDF
  ========================================================= */

  function generarPdfConciliacion() {

    if (!ultimaConciliacion) {

      try {
        ultimaConciliacion =
          JSON.parse(
            sessionStorage.getItem("ultimaConciliacionHHEE")
            ||
            "null"
          );
      }
      catch {}
    }


    if (!ultimaConciliacion) {
      return toast(
        "Primero realiza una conciliación SAP.",
        "error"
      );
    }


    if (!window.jspdf?.jsPDF) {
      return toast(
        "No se cargó jsPDF.",
        "error"
      );
    }


    const { jsPDF } = window.jspdf;

    const pdf =
      new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4"
      });


    const resumen =
      ultimaConciliacion.resumenTrabajadores
      ||
      resumenConciliacionPorTrabajador(
        ultimaConciliacion.resultados || []
      );


    const conformes =
      resumen.filter(x => x.estado === "Conforme");


    const revisar =
      resumen.filter(x => x.estado !== "Conforme");


    const totalBonificacion =
      red2(
        resumen.reduce(
          (a, x) => a + numero(x.bonificacion),
          0
        )
      );


    const totalPagar =
      red2(
        resumen.reduce(
          (a, x) => a + numero(x.pagar),
          0
        )
      );


    const totalSap =
      red2(
        resumen.reduce(
          (a, x) => a + numero(x.sap),
          0
        )
      );


    const azul = [0, 59, 113];
    const gris = [86, 105, 125];
    const borde = [215, 226, 234];


    pdf.setTextColor(...azul);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(22);
    pdf.text("transelec", 14, 16);


    pdf.setFontSize(15);
    pdf.text(
      "Informe de conciliación de horas extras",
      283,
      14,
      { align: "right" }
    );


    pdf.setFont("helvetica", "normal");
    pdf.setTextColor(...gris);
    pdf.setFontSize(8);
    pdf.text(
      `Validación administrativa previa al pago · ${MESES[ultimaConciliacion.mes]} ${ultimaConciliacion.anio}`,
      283,
      19,
      { align: "right" }
    );


    pdf.setDrawColor(...azul);
    pdf.setLineWidth(.8);
    pdf.line(14, 23, 283, 23);


    pdf.setFillColor(247, 250, 252);
    pdf.setDrawColor(...borde);
    pdf.roundedRect(14, 27, 269, 10, 1.5, 1.5, "FD");

    pdf.setTextColor(...gris);
    pdf.setFontSize(7.5);
    pdf.text(
      `Fuente BD.HE: solicitudes aprobadas · Regla mínima: ${vistaHoras(HORAS_MINIMAS_PAGO)} h · SAP: ${ultimaConciliacion.archivoSap || "archivo cargado"} · Generado: ${new Date().toLocaleString("es-CL")}`,
      17,
      33
    );


    if (revisar.length) {
      pdf.setFillColor(255, 247, 230);
      pdf.setDrawColor(235, 190, 105);
      pdf.roundedRect(14, 40, 269, 12, 1.5, 1.5, "FD");
      pdf.setFont("helvetica", "bold");
      pdf.setTextColor(132, 84, 0);
      pdf.setFontSize(8);
      pdf.text("Revisión administrativa requerida", 18, 45);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(7);
      pdf.text(
        `${revisar.length} trabajador(es) presentan diferencias o incidencias y deben revisarse antes de autorizar el pago.`,
        18,
        49
      );
    }
    else {
      pdf.setFillColor(236, 249, 242);
      pdf.setDrawColor(150, 210, 179);
      pdf.roundedRect(14, 40, 269, 12, 1.5, 1.5, "FD");
      pdf.setFont("helvetica", "bold");
      pdf.setTextColor(24, 134, 91);
      pdf.setFontSize(8);
      pdf.text("Conciliación conforme", 18, 45);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(7);
      pdf.text("No se detectaron diferencias entre las horas a pagar y SAP.", 18, 49);
    }


    const tarjetas = [
      ["TRABAJADORES", String(resumen.length)],
      ["CONFORMES", String(conformes.length)],
      ["POR REVISAR", String(revisar.length)],
      ["BONIFICACIÓN", vistaHoras(totalBonificacion)],
      ["HORAS A PAGAR", vistaHoras(totalPagar)],
      ["HORAS SAP", vistaHoras(totalSap)]
    ];


    const anchoTarjeta = 42.2;

    tarjetas.forEach(
      (t, i) => {
        const x = 14 + i * 45.3;
        pdf.setFillColor(249, 251, 253);
        pdf.setDrawColor(...borde);
        pdf.roundedRect(x, 56, anchoTarjeta, 16, 1.5, 1.5, "FD");
        pdf.setTextColor(...gris);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(6.5);
        pdf.text(t[0], x + 3, 61);
        pdf.setTextColor(...azul);
        pdf.setFontSize(11);
        pdf.text(t[1], x + 3, 68);
      }
    );


    let yActual = 79;


    const dibujarSeccion = (
      titulo,
      lista,
      decision
    ) => {

      pdf.setTextColor(...azul);
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(9);
      pdf.text(titulo, 14, yActual);
      pdf.setFontSize(7);
      pdf.text(
        `${lista.length} trabajador(es)`,
        283,
        yActual,
        { align: "right" }
      );

      yActual += 3;


      const body =
        lista.length
          ?
            lista.map(
              x => [
                x.numeroPersonal || "Sin N°",
                x.nombre,
                `${vistaHoras(x.registradas)} h`,
                `${vistaHoras(x.bonificacion)} h`,
                `${vistaHoras(x.pagar)} h`,
                `${vistaHoras(x.sap)} h`,
                `${vistaHoras(x.diferencia)} h`,
                String(x.incidencias),
                decision
              ]
            )
          :
            [["", "No se identificaron trabajadores en esta sección.", "", "", "", "", "", "", ""]];


      if (typeof pdf.autoTable === "function") {

        pdf.autoTable({
          startY: yActual,
          margin: { left: 14, right: 14 },
          head: [[
            "N° personal",
            "Trabajador",
            "Registradas",
            "Bonificación",
            "Horas a pagar",
            "SAP",
            "Diferencia",
            "Inc.",
            "Decisión"
          ]],
          body,
          theme: "grid",
          styles: {
            fontSize: 6.4,
            cellPadding: 1.6,
            lineColor: borde,
            lineWidth: .2,
            textColor: [47, 63, 82]
          },
          headStyles: {
            fillColor: [238, 245, 249],
            textColor: azul,
            fontStyle: "bold"
          },
          columnStyles: {
            0: { cellWidth: 22 },
            1: { cellWidth: 48 },
            2: { cellWidth: 23, halign: "right" },
            3: { cellWidth: 25, halign: "right" },
            4: { cellWidth: 27, halign: "right" },
            5: { cellWidth: 22, halign: "right" },
            6: { cellWidth: 25, halign: "right" },
            7: { cellWidth: 12, halign: "center" },
            8: { cellWidth: 25, halign: "center", fontStyle: "bold" }
          },
          didParseCell(data) {
            if (
              data.section === "body"
              &&
              data.column.index === 8
              &&
              decision === "REVISAR"
            ) {
              data.cell.styles.textColor = [138, 91, 0];
            }
          }
        });

        yActual = (pdf.lastAutoTable?.finalY || yActual) + 9;
      }
    };


    dibujarSeccion(
      "1. Trabajadores conformes",
      conformes,
      "CONFORME"
    );


    dibujarSeccion(
      "2. Trabajadores con diferencias",
      revisar,
      "REVISAR"
    );


    {
      const yNota = Math.min(
        (pdf.lastAutoTable?.finalY || 175) + 7,
        193
      );
      pdf.setTextColor(132, 84, 0);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(6.5);
      pdf.text(
        `Regla de horas mínimas: si un trabajador registra menos de ${vistaHoras(HORAS_MINIMAS_PAGO)} h aprobadas para una misma fecha y recargo, se completa la diferencia como bonificación al mismo recargo.`,
        14,
        yNota
      );
    }


    const totalPaginas =
      pdf.getNumberOfPages();


    for (let i = 1; i <= totalPaginas; i++) {
      pdf.setPage(i);
      pdf.setTextColor(...gris);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(6.5);
      pdf.text(
        `Página ${i} de ${totalPaginas}`,
        283,
        203,
        { align: "right" }
      );
    }


    pdf.save(
      `Informe_Conciliacion_HHEE_${MESES[ultimaConciliacion.mes]}_${ultimaConciliacion.anio}.pdf`
    );


    toast(
      "PDF generado.",
      "exito"
    );
  }


  function renderInformePdf() {

    if (!ultimaConciliacion) {

      try {
        ultimaConciliacion =
          JSON.parse(
            sessionStorage.getItem("ultimaConciliacionHHEE")
            ||
            "null"
          );
      }
      catch {}
    }


    if (!ultimaConciliacion) {

      $("contenidoModulo").innerHTML =
        `
          <div class="mensaje mensajeAviso">
            Todavía no hay una conciliación cargada en esta sesión.
            Ve a <b>Conciliación SAP</b>, realiza el cruce y genera el PDF desde esa misma pestaña.
          </div>

          <div class="acciones" style="justify-content:flex-start">
            <button class="btn btnPrimario" id="irConc">Ir a Conciliación SAP</button>
          </div>
        `;


      $("irConc").addEventListener(
        "click",
        () => abrirModulo("conciliacion")
      );

      return;
    }


    const resumen =
      ultimaConciliacion.resumenTrabajadores
      ||
      resumenConciliacionPorTrabajador(
        ultimaConciliacion.resultados || []
      );


    $("contenidoModulo").innerHTML =
      `
        <div class="resumenKpi">
          <div class="kpi"><span>Período</span><strong style="font-size:16px">${MESES[ultimaConciliacion.mes]} ${ultimaConciliacion.anio}</strong></div>
          <div class="kpi"><span>Trabajadores</span><strong>${resumen.length}</strong></div>
          <div class="kpi"><span>Conformes</span><strong>${resumen.filter(x => x.estado === "Conforme").length}</strong></div>
          <div class="kpi"><span>Por revisar</span><strong>${resumen.filter(x => x.estado !== "Conforme").length}</strong></div>
        </div>

        <div class="acciones" style="justify-content:flex-start">
          <button class="btn btnPrimario" id="generarPdfFinal">Generar informe PDF</button>
        </div>
      `;


    $("generarPdfFinal").addEventListener(
      "click",
      generarPdfConciliacion
    );
  }


  /* =========================================================
     NAVEGACIÓN
  ========================================================= */

  async function abrirModulo(clave) {

    const modulosAdmin =
      new Set([
        "resumenAdmin",
        "todasSolicitudes",
        "horasPagar",
        "conciliacion",
        "exportar",
        "usuarios",
        "informePdf"
      ]);


    const esModuloAdmin =
      modulosAdmin.has(clave)
      &&
      esAdmin();


    document
      .querySelector(".contenedor")
      ?.classList
      .toggle(
        "adminModuloActivo",
        esModuloAdmin
      );


    document
      .querySelectorAll(".adminTab")
      .forEach(
        tab =>
          tab.classList.toggle(
            "activa",
            esModuloAdmin
            &&
            tab.dataset.modulo === clave
          )
      );


    $("panelModulo")
      .classList
      .remove(
        "hidden"
      );


    $("contenidoModulo")
      .innerHTML =
        `
          <div class="mensaje mensajeInfo">
            Cargando...
          </div>
        `;


    const titulos = {

      ingresoNormal: [
        "Ingresar horas extras",
        "Formulario normal de lunes a sábado."
      ],

      domingoFeriado: [
        "Domingos y feriados",
        "Formulario especial para domingo o feriado."
      ],

      misSolicitudes: [
        "Mis solicitudes",
        "Consulta tus solicitudes y edita o elimina solo las pendientes."
      ],

      misHoras: [
        "Mis horas",
        "Resumen de solicitudes aprobadas."
      ],

      pendientesSupervisor: [
        "Solicitudes pendientes",
        "Solicitudes de tus trabajadores asignados."
      ],

      historialSupervisor: [
        "Historial de revisión",
        "Solicitudes aprobadas o rechazadas."
      ],

      resumenAdmin: [
        "Resumen",
        "Indicadores generales de solicitudes y horas aprobadas."
      ],

      todasSolicitudes: [
        "Solicitudes",
        "Vista administrativa general de todas las áreas y estados."
      ],

      horasPagar: [
        "Horas a pagar",
        "Solicitudes aprobadas por supervisor y cálculo por recargo."
      ],

      usuarios: [
        "Usuarios",
        "Administración de perfiles y permisos."
      ],

      exportar: [
        "Exportar solicitudes",
        "Descarga un rango de meses en un solo Excel."
      ],

      conciliacion: [
        "Conciliación SAP",
        "Cruce por N° de personal, fecha y recargo antes del pago."
      ],

      informePdf: [
        "Informe PDF",
        "Informe final para el jefe de Administración."
      ]
    };


    $("tituloModulo")
      .textContent =
        titulos[clave]?.[0]
        ||
        "Módulo";


    $("descripcionModulo")
      .textContent =
        titulos[clave]?.[1]
        ||
        "";


    try {

      if (
        clave ===
        "ingresoNormal"
      ) {

        $("contenidoModulo")
          .innerHTML =
            formularioSolicitud(
              "Normal"
            );


        configurarFormularioSolicitud(
          "Normal"
        );

      }
      else if (
        clave ===
        "domingoFeriado"
      ) {

        $("contenidoModulo")
          .innerHTML =
            formularioSolicitud(
              "DomingoFeriado"
            );


        configurarFormularioSolicitud(
          "DomingoFeriado"
        );

      }
      else if (
        clave ===
        "misSolicitudes"
      ) {

        await renderMisSolicitudes();

      }
      else if (
        clave ===
        "misHoras"
      ) {

        await renderMisHoras();

      }
      else if (
        clave ===
        "pendientesSupervisor"
      ) {

        await renderSupervisor(
          true
        );

      }
      else if (
        clave ===
        "historialSupervisor"
      ) {

        await renderSupervisor(
          false
        );

      }
      else if (
        clave ===
        "resumenAdmin"
      ) {

        await renderResumenAdmin();

      }
      else if (
        clave ===
        "todasSolicitudes"
      ) {

        await renderTodasSolicitudes();

      }
      else if (
        clave ===
        "horasPagar"
      ) {

        await renderHorasPagar();

      }
      else if (
        clave ===
        "usuarios"
      ) {

        await renderUsuarios();

      }
      else if (
        clave ===
        "exportar"
      ) {

        await renderExportar();

      }
      else if (
        clave ===
        "conciliacion"
      ) {

        await renderConciliacion();

      }
      else if (
        clave ===
        "informePdf"
      ) {

        renderInformePdf();
      }

    }
    catch (err) {

      console.error(err);


      $("contenidoModulo")
        .innerHTML =
          `
            <div class="mensaje mensajeError">
              ${
                err.code?.includes(
                  "permission"
                )
                  ?
                    "Firestore rechazó la consulta por permisos. Revisa las reglas."
                  :
                    "Ocurrió un error al cargar este módulo."
              }
            </div>
          `;
    }


    if (esModuloAdmin) {

      $("barraAdminTabs")
        ?.scrollIntoView({
          behavior: "smooth",
          block: "start"
        });

    }
    else {

      $("panelModulo")
        .scrollIntoView({
          behavior: "smooth",
          block: "start"
        });
    }
  }


  window.abrirModulo =
    abrirModulo;


  document
    .querySelectorAll(
      "[data-modulo]"
    )
    .forEach(
      b =>
        b.addEventListener(
          "click",
          () =>
            abrirModulo(
              b.dataset.modulo
            )
        )
    );


  function cerrarPanel() {

    document
      .querySelector(".contenedor")
      ?.classList
      .remove("adminModuloActivo");


    document
      .querySelectorAll(".adminTab")
      .forEach(
        tab =>
          tab.classList.remove("activa")
      );


    if (esAdmin()) {
      $("introAdmin")?.classList.remove("hidden");
    }


    $("panelModulo")
      .classList
      .add(
        "hidden"
      );


    $("contenidoModulo")
      .innerHTML =
        "";
  }


  window.cerrarPanel =
    cerrarPanel;


  $("cerrarPanel")
    .addEventListener(
      "click",
      cerrarPanel
    );


  /* =========================================================
     SESIÓN
  ========================================================= */

  onAuthStateChanged(
    auth,
    async user => {

      if (!user) {

        usuarioActual =
          null;

        perfilActual =
          null;

        cacheSolicitudes.clear();
        cacheUsuarios.clear();
        cacheAutorizaciones.clear();

        fidActual =
          "";

        notificacionesActivas =
          false;


        mostrarLogin();

        return;
      }


      if (
        activacionEnProceso
      ) {

        mostrarCarga();

        return;
      }


      mostrarCarga();


      try {

        usuarioActual =
          user;


        perfilActual =
          await obtenerPerfil(
            user
          );


        pintarPerfil(
          perfilActual,
          user
        );


        mostrarApp();


        /*
          Si el usuario ya había dado permiso anteriormente,
          sincronizamos FCM automáticamente.

          Si nunca dio permiso, NO mostramos la ventana aquí.
          El permiso solo se solicita cuando pulsa:
          "Activar notificaciones".
        */

        await sincronizarNotificacionesSiCorresponde();


        /*
          Si el usuario llegó tocando una notificación:

          ?modulo=pendientesSupervisor&solicitud=ID
          o
          ?modulo=misSolicitudes&solicitud=ID

          Primero abrimos el módulo y después la solicitud exacta.
        */

        const params =
          new URLSearchParams(
            window.location.search
          );


        const modulo =
          params.get(
            "modulo"
          );


        const solicitudId =
          params.get(
            "solicitud"
          );


        if (modulo) {

          await abrirModulo(
            modulo
          );


          if (solicitudId) {

            let solicitud =
              cacheSolicitudes.get(
                solicitudId
              );


            if (!solicitud) {

              try {

                const solicitudSnap =
                  await getDoc(
                    doc(
                      db,
                      "solicitudes",
                      solicitudId
                    )
                  );


                if (
                  solicitudSnap.exists()
                ) {

                  solicitud = {
                    id:
                      solicitudSnap.id,

                    ...solicitudSnap.data()
                  };


                  cacheSolicitudes.set(
                    solicitudId,
                    solicitud
                  );
                }

              }
              catch (err) {

                console.warn(
                  "[HHEE] No se pudo cargar la solicitud de la notificación:",
                  err
                );
              }
            }


            if (solicitud) {

              window.verSolicitud(
                solicitudId
              );

            }
            else {

              toast(
                "La solicitud indicada por la notificación ya no está disponible.",
                "info"
              );
            }
          }


          history.replaceState(
            {},
            "",
            window.location.pathname
          );
        }

      }
      catch (err) {

        console.error(err);


        try {

          await signOut(
            auth
          );

        }
        catch {}


        mostrarLogin();


        loginMsg(
          err.message ===
          "PERFIL_NO_EXISTE"
            ?
              "La cuenta existe, pero todavía no tiene perfil HHEE. Si es tu primera vez, usa “Activar mi cuenta”."
            :
          err.message ===
          "USUARIO_INACTIVO"
            ?
              "Tu usuario está desactivado."
            :
              "No fue posible cargar tu perfil."
        );
      }
    }
  );

</script>

</body>
</html>
