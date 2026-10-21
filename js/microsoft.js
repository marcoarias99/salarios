// Conexión con Microsoft 365: iniciar sesión (ventana emergente) y enviar correos con Microsoft Graph
// desde la cuenta que inició sesión. Fase 1: se inicia sesión con la cuenta gestiones y se envía desde ella.
(function (root) {
  'use strict';

  const config = root.CONFIG_MICROSOFT || {};
  // Permisos delegados que se piden al iniciar sesión: leer el perfil y enviar correo como esa cuenta.
  const PERMISOS = ['User.Read', 'Mail.Send'];
  // Permiso para SharePoint (respaldo y lista de envíos). Con SharePoint configurado se pide al iniciar sesión.
  const PERMISOS_SHAREPOINT = ['Sites.ReadWrite.All'];
  const GRAPH_ENVIAR = 'https://graph.microsoft.com/v1.0/me/sendMail';
  const REINTENTOS_POR_LIMITE = 3;

  let app = null;
  let cuenta = null;

  const configurado = () => Boolean(config.clientId && config.tenantId);
  // El inicio de sesión necesita que el sitio se abra desde http(s), no con doble clic (file://).
  const desdeServidor = () => root.location.protocol === 'http:' || root.location.protocol === 'https:';

  function estado() {
    if (!root.msal) return 'sin-libreria';
    if (!configurado()) return 'sin-configurar';
    if (!desdeServidor()) return 'archivo-local';
    return cuenta ? 'conectado' : 'desconectado';
  }

  async function iniciar() {
    if (estado() !== 'desconectado') return null;
    app = new root.msal.PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: new URL('redirect.html', root.location.href).href,
      },
      // La sesión dura mientras la pestaña esté abierta; al cerrarla hay que volver a iniciar sesión.
      cache: { cacheLocation: 'sessionStorage' },
    });
    await app.initialize();
    cuenta = app.getAllAccounts()[0] || null;
    if (cuenta) app.setActiveAccount(cuenta);
    return cuenta;
  }

  async function iniciarSesion() {
    if (!app) await iniciar();
    const permisos = config.sharepoint && config.sharepoint.sitio ? [...PERMISOS, ...PERMISOS_SHAREPOINT] : PERMISOS;
    const respuesta = await app.loginPopup({ scopes: permisos, prompt: 'select_account' });
    cuenta = respuesta.account;
    app.setActiveAccount(cuenta);
    return cuenta;
  }

  // Cierra la sesión solo en este sitio (no en Outlook ni en el resto de Microsoft 365).
  async function cerrarSesion() {
    if (app && cuenta) await app.clearCache({ account: cuenta });
    cuenta = null;
  }

  async function obtenerToken(permisos = PERMISOS) {
    try {
      return (await app.acquireTokenSilent({ scopes: permisos, account: cuenta })).accessToken;
    } catch (e) {
      if (e instanceof root.msal.InteractionRequiredAuthError) {
        return (await app.acquireTokenPopup({ scopes: permisos, account: cuenta })).accessToken;
      }
      throw e;
    }
  }

  const esperar = ms => new Promise(r => setTimeout(r, ms));

  // Mensajes en español para los errores más comunes al enviar.
  function mensajeDeError(status, detalle) {
    const codigo = detalle && detalle.error ? detalle.error.code : '';
    if (status === 401) return 'La sesión venció; inicie sesión de nuevo';
    if (status === 403) return 'Sin permiso para enviar correo (permiso Mail.Send de la aplicación)';
    if (codigo === 'ErrorInvalidRecipients' || status === 400) return 'Dirección de correo no válida';
    if (codigo === 'ErrorQuotaExceeded' || codigo === 'ErrorMessageSubmissionBlocked') return 'El buzón no puede enviar más correos por ahora';
    const texto = detalle && detalle.error ? detalle.error.message : '';
    return `Error ${status}${texto ? ': ' + texto : ''}`;
  }

  // Envía un correo HTML desde la cuenta conectada y deja copia en Elementos enviados.
  // Microsoft responde 202 cuando acepta el correo; si pide esperar (429), se reintenta.
  async function enviarCorreo({ para, nombre, asunto, html }) {
    if (!cuenta) throw new Error('No hay sesión iniciada con Microsoft');
    const cuerpo = JSON.stringify({
      message: {
        subject: asunto,
        body: { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: para, name: nombre || para } }],
      },
      saveToSentItems: true,
    });
    for (let intento = 0; ; intento++) {
      let respuesta;
      try {
        respuesta = await fetch(GRAPH_ENVIAR, {
          method: 'POST',
          headers: { Authorization: `Bearer ${await obtenerToken()}`, 'Content-Type': 'application/json' },
          body: cuerpo,
        });
      } catch (e) {
        throw new Error('Sin conexión con Microsoft; revise el internet');
      }
      if (respuesta.status === 202) return;
      if (respuesta.status === 429 && intento < REINTENTOS_POR_LIMITE) {
        await esperar((Number(respuesta.headers.get('Retry-After')) || 10) * 1000);
        continue;
      }
      throw new Error(mensajeDeError(respuesta.status, await respuesta.json().catch(() => null)));
    }
  }

  // ---------- SharePoint ----------

  const sitios = new Map();

  function mensajeDeSharePoint(status, detalle) {
    if (status === 401) return 'La sesión venció; inicie sesión de nuevo';
    if (status === 403) return 'Sin permiso para guardar en SharePoint: la cuenta necesita acceso de edición al sitio';
    if (status === 404) return 'No se encontró el sitio de SharePoint; revise la dirección en js/config.js';
    const texto400 = detalle && detalle.error ? detalle.error.message : '';
    if (status === 400 && /field|column/i.test(texto400)) return `A la lista de SharePoint le falta una columna o tiene otro nombre (${texto400})`;
    if (status === 507) return 'El espacio de SharePoint está lleno';
    const texto = detalle && detalle.error ? detalle.error.message : '';
    return `Error ${status} de SharePoint${texto ? ': ' + texto : ''}`;
  }

  async function graphSharePoint(url, opciones = {}) {
    let respuesta;
    try {
      const token = await obtenerToken(PERMISOS_SHAREPOINT);
      respuesta = await fetch(url, { ...opciones, headers: { Authorization: `Bearer ${token}`, ...(opciones.headers || {}) } });
    } catch (e) {
      if (e && e.errorCode) throw new Error(mensajeDeSesion(e) || 'Se canceló el permiso para SharePoint');
      throw new Error('Sin conexión con Microsoft; revise el internet');
    }
    const datos = await respuesta.json().catch(() => null);
    if (!respuesta.ok) throw new Error(mensajeDeSharePoint(respuesta.status, datos));
    return datos;
  }

  // "https://empresa.sharepoint.com/sites/Planillas" → id del sitio (se guarda para no buscarlo cada vez).
  async function idDelSitio(direccion) {
    if (!sitios.has(direccion)) {
      const url = new URL(direccion);
      const ruta = url.pathname.replace(/\/+$/, '');
      const sitio = await graphSharePoint(`https://graph.microsoft.com/v1.0/sites/${url.hostname}:${ruta}?$select=id`);
      sitios.set(direccion, sitio.id);
    }
    return sitios.get(direccion);
  }

  // Sube un archivo a la biblioteca "Documentos" del sitio. Las carpetas de la ruta se crean solas y un
  // archivo con el mismo nombre se reemplaza. Devuelve la dirección web del archivo.
  async function subirArchivo({ sitio, ruta, contenido, tipo }) {
    if (!cuenta) throw new Error('No hay sesión iniciada con Microsoft');
    const id = await idDelSitio(sitio);
    const rutaUrl = ruta.split('/').map(encodeURIComponent).join('/');
    const archivo = await graphSharePoint(`https://graph.microsoft.com/v1.0/sites/${id}/drive/root:/${rutaUrl}:/content`, {
      method: 'PUT',
      headers: { 'Content-Type': tipo || 'application/octet-stream' },
      body: contenido,
    });
    return archivo.webUrl;
  }

  // ---------- Lista de envíos ----------
  // Una fila por boleta enviada. Title lleva la quincena ("2026-10-Q2") y está indexada para buscar rápido.

  const listas = new Map();

  async function idDeLista(sitio, nombre) {
    const clave = `${sitio}|${nombre}`;
    if (!listas.has(clave)) {
      const id = await idDelSitio(sitio);
      const datos = await graphSharePoint(`https://graph.microsoft.com/v1.0/sites/${id}/lists?$select=id,displayName&$top=200`);
      const lista = (datos.value || []).find(l => l.displayName.trim().toLowerCase() === nombre.trim().toLowerCase());
      if (!lista) throw new Error(`No se encontró la lista "${nombre}" en SharePoint`);
      listas.set(clave, { sitio: id, lista: lista.id });
    }
    return listas.get(clave);
  }

  // Todas las filas de una quincena (sigue las páginas si hay muchas).
  async function leerEnvios({ sitio, lista, quincena }) {
    if (!cuenta) throw new Error('No hay sesión iniciada con Microsoft');
    const ids = await idDeLista(sitio, lista);
    const filtro = encodeURIComponent(`fields/Title eq '${quincena.replace(/'/g, "''")}'`);
    let url = `https://graph.microsoft.com/v1.0/sites/${ids.sitio}/lists/${ids.lista}/items` +
      `?$expand=fields($select=Title,Cedula,Nombre,Correo,Liquido,Fecha,EnviadoPor)&$filter=${filtro}&$top=500`;
    const filas = [];
    while (url) {
      const datos = await graphSharePoint(url, { headers: { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
      (datos.value || []).forEach(item => filas.push(item.fields || {}));
      url = datos['@odata.nextLink'] || null;
    }
    return filas;
  }

  async function registrarEnvio({ sitio, lista, campos }) {
    if (!cuenta) throw new Error('No hay sesión iniciada con Microsoft');
    const ids = await idDeLista(sitio, lista);
    await graphSharePoint(`https://graph.microsoft.com/v1.0/sites/${ids.sitio}/lists/${ids.lista}/items`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields: campos }),
    });
  }

  // Errores al iniciar sesión, en palabras simples. Devuelve null si la persona cerró la ventana.
  function mensajeDeSesion(e) {
    const codigo = (e && (e.errorCode || e.name)) || '';
    const texto = (e && e.message) || '';
    if (codigo === 'user_cancelled') return null;
    if (codigo === 'popup_window_error' || codigo === 'empty_window_error') {
      return 'El navegador bloqueó la ventana de inicio de sesión. Permita ventanas emergentes para este sitio e intente de nuevo.';
    }
    if (/AADSTS65001|AADSTS90094|consent/i.test(texto)) {
      return 'La aplicación necesita aprobación de un administrador de Microsoft 365 (IT) antes de poder usarse.';
    }
    if (/AADSTS50011/.test(texto)) return 'La dirección del sitio no está registrada en la aplicación (Redirect URI). Revise el registro en Entra ID.';
    if (/AADSTS700016|AADSTS90002/.test(texto)) return 'El Client ID o el Tenant ID de js/config.js no son correctos.';
    return `No se pudo iniciar sesión: ${texto || codigo}`;
  }

  root.Microsoft = {
    estado,
    iniciar,
    iniciarSesion,
    cerrarSesion,
    enviarCorreo,
    subirArchivo,
    leerEnvios,
    registrarEnvio,
    mensajeDeSesion,
    cuenta: () => cuenta,
    sharepoint: () => config.sharepoint || {},
    correo: () => (cuenta ? cuenta.username : ''),
  };
})(this);
