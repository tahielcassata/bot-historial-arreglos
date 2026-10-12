/**
 * Bot de historial de arreglos
 * Telegram + Google Sheets + Google Apps Script
 *
 * Funciones para correr a mano desde el editor (menú de arriba):
 *   configurarTodo()   -> crea pestañas, columnas y la carpeta de boletas en Drive
 *   conectarTelegram() -> conecta el bot con esta aplicación web
 *   estadoTelegram()   -> muestra si la conexión con Telegram anda bien
 *
 * Propiedades del script (Configuración del proyecto > Propiedades del script):
 *   TELEGRAM_TOKEN  -> el token de @BotFather (lo pega el dueño, nunca va en el código)
 *   FOLDER_ID       -> lo crea configurarTodo()
 *   WEBHOOK_SECRET  -> lo crea configurarTodo()
 *   WEBAPP_URL      -> opcional, solo si conectarTelegram() no encuentra la URL sola
 */

const HOJA_ARREGLOS = 'Arreglos';
const HOJA_CONFIG = 'Configuración';
const ENCABEZADOS = ['ID', 'Fecha', 'Celular', 'Cliente', 'Bici', 'Trabajo', 'Precio', 'Foto', 'Foto Telegram'];
const COL = { ID: 1, FECHA: 2, CELULAR: 3, CLIENTE: 4, BICI: 5, TRABAJO: 6, PRECIO: 7, FOTO: 8, FOTO_TG: 9 };
const MAX_RESULTADOS = 5;
const MAX_TEXTO = 500;
const MAX_QR = 1200; // más texto que esto hace un QR difícil de escanear
const MAX_BICIS = 4; // bicis que se ofrecen al cargar un cliente que ya existe

const BTN_NUEVO = 'Nuevo';
const BTN_CONSULTAR = 'Consultar';
const BTN_CANCELAR = 'Cancelar';

const TECLADO_MENU = {
  keyboard: [[{ text: BTN_NUEVO }, { text: BTN_CONSULTAR }]],
  resize_keyboard: true,
  is_persistent: true,
};
const TECLADO_CANCELAR = {
  keyboard: [[{ text: BTN_CANCELAR }]],
  resize_keyboard: true,
  is_persistent: true,
};

// =====================================================================
// Funciones para correr a mano
// =====================================================================

function configurarTodo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Pestaña Arreglos
  let hoja = ss.getSheetByName(HOJA_ARREGLOS);
  if (!hoja) {
    const primera = ss.getSheets()[0];
    const primeraVacia = primera.getLastRow() === 0 && primera.getName() !== HOJA_CONFIG;
    hoja = primeraVacia ? primera.setName(HOJA_ARREGLOS) : ss.insertSheet(HOJA_ARREGLOS);
  }
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, ENCABEZADOS.length)
      .setValues([ENCABEZADOS])
      .setFontWeight('bold')
      .setBackground('#e8eaed');
    hoja.setFrozenRows(1);
  }
  hoja.getRange('B:C').setNumberFormat('@'); // fecha y celular como texto
  hoja.setColumnWidth(COL.CLIENTE, 160);
  hoja.setColumnWidth(COL.BICI, 180);
  hoja.setColumnWidth(COL.TRABAJO, 300);
  hoja.hideColumns(COL.FOTO_TG);

  // Pestaña Configuración
  let conf = ss.getSheetByName(HOJA_CONFIG);
  if (!conf) {
    conf = ss.insertSheet(HOJA_CONFIG);
    conf.getRange('B2:B3').setNumberFormat('@');
    conf.getRange('A1:B3').setValues([
      ['Clave', 'Valor'],
      ['Nombre del negocio', 'Bicicletería'],
      ['Usuarios autorizados', ''],
    ]);
    conf.getRange('A1:B1').setFontWeight('bold').setBackground('#e8eaed');
    conf.getRange('A5').setValue(
      'Usuarios autorizados: IDs de Telegram separados por coma. ' +
      'Si alguien no autorizado le escribe al bot, el bot le muestra su ID.'
    );
    conf.setColumnWidth(1, 200);
    conf.setColumnWidth(2, 320);
  }

  // Carpeta de boletas en Drive y clave secreta del webhook
  const props = PropertiesService.getScriptProperties();
  let carpeta = null;
  const folderId = props.getProperty('FOLDER_ID');
  if (folderId) {
    try { carpeta = DriveApp.getFolderById(folderId); } catch (e) { carpeta = null; }
  }
  if (!carpeta) {
    carpeta = DriveApp.createFolder('Boletas - ' + ss.getName());
    props.setProperty('FOLDER_ID', carpeta.getId());
  }
  if (!props.getProperty('WEBHOOK_SECRET')) {
    props.setProperty('WEBHOOK_SECRET', Utilities.getUuid().replace(/-/g, ''));
  }

  Logger.log('Listo. Hoja configurada. Carpeta de boletas: ' + carpeta.getUrl());
}

function conectarTelegram() {
  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty('WEBHOOK_SECRET');
  if (!secret) throw new Error('Primero corré configurarTodo().');
  getToken_(); // avisa si falta el token

  const url = (props.getProperty('WEBAPP_URL') || ScriptApp.getService().getUrl() || '').trim();
  if (!/\/exec$/.test(url)) {
    throw new Error(
      'No encontré la URL /exec de la aplicación web. Publicá el script ' +
      '(Implementar > Nueva implementación > Aplicación web) o cargá WEBAPP_URL en las propiedades.'
    );
  }

  const r = tg_('setWebhook', {
    url: url + '?secret=' + secret,
    allowed_updates: ['message', 'callback_query'],
    drop_pending_updates: true,
  });
  Logger.log(r.ok ? 'Conectado con Telegram.' : 'Telegram respondió: ' + r.description);
}

function estadoTelegram() {
  const yo = tg_('getMe', {});
  const info = tg_('getWebhookInfo', {});
  if (info.result && info.result.url) {
    info.result.url = info.result.url.replace(/secret=.*/, 'secret=***');
  }
  Logger.log('Bot: ' + (yo.ok ? '@' + yo.result.username : 'token inválido'));
  Logger.log(JSON.stringify(info.result, null, 2));
}

// =====================================================================
// Entrada desde Telegram
// =====================================================================

function doGet() {
  return ContentService.createTextOutput('El bot está funcionando.');
}

function doPost(e) {
  let chatId = null;
  try {
    const secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');
    if (!e || !e.parameter || !secret || e.parameter.secret !== secret) return ok_();
    if (!e.postData || !e.postData.contents) return ok_();

    const update = JSON.parse(e.postData.contents);

    // Telegram puede reenviar el mismo mensaje: se procesa una sola vez
    const cache = CacheService.getScriptCache();
    const clave = 'upd_' + update.update_id;
    if (cache.get(clave)) return ok_();
    cache.put(clave, '1', 21600);

    if (update.message) {
      chatId = update.message.chat.id;
      manejarMensaje_(update.message);
    } else if (update.callback_query) {
      chatId = update.callback_query.message ? update.callback_query.message.chat.id : null;
      manejarBoton_(update.callback_query);
    }
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    if (chatId) {
      try { enviar_(chatId, 'Hubo un error. Probá de nuevo o tocá Cancelar.', TECLADO_MENU); } catch (e2) {}
    }
  }
  return ok_();
}

// Se responde con HtmlService porque ContentService hace una redirección (302)
// y Telegram la toma como error y reintenta el mismo mensaje sin parar.
function ok_() {
  return HtmlService.createHtmlOutput('ok');
}

// =====================================================================
// Mensajes
// =====================================================================

function manejarMensaje_(msg) {
  if (!msg.chat || msg.chat.type !== 'private' || !msg.from) return;
  const chatId = msg.chat.id;
  const conf = leerConfig_();

  if (!estaAutorizado_(msg.from.id, conf)) {
    enviar_(chatId,
      'No estás autorizado para usar este bot.\n' +
      'Tu ID de Telegram es: ' + msg.from.id + '\n' +
      'Pedile al dueño que lo agregue en la hoja, pestaña Configuración.');
    return;
  }

  const texto = (msg.text || '').trim();

  if (texto === '/start' || texto === '/menu') {
    borrarEstado_(chatId);
    enviar_(chatId, 'Hola 👋 Elegí una opción:', TECLADO_MENU);
    return;
  }
  if (texto === BTN_CANCELAR || texto === '/cancelar') {
    borrarEstado_(chatId);
    enviar_(chatId, 'Cancelado.', TECLADO_MENU);
    return;
  }
  if (texto === BTN_NUEVO || texto === '/nuevo') {
    setEstado_(chatId, { paso: 'celular', datos: {} });
    enviar_(chatId, 'Nuevo arreglo.\nCelular del cliente (solo números):', TECLADO_CANCELAR);
    return;
  }
  if (texto === BTN_CONSULTAR || texto === '/consultar') {
    setEstado_(chatId, { paso: 'consulta' });
    enviar_(chatId, 'Escribí el celular del cliente:', TECLADO_CANCELAR);
    return;
  }

  const st = getEstado_(chatId);

  // Esperando la foto de la boleta
  if (st && st.paso === 'foto') {
    const foto = extraerFoto_(msg);
    if (foto) {
      guardarArreglo_(chatId, st.datos, foto);
      return;
    }
    if (texto.toLowerCase() === 'sin foto') {
      guardarArreglo_(chatId, st.datos, null);
      return;
    }
    enviar_(chatId, 'Mandá la foto de la boleta.\nSi cambiaste de idea, escribí "sin foto".');
    return;
  }

  if (!st) {
    enviar_(chatId, 'Elegí una opción:', TECLADO_MENU);
    return;
  }
  if (!texto) {
    enviar_(chatId, 'Respondé con texto, o tocá Cancelar.');
    return;
  }

  const valor = texto.slice(0, MAX_TEXTO);

  switch (st.paso) {
    case 'celular': {
      const cel = normalizarCelular_(valor);
      if (!cel) {
        enviar_(chatId, 'Ese número no parece válido. Escribí solo los números, por ejemplo 1123456789.');
        return;
      }
      st.datos.celular = cel;
      // Si el celular ya está guardado, no vuelve a pedir el nombre y ofrece sus bicis
      const previos = filasDeCelular_(cel);
      if (previos.length) {
        st.datos.cliente = String(previos[0][COL.CLIENTE - 1]);
        st.bicis = bicisDistintas_(previos);
        st.paso = 'pregunta_bici';
        // El botón lleva solo el índice: el texto de la bici no entra en callback_data
        const botones = st.bicis.map(function (b, i) {
          return [{ text: '🚲 ' + recortar_(b, 30), callback_data: 'bici_' + i }];
        });
        botones.push([{ text: '➕ Otra bici', callback_data: 'bici_otra' }]);
        botones.push([{ text: '✖ Cancelar', callback_data: 'agregar_no' }]);
        const r = tg_('sendMessage', {
          chat_id: chatId,
          text: 'Este celular ya está guardado: ' + st.datos.cliente + ', ' + previos.length + ' arreglo(s).\n' +
            '¿En qué bici es el trabajo nuevo?',
          reply_markup: { inline_keyboard: botones },
        });
        // Solo valen los botones de este mensaje (los índices de otro apuntarían a otra lista)
        st.msgId = r.ok && r.result ? r.result.message_id : null;
        setEstado_(chatId, st);
        return;
      }
      st.paso = 'cliente';
      setEstado_(chatId, st);
      enviar_(chatId, 'Nombre del cliente:');
      return;
    }
    case 'cliente':
      st.datos.cliente = valor;
      st.paso = 'bici';
      setEstado_(chatId, st);
      enviar_(chatId, 'Bici (marca, modelo, color):');
      return;
    case 'bici':
      st.datos.bici = valor;
      st.paso = 'trabajo';
      setEstado_(chatId, st);
      enviar_(chatId, 'Trabajo realizado:');
      return;
    case 'trabajo':
      st.datos.trabajo = valor;
      st.paso = 'precio';
      setEstado_(chatId, st);
      enviar_(chatId, 'Precio (solo el número, ej. 15000):');
      return;
    case 'precio': {
      const precio = normalizarPrecio_(valor);
      if (precio === null) {
        enviar_(chatId, 'No entendí el precio. Escribí solo el número, por ejemplo 15000.');
        return;
      }
      st.datos.precio = precio;
      st.paso = 'pregunta_foto';
      setEstado_(chatId, st);
      tg_('sendMessage', {
        chat_id: chatId,
        text: '¿Querés adjuntar la foto de la boleta?',
        reply_markup: { inline_keyboard: [[
          { text: 'Sí', callback_data: 'foto_si' },
          { text: 'No', callback_data: 'foto_no' },
        ]] },
      });
      return;
    }
    case 'pregunta_bici':
      enviar_(chatId, 'Tocá una opción en el mensaje de arriba.');
      return;
    case 'pregunta_foto':
      enviar_(chatId, 'Tocá Sí o No en el mensaje de arriba.');
      return;
    case 'consulta': {
      const cel = normalizarCelular_(valor);
      if (!cel) {
        enviar_(chatId, 'Ese número no parece válido. Escribí solo los números, por ejemplo 1123456789.');
        return;
      }
      borrarEstado_(chatId);
      mostrarArreglos_(chatId, cel);
      return;
    }
    default:
      borrarEstado_(chatId);
      enviar_(chatId, 'Elegí una opción:', TECLADO_MENU);
  }
}

// =====================================================================
// Botones debajo de los mensajes
// =====================================================================

function manejarBoton_(cq) {
  if (!cq.message) return;
  const chatId = cq.message.chat.id;
  const conf = leerConfig_();

  if (!estaAutorizado_(cq.from.id, conf)) {
    tg_('answerCallbackQuery', { callback_query_id: cq.id, text: 'No autorizado' });
    return;
  }
  tg_('answerCallbackQuery', { callback_query_id: cq.id });
  // Saca los botones para que no se toquen dos veces
  tg_('editMessageReplyMarkup', {
    chat_id: chatId,
    message_id: cq.message.message_id,
    reply_markup: { inline_keyboard: [] },
  });

  const data = cq.data || '';

  // 'agregar_si' ya no se manda: quedó en botones de chats de antes
  const mBici = data.match(/^bici_(\d+)$/);
  if (mBici || data === 'bici_otra' || data === 'agregar_no' || data === 'agregar_si') {
    const st = getEstado_(chatId);
    const vigente = st && st.paso === 'pregunta_bici' && data !== 'agregar_si' &&
      (!st.msgId || st.msgId === cq.message.message_id);
    const bici = vigente && mBici ? st.bicis[Number(mBici[1])] : null;
    if (!vigente || (mBici && !bici)) {
      if (st) {
        // Hay otra carga en curso: no se le cambia el teclado
        enviar_(chatId, 'Ese botón es de una carga anterior, ya no está vigente.');
      } else {
        enviar_(chatId, 'Esa carga ya terminó o se canceló. Tocá Nuevo para empezar otra.', TECLADO_MENU);
      }
      return;
    }
    if (data === 'agregar_no') {
      borrarEstado_(chatId);
      enviar_(chatId, 'Cancelado.', TECLADO_MENU);
      return;
    }
    delete st.bicis;
    delete st.msgId;
    if (mBici) {
      st.datos.bici = bici;
      st.paso = 'trabajo';
      setEstado_(chatId, st);
      enviar_(chatId, 'Trabajo realizado:');
    } else {
      st.paso = 'bici';
      setEstado_(chatId, st);
      enviar_(chatId, 'Bici (marca, modelo, color):');
    }
    return;
  }

  if (data === 'foto_si' || data === 'foto_no') {
    const st = getEstado_(chatId);
    if (!st || st.paso !== 'pregunta_foto') {
      enviar_(chatId, 'Esa carga ya terminó o se canceló. Tocá Nuevo para empezar otra.', TECLADO_MENU);
      return;
    }
    if (data === 'foto_si') {
      st.paso = 'foto';
      setEstado_(chatId, st);
      enviar_(chatId, 'Mandá la foto de la boleta.\nSi cambiaste de idea, escribí "sin foto".');
    } else {
      guardarArreglo_(chatId, st.datos, null);
    }
    return;
  }

  if (data === 'qr_no') {
    enviar_(chatId, 'Listo.', TECLADO_MENU);
    return;
  }

  const m = data.match(/^qrc_(\d+)$/);
  if (m) {
    enviarQR_(chatId, m[1], conf);
    return;
  }

  // Botones "QR del #N" que quedaron en chats de antes: ya no hay QR por arreglo
  if (/^qr_\d+$/.test(data)) {
    enviar_(chatId, 'Ese botón es viejo. Consultá de nuevo el celular para obtener el QR del cliente.', TECLADO_MENU);
  }
}

// =====================================================================
// Guardar
// =====================================================================

function guardarArreglo_(chatId, datos, foto) {
  if (!datos || !datos.celular) {
    borrarEstado_(chatId);
    enviar_(chatId, 'Se perdieron los datos de la carga. Tocá Nuevo para empezar de nuevo.', TECLADO_MENU);
    return;
  }

  // La foto se baja antes de bloquear la hoja, porque tarda
  let link = '';
  if (foto) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const nombre = 'Boleta_' + datos.celular + '_' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HH-mm-ss');
    link = guardarFotoEnDrive_(foto.fileId, nombre);
    if (!link) {
      enviar_(chatId, 'No pude guardar la foto en Drive. Mandala de nuevo, o escribí "sin foto".');
      return; // queda esperando la foto
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let id;
  try {
    const hoja = hojaArreglos_();
    id = siguienteId_(hoja);
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const fecha = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    hoja.appendRow([
      id,
      fecha,
      datos.celular,
      seguro_(datos.cliente),
      seguro_(datos.bici),
      seguro_(datos.trabajo),
      datos.precio,
      link,
      foto ? foto.tipo + ':' + foto.fileId : '',
    ]);
  } finally {
    lock.releaseLock();
  }

  borrarEstado_(chatId);
  enviar_(chatId,
    '✅ Listo, guardado.\n' +
    'Arreglo #' + id + ' · celular ' + datos.celular + (foto ? ' · con foto' : ''),
    TECLADO_MENU);
}

function guardarFotoEnDrive_(fileId, nombre) {
  const info = tg_('getFile', { file_id: fileId });
  if (!info.ok || !info.result.file_path) return '';
  const ruta = info.result.file_path;
  const res = UrlFetchApp.fetch(
    'https://api.telegram.org/file/bot' + getToken_() + '/' + ruta,
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) return '';
  const ext = ruta.indexOf('.') >= 0 ? ruta.split('.').pop() : 'jpg';
  const blob = res.getBlob().setName(nombre + '.' + ext);
  const carpeta = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  return carpeta.createFile(blob).getUrl();
}

// =====================================================================
// Consultar
// =====================================================================

function mostrarArreglos_(chatId, celular) {
  const filas = filasDeCelular_(celular);

  if (!filas.length) {
    enviar_(chatId, 'No encontré arreglos para el celular ' + celular + '.', TECLADO_MENU);
    return;
  }

  const total = filas.length;
  const mostrar = filas.slice(0, MAX_RESULTADOS);
  let encabezado = total === 1
    ? 'Encontré 1 arreglo para ' + celular + ':'
    : 'Encontré ' + total + ' arreglos para ' + celular + ':';
  if (total > MAX_RESULTADOS) {
    encabezado = 'Encontré ' + total + ' arreglos para ' + celular +
      '. Te muestro los ' + MAX_RESULTADOS + ' más recientes:';
  }
  enviar_(chatId, encabezado, TECLADO_MENU);

  mostrar.forEach(function (f) {
    const texto = textoArreglo_(f);
    const fotoTg = String(f[COL.FOTO_TG - 1] || '');
    if (!fotoTg) {
      enviar_(chatId, texto);
      return;
    }
    const sep = fotoTg.indexOf(':');
    const tipo = fotoTg.slice(0, sep);
    const id = fotoTg.slice(sep + 1);
    const caption = texto.slice(0, 1000);
    const r = tipo === 'document'
      ? tg_('sendDocument', { chat_id: chatId, document: id, caption: caption })
      : tg_('sendPhoto', { chat_id: chatId, photo: id, caption: caption });
    if (!r.ok) {
      enviar_(chatId, texto + '\nFoto: ' + (f[COL.FOTO - 1] || 'no disponible'));
    }
  });

  // Un solo QR por cliente, con todos sus arreglos
  tg_('sendMessage', {
    chat_id: chatId,
    text: '¿Querés el QR?',
    reply_markup: { inline_keyboard: [[
      { text: 'Sí', callback_data: 'qrc_' + celular },
      { text: 'No', callback_data: 'qr_no' },
    ]] },
  });
}

// Arreglos de un celular, del más nuevo al más viejo
function filasDeCelular_(celular) {
  const hoja = hojaArreglos_();
  const n = hoja.getLastRow();
  if (n < 2) return [];
  return hoja.getRange(2, 1, n - 1, ENCABEZADOS.length).getValues()
    .filter(function (f) { return normalizarCelular_(String(f[COL.CELULAR - 1])) === celular; })
    .sort(function (a, b) { return Number(b[COL.ID - 1]) - Number(a[COL.ID - 1]); });
}

// Bicis distintas de un cliente, las de los arreglos más recientes primero.
// Dos bicis son la misma si coinciden sin mirar mayúsculas ni espacios de más.
function bicisDistintas_(filas) {
  const claves = [];
  const bicis = [];
  filas.forEach(function (f) {
    const bici = String(f[COL.BICI - 1]).trim();
    const clave = bici.toLowerCase().replace(/\s+/g, ' ');
    if (!bici || claves.indexOf(clave) >= 0) return;
    claves.push(clave);
    bicis.push(bici);
  });
  return bicis.slice(0, MAX_BICIS);
}

function textoArreglo_(f) {
  return 'Arreglo #' + f[COL.ID - 1] + ' · ' + textoFecha_(f[COL.FECHA - 1]) + '\n' +
    'Cliente: ' + f[COL.CLIENTE - 1] + '\n' +
    'Bici: ' + f[COL.BICI - 1] + '\n' +
    'Trabajo: ' + f[COL.TRABAJO - 1] + '\n' +
    'Precio: ' + formatoPrecio_(f[COL.PRECIO - 1]);
}

// =====================================================================
// QR
// =====================================================================

function enviarQR_(chatId, celular, conf) {
  const filas = filasDeCelular_(celular);
  if (!filas.length) {
    enviar_(chatId, 'No encontré arreglos para el celular ' + celular + '.', TECLADO_MENU);
    return;
  }
  const contenido = textoQR_(conf.negocio, filas);

  // Va por POST: con varios arreglos el texto no entra en el largo máximo de una URL
  const res = UrlFetchApp.fetch('https://api.qrserver.com/v1/create-qr-code/', {
    method: 'post',
    payload: { size: '1000x1000', margin: '10', data: contenido },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    enviar_(chatId, 'No pude generar el QR en este momento. Probá de nuevo en un rato.', TECLADO_MENU);
    return;
  }
  const blob = res.getBlob().setName('qr_historial.png');
  tg_('sendPhoto',
    { chat_id: chatId, caption: 'QR del historial de arreglos', reply_markup: TECLADO_MENU },
    { campo: 'photo', blob: blob });
}

// El QR lleva solo negocio y, por arreglo, número, fecha, bici y trabajo:
// nada de celular, nombre del cliente ni precio
function textoQR_(negocio, filas) {
  const lineas = filas.map(function (f) {
    return '#' + f[COL.ID - 1] + ' ' + soloFecha_(f[COL.FECHA - 1]) + ' - ' +
      recortar_(f[COL.BICI - 1], 60) + ' - ' + recortar_(f[COL.TRABAJO - 1], 150);
  });
  let texto = recortar_(negocio, 100);
  for (let i = 0; i < lineas.length; i++) {
    // Si después de esta línea quedan más, tiene que entrar también el aviso
    const conLinea = texto + '\n' + lineas[i];
    const resto = lineas.length - i - 1;
    if ((conLinea + (resto ? '\n' + avisoAnteriores_(resto) : '')).length > MAX_QR) {
      return texto + '\n' + avisoAnteriores_(lineas.length - i);
    }
    texto = conLinea;
  }
  return texto;
}

function avisoAnteriores_(n) {
  return n === 1 ? '+1 arreglo anterior' : '+' + n + ' arreglos anteriores';
}

function recortar_(v, max) {
  const t = String(v).replace(/\s+/g, ' ').trim();
  return t.length > max ? t.slice(0, max - 1) + '…' : t;
}

// =====================================================================
// Utilidades
// =====================================================================

function hojaArreglos_() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_ARREGLOS);
  if (!hoja) throw new Error('Falta la pestaña Arreglos. Corré configurarTodo().');
  return hoja;
}

function siguienteId_(hoja) {
  const n = hoja.getLastRow();
  if (n < 2) return 1;
  const ids = hoja.getRange(2, COL.ID, n - 1, 1).getValues()
    .map(function (r) { return Number(r[0]) || 0; });
  return Math.max.apply(null, ids) + 1;
}

function leerConfig_() {
  const conf = { negocio: 'Bicicletería', usuarios: [] };
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  if (!hoja || hoja.getLastRow() < 2) return conf;
  hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(function (fila) {
    const clave = String(fila[0]).trim().toLowerCase();
    const valor = String(fila[1]).trim();
    if (clave === 'nombre del negocio' && valor) conf.negocio = valor;
    if (clave === 'usuarios autorizados') {
      conf.usuarios = valor.split(/[\s,;]+/).filter(Boolean);
    }
  });
  return conf;
}

function estaAutorizado_(userId, conf) {
  return conf.usuarios.indexOf(String(userId)) >= 0;
}

function extraerFoto_(msg) {
  if (msg.photo && msg.photo.length) {
    return { tipo: 'photo', fileId: msg.photo[msg.photo.length - 1].file_id };
  }
  if (msg.document && /^image\//.test(msg.document.mime_type || '')) {
    return { tipo: 'document', fileId: msg.document.file_id };
  }
  return null;
}

function normalizarCelular_(texto) {
  const digitos = String(texto).replace(/\D/g, '');
  return digitos.length >= 6 && digitos.length <= 15 ? digitos : null;
}

function normalizarPrecio_(texto) {
  let t = String(texto).replace(/[$\s]/g, '');
  if (!t) return null;
  if (/^\d{1,3}(\.\d{3})+(,\d{1,2})?$/.test(t)) {
    t = t.replace(/\./g, '').replace(',', '.'); // 15.000 o 15.000,50
  } else if (/^\d+(,\d{1,2})?$/.test(t)) {
    t = t.replace(',', '.');                    // 15000 o 15000,50
  } else if (!/^\d+(\.\d{1,2})?$/.test(t)) {
    return null;                                // 15000 o 15000.50
  }
  const n = Number(t);
  return isNaN(n) || n < 0 ? null : Math.round(n * 100) / 100;
}

function formatoPrecio_(v) {
  const n = Number(v);
  if (v === '' || isNaN(n)) return String(v);
  const partes = n.toFixed(2).split('.');
  const entero = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return '$' + entero + (partes[1] !== '00' ? ',' + partes[1] : '');
}

function textoFecha_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(v);
}

function soloFecha_(v) {
  return textoFecha_(v).split(' ')[0];
}

// Evita que un texto que empieza con = + - @ se tome como fórmula
function seguro_(texto) {
  const t = String(texto || '');
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function getEstado_(chatId) {
  const v = CacheService.getScriptCache().get('st_' + chatId);
  return v ? JSON.parse(v) : null;
}
function setEstado_(chatId, estado) {
  CacheService.getScriptCache().put('st_' + chatId, JSON.stringify(estado), 21600);
}
function borrarEstado_(chatId) {
  CacheService.getScriptCache().remove('st_' + chatId);
}

function enviar_(chatId, texto, teclado) {
  const datos = { chat_id: chatId, text: texto };
  if (teclado) datos.reply_markup = teclado;
  return tg_('sendMessage', datos);
}

function getToken_() {
  const token = (PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') || '').trim();
  if (!token) throw new Error('Falta TELEGRAM_TOKEN en las propiedades del script.');
  return token;
}

function tg_(metodo, datos, archivo) {
  const opciones = { method: 'post', muteHttpExceptions: true };
  if (archivo) {
    const payload = {};
    Object.keys(datos).forEach(function (k) {
      const v = datos[k];
      payload[k] = typeof v === 'object' ? JSON.stringify(v) : String(v);
    });
    payload[archivo.campo] = archivo.blob;
    opciones.payload = payload;
  } else {
    opciones.contentType = 'application/json';
    opciones.payload = JSON.stringify(datos);
  }
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + getToken_() + '/' + metodo, opciones);
  let json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    json = { ok: false, description: 'Respuesta inválida (' + res.getResponseCode() + ')' };
  }
  if (!json.ok) console.error(metodo + ' falló: ' + json.description);
  return json;
}
