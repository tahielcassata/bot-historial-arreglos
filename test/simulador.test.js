// Simulador de Apps Script + Telegram para probar Code.gs
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(require('path').join(__dirname, '..', 'src', 'Code.js'), 'utf8');

// ---------- Sheets mock ----------
function colToNum(c) { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
class Range {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  setValues(v) { v.forEach((row, i) => row.forEach((val, j) => this.sheet.set(this.r + i, this.c + j, val))); return this; }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sheet.get(this.r + i, this.c + j)); out.push(row); } return out; }
  setValue(v) { this.sheet.set(this.r, this.c, v); return this; }
  setFontWeight() { return this; } setBackground() { return this; } setNumberFormat() { return this; }
}
class Sheet {
  constructor(name) { this.name = name; this.data = []; this.hidden = []; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  set(r, c, v) { while (this.data.length < r) this.data.push([]); this.data[r - 1][c - 1] = v; }
  get(r, c) { const row = this.data[r - 1]; const v = row ? row[c - 1] : undefined; return v === undefined ? '' : v; }
  getLastRow() { for (let i = this.data.length; i > 0; i--) if (this.data[i - 1].some(v => v !== undefined && v !== '')) return i; return 0; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = a.match(/^([A-Z]+)(\d*):([A-Z]+)(\d*)$/) || a.match(/^([A-Z]+)(\d+)$/);
      if (m && m.length === 3) return new Range(this, +m[2], colToNum(m[1]), 1, 1);
      const r1 = m[2] ? +m[2] : 1, r2 = m[4] ? +m[4] : 1000;
      return new Range(this, r1, colToNum(m[1]), r2 - r1 + 1, colToNum(m[3]) - colToNum(m[1]) + 1);
    }
    return new Range(this, a, b, c || 1, d || 1);
  }
  appendRow(row) { const r = this.getLastRow() + 1; row.forEach((v, j) => this.set(r, j + 1, v)); }
  setFrozenRows() {} setColumnWidth() {} hideColumns(c) { this.hidden.push(c); }
}
const ss = {
  sheets: [new Sheet('Hoja 1')],
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; },
  getSheets() { return this.sheets; },
  insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; },
  getName() { return 'Historial de arreglos'; },
  getSpreadsheetTimeZone() { return 'America/Argentina/Buenos_Aires'; },
};

// ---------- Otros servicios ----------
const props = { TELEGRAM_TOKEN: 'TEST_TOKEN' };
const cache = {};
const calls = [];
const drive = { folders: {}, files: [] };
let qrFail = false;
let msgId = 1000;            // id que Telegram le pone a cada mensaje del bot
let ultimoConBotones = null; // último mensaje del bot con botones debajo

function resp(code, obj, blob) {
  return {
    getResponseCode: () => code,
    getContentText: () => JSON.stringify(obj),
    getBlob: () => blob || { name: '', setName(n) { this.name = n; return this; } },
  };
}

const ctx = {
  console: { log() {}, error: (...a) => ctx.__errors.push(a.join(' ')) },
  __errors: [],
  Logger: { log: m => ctx.__logs.push(m) }, __logs: [],
  SpreadsheetApp: { getActiveSpreadsheet: () => ss },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: {
    getUuid: () => '1234-5678-90ab',
    formatDate: (d, tz, f) => f.replace('dd', '12').replace('MM', '10').replace('yyyy', '2026').replace('HH', '14').replace('mm', '30').replace('ss', '00'),
  },
  DriveApp: {
    createFolder: n => { const id = 'F1'; drive.folders[id] = n; return { getId: () => id, getUrl: () => 'https://drive/F1' }; },
    getFolderById: id => { if (!drive.folders[id]) throw new Error('no folder'); return { getUrl: () => 'https://drive/' + id, createFile: b => { drive.files.push(b.name); return { getUrl: () => 'https://drive/file/' + b.name }; } }; },
  },
  ContentService: { createTextOutput: t => ({ text: t }) },
  HtmlService: { createHtmlOutput: t => ({ html: t }) },
  ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ABC/exec' }) },
  UrlFetchApp: {
    fetch(url, opts) {
      if (url.startsWith('https://api.qrserver.com')) {
        calls.push({ method: 'QRSERVER', url, payload: opts.payload });
        return qrFail ? resp(500, {}) : resp(200, {}, { name: '', setName(n) { this.name = n; return this; } });
      }
      if (url.startsWith('https://api.telegram.org/file/')) {
        calls.push({ method: 'DOWNLOAD', url });
        return resp(200, {}, { name: '', setName(n) { this.name = n; return this; } });
      }
      const m = url.match(/bot(.+?)\/(\w+)$/);
      assert.strictEqual(m[1], 'TEST_TOKEN');
      const method = m[2];
      const payload = opts.contentType === 'application/json' ? JSON.parse(opts.payload) : opts.payload;
      calls.push({ method, payload, multipart: !opts.contentType });
      if (method === 'getFile') return resp(200, { ok: true, result: { file_path: 'photos/file_1.jpg' } });
      if (method === 'getMe') return resp(200, { ok: true, result: { username: 'Historial_Control_bot' } });
      if (method === 'getWebhookInfo') return resp(200, { ok: true, result: { url: 'https://x/exec?secret=abc', pending_update_count: 0 } });
      if (method === 'sendMessage') {
        msgId++;
        if (payload.reply_markup && payload.reply_markup.inline_keyboard) ultimoConBotones = msgId;
        return resp(200, { ok: true, result: { message_id: msgId } });
      }
      return resp(200, { ok: true, result: {} });
    },
  },
};
vm.createContext(ctx);
vm.runInContext(code, ctx);

// ---------- Helpers ----------
let updId = 100;
const USER = 555;
function post(update, secret) {
  return ctx.doPost({ parameter: { secret: secret === undefined ? props.WEBHOOK_SECRET : secret }, postData: { contents: JSON.stringify(update) } });
}
function msg(text, extra = {}) {
  calls.length = 0;
  post({ update_id: ++updId, message: Object.assign({ message_id: updId, chat: { id: USER, type: 'private' }, from: { id: USER }, text }, extra) });
  return calls.slice();
}
// Toca un botón: por defecto, del último mensaje con botones; con `mensaje`, de uno viejo
function btn(data, mensaje = ultimoConBotones) {
  calls.length = 0;
  post({ update_id: ++updId, callback_query: { id: 'cq' + updId, from: { id: USER }, data, message: { message_id: mensaje, chat: { id: USER } } } });
  return calls.slice();
}
const botones = cs => cs.filter(c => c.method === 'sendMessage').slice(-1)[0].payload.reply_markup.inline_keyboard;
const texts = cs => cs.filter(c => c.method === 'sendMessage').map(c => c.payload.text);
const last = cs => texts(cs).slice(-1)[0];

// ---------- Pruebas ----------
let ok = 0;
function t(nombre, fn) { fn(); ok++; console.log('✓ ' + nombre); }

t('configurarTodo crea pestañas, columnas, carpeta y secreto', () => {
  ctx.configurarTodo();
  const arr = ss.getSheetByName('Arreglos');
  assert(arr, 'falta Arreglos');
  assert.deepStrictEqual(arr.data[0], ['ID', 'Fecha', 'Celular', 'Cliente', 'Bici', 'Trabajo', 'Precio', 'Foto', 'Foto Telegram']);
  assert(ss.getSheetByName('Configuración'));
  assert.strictEqual(ss.sheets.length, 2);
  assert(props.FOLDER_ID && props.WEBHOOK_SECRET);
  ctx.configurarTodo(); // segunda vez no rompe nada
  assert.strictEqual(ss.sheets.length, 2);
  assert.strictEqual(Object.keys(drive.folders).length, 1);
});

t('secreto incorrecto: se ignora', () => {
  calls.length = 0;
  post({ update_id: 1, message: { chat: { id: USER, type: 'private' }, from: { id: USER }, text: '/start' } }, 'malo');
  assert.strictEqual(calls.length, 0);
});

t('usuario no autorizado ve su ID y nada más', () => {
  const r = msg('/start');
  assert.match(last(r), /No estás autorizado/);
  assert.match(last(r), /555/);
});

t('se autoriza agregando el ID en Configuración', () => {
  ss.getSheetByName('Configuración').set(3, 2, '555');
  const r = msg('/start');
  assert.match(last(r), /Elegí una opción/);
  const kb = r.find(c => c.method === 'sendMessage').payload.reply_markup.keyboard;
  assert.deepStrictEqual(kb[0].map(b => b.text), ['Nuevo', 'Consultar']);
});

t('carga completa con foto', () => {
  assert.match(last(msg('Nuevo')), /Celular del cliente/);
  assert.match(last(msg('abc')), /no parece válido/);
  assert.match(last(msg('11 2345-6789')), /Nombre del cliente/);
  assert.match(last(msg('Juan Pérez')), /Bici/);
  assert.match(last(msg('Venzo rodado 29 negra')), /Trabajo/);
  assert.match(last(msg('=cambio de cadena')), /Precio/);
  assert.match(last(msg('quince')), /No entendí el precio/);
  const r = msg('$15.000');
  assert.strictEqual(r[0].payload.text, '¿Querés adjuntar la foto de la boleta?');
  assert.strictEqual(r[0].payload.reply_markup.inline_keyboard[0][0].callback_data, 'foto_si');
  assert.match(last(btn('foto_si')), /Mandá la foto/);
  assert.match(last(msg('hola')), /Mandá la foto/);
  const g = msg(undefined, { photo: [{ file_id: 'chica' }, { file_id: 'grande' }] });
  assert(g.some(c => c.method === 'DOWNLOAD'));
  assert.match(last(g), /Listo, guardado/);
  assert.match(last(g), /#1/);
  const fila = ss.getSheetByName('Arreglos').data[1];
  assert.strictEqual(fila[0], 1);
  assert.strictEqual(fila[2], '1123456789');
  assert.strictEqual(fila[5], "'=cambio de cadena", 'no se protegió la fórmula');
  assert.strictEqual(fila[6], 15000);
  assert.match(fila[7], /Boleta_1123456789/);
  assert.strictEqual(fila[8], 'photo:grande');
  assert.strictEqual(drive.files.length, 1);
  // guardar NO manda QR
  assert(!g.some(c => c.method === 'QRSERVER' || c.method === 'sendPhoto'));
});

t('celular existente: ofrece sus bicis; al elegir una no pide nombre ni bici', () => {
  msg('Nuevo');
  const r = msg('11 2345-6789');
  assert.strictEqual(r.length, 1); // una sola pregunta
  assert.strictEqual(r[0].payload.text, 'Este celular ya está guardado: Juan Pérez, 1 arreglo(s).\n¿En qué bici es el trabajo nuevo?');
  const kb = botones(r);
  assert.deepStrictEqual(kb.map(f => f.length), [1, 1, 1]); // un botón por fila
  assert.deepStrictEqual(kb.map(f => f[0].text), ['🚲 Venzo rodado 29 negra', '➕ Otra bici', '✖ Cancelar']);
  assert.deepStrictEqual(kb.map(f => f[0].callback_data), ['bici_0', 'bici_otra', 'agregar_no']);
  assert.match(last(msg('hola')), /Tocá una opción/);
  assert.match(last(btn('bici_0')), /^Trabajo realizado/); // salta nombre y bici
  assert.match(last(msg('Service completo')), /Precio/);
  msg('20000,50');
  const g = btn('foto_no');
  assert.match(last(g), /#2/);
  const fila = ss.getSheetByName('Arreglos').data[2];
  assert.strictEqual(fila[2], '1123456789');
  assert.strictEqual(fila[3], 'Juan Pérez');
  assert.strictEqual(fila[4], 'Venzo rodado 29 negra');
  assert.strictEqual(fila[6], 20000.5);
  assert.strictEqual(fila[8], '');
  // guardar NO manda QR
  assert(!g.some(c => c.method === 'QRSERVER' || c.method === 'sendPhoto'));
});

t('tocar el botón viejo de foto después de guardar no duplica', () => {
  const r = btn('foto_no');
  assert.match(last(r), /ya terminó/);
  assert.strictEqual(ss.getSheetByName('Arreglos').getLastRow(), 3);
});

t('celular existente: Cancelar corta la carga y vuelve al menú', () => {
  msg('Nuevo'); msg('1123456789');
  const r = btn('agregar_no');
  assert.match(last(r), /Cancelado/);
  const kb = r.filter(c => c.method === 'sendMessage').slice(-1)[0].payload.reply_markup.keyboard;
  assert.deepStrictEqual(kb[0].map(b => b.text), ['Nuevo', 'Consultar']);
  assert.match(last(msg('algo suelto')), /Elegí una opción/);
  assert.strictEqual(ss.getSheetByName('Arreglos').getLastRow(), 3);
});

t('botones de una carga cancelada y botón viejo "agregar_si": ya no están vigentes', () => {
  ['agregar_si', 'bici_0', 'bici_otra', 'agregar_no'].forEach(d => {
    assert.match(last(btn(d)), /ya terminó o se canceló/, d);
  });
  assert.match(last(msg('Venzo')), /Elegí una opción/); // no quedó ninguna carga abierta
  assert.strictEqual(ss.getSheetByName('Arreglos').getLastRow(), 3);
});

t('botón de bici de un mensaje viejo, con otra carga abierta: no elige nada', () => {
  msg('Nuevo'); msg('1123456789');
  const viejo = ultimoConBotones;
  msg('Cancelar');
  msg('Nuevo'); msg('1123456789');
  assert.notStrictEqual(ultimoConBotones, viejo);
  const r = btn('bici_0', viejo);
  assert.match(last(r), /ya no está vigente/);
  assert(!r.filter(c => c.method === 'sendMessage').slice(-1)[0].payload.reply_markup, 'no debe cambiar el teclado');
  assert.match(last(btn('agregar_si')), /ya no está vigente/);
  assert.match(last(btn('bici_7')), /ya no está vigente/); // índice que no existe
  assert.match(last(msg('hola')), /Tocá una opción/);      // la carga nueva sigue esperando
  assert.match(last(btn('agregar_no')), /Cancelado/);
  assert.strictEqual(ss.getSheetByName('Arreglos').getLastRow(), 3);
});

t('Cancelar a mitad de carga', () => {
  msg('Nuevo'); msg('1199999999');
  assert.match(last(msg('Cancelar')), /Cancelado/);
  assert.match(last(msg('algo suelto')), /Elegí una opción/);
});

t('Consultar: muestra historial nuevo→viejo, foto y una sola pregunta de QR', () => {
  assert.match(last(msg('Consultar')), /celular del cliente/);
  const r = msg('11-2345-6789');
  const tx = texts(r);
  assert.match(tx[0], /Encontré 2 arreglos/);
  assert.match(tx[1], /Arreglo #2/);          // el más nuevo primero, sin foto -> texto
  assert.match(tx[1], /Precio: \$20\.000,50/);
  const foto = r.find(c => c.method === 'sendPhoto');
  assert.strictEqual(foto.payload.photo, 'grande'); // el #1 con su foto
  assert.match(foto.payload.caption, /Arreglo #1/);
  assert.match(foto.payload.caption, /\$15\.000$/);
  const pregunta = r.filter(c => c.method === 'sendMessage').slice(-1)[0];
  assert.strictEqual(pregunta.payload.text, '¿Querés el QR?');
  const kb = pregunta.payload.reply_markup.inline_keyboard;
  assert.strictEqual(kb.length, 1); // ya no hay un botón por arreglo
  assert.deepStrictEqual(kb[0].map(b => b.text), ['Sí', 'No']);
  assert.deepStrictEqual(kb[0].map(b => b.callback_data), ['qrc_1123456789', 'qr_no']);
});

t('QR unificado: todos los arreglos nuevo→viejo; sin celular, nombre ni precio', () => {
  const r = btn('qrc_1123456789');
  const qs = r.filter(c => c.method === 'QRSERVER');
  assert.strictEqual(qs.length, 1); // un solo QR
  const lineas = qs[0].payload.data.split('\n');
  assert.strictEqual(lineas.length, 3);
  assert.strictEqual(lineas[0], 'Bicicletería');
  assert.strictEqual(lineas[1], '#2 12/10/2026 - Venzo rodado 29 negra - Service completo');
  assert.match(lineas[2], /^#1 12\/10\/2026 - Venzo rodado 29 negra - .*cambio de cadena$/);
  const todo = qs[0].url + JSON.stringify(qs[0].payload);
  assert(!/1123456789/.test(todo), 'el QR tiene el celular');
  assert(!/Juan|Pérez/.test(todo), 'el QR tiene el nombre del cliente');
  assert(!/15000|15\.000|20000|20\.000/.test(todo), 'el QR tiene el precio');
  const envio = r.find(c => c.method === 'sendPhoto');
  assert(envio.multipart && envio.payload.photo.name === 'qr_historial.png');
  assert(!/1123456789|Juan/.test(envio.payload.caption));
  assert(r.some(c => c.method === 'answerCallbackQuery'));
  assert(r.some(c => c.method === 'editMessageReplyMarkup'));
});

t('QR: si el servicio falla, avisa', () => {
  qrFail = true;
  assert.match(last(btn('qrc_1123456789')), /No pude generar el QR/);
  qrFail = false;
});

t('botón viejo "qr_N": pide consultar de nuevo y no genera QR', () => {
  const r = btn('qr_1');
  assert.match(last(r), /Consultá de nuevo/);
  assert(!r.some(c => c.method === 'QRSERVER' || c.method === 'sendPhoto'));
});

t('QR: con muchos arreglos respeta el límite y avisa cuántos quedaron afuera', () => {
  const hoja = ss.getSheetByName('Arreglos');
  for (let i = 101; i <= 130; i++) {
    hoja.appendRow([i, '0' + (i % 9 + 1) + '/09/2026 10:00', '1177770000', 'Carlos Secreto',
      'Bici ' + 'b'.repeat(100), 'Trabajo ' + i + ' ' + 't'.repeat(300), 98765, '', '']);
  }
  const r = btn('qrc_1177770000');
  const data = r.find(c => c.method === 'QRSERVER').payload.data;
  assert(data.length <= 1200, 'el QR pasa el límite: ' + data.length);
  const lineas = data.split('\n');
  const arreglos = lineas.filter(l => l.startsWith('#'));
  assert(arreglos.length >= 2 && arreglos.length < 30);
  assert.match(arreglos[0], /^#130 /);          // entran los más recientes
  assert.match(arreglos[1], /^#129 /);
  arreglos.forEach(l => assert(l.length <= 235, 'trabajo o bici sin recortar'));
  const m = lineas[lineas.length - 1].match(/^\+(\d+) arreglos anteriores$/);
  assert(m, 'falta el aviso de arreglos anteriores');
  assert.strictEqual(arreglos.length + Number(m[1]), 30);
  assert(!/1177770000|Carlos|98765/.test(data));
});

t('Consultar con un solo arreglo: botones Sí / No', () => {
  msg('Nuevo');
  assert.match(last(msg('1155554444')), /Nombre del cliente/); // celular nuevo: pide el nombre como siempre
  msg('Ana'); msg('Trek'); msg('Frenos'); msg('8000'); btn('foto_no');
  msg('Consultar');
  const r = msg('1155554444');
  const kb = r.filter(c => c.method === 'sendMessage').slice(-1)[0].payload.reply_markup.inline_keyboard;
  assert.deepStrictEqual(kb[0].map(b => b.text), ['Sí', 'No']);
  const no = btn('qr_no');
  assert.match(last(no), /Listo/);
  assert(!no.some(c => c.method === 'QRSERVER'));
});

const BICI_LARGA = 'Scott Aspect 950 azul con portaequipaje y guardabarros';
const ultimaFila = () => { const h = ss.getSheetByName('Arreglos'); return h.data[h.getLastRow() - 1]; };

t('celular existente: "Otra bici" pide la bici y sigue como siempre', () => {
  msg('Nuevo');
  const r = msg('1155554444');
  assert.deepStrictEqual(botones(r).map(f => f[0].text), ['🚲 Trek', '➕ Otra bici', '✖ Cancelar']);
  assert.match(last(btn('bici_otra')), /^Bici \(marca/);
  assert.match(last(msg(BICI_LARGA)), /Trabajo/);
  msg('Parche'); msg('3000');
  assert.match(last(btn('foto_no')), /Listo, guardado/);
  assert.strictEqual(ultimaFila()[3], 'Ana'); // el nombre no se volvió a pedir
  assert.strictEqual(ultimaFila()[4], BICI_LARGA);
});

t('bici de nombre largo: el botón va recortado, pero se guarda el texto completo', () => {
  msg('Nuevo');
  const kb = botones(msg('1155554444'));
  assert.deepStrictEqual(kb.map(f => f[0].callback_data), ['bici_0', 'bici_1', 'bici_otra', 'agregar_no']);
  assert.strictEqual(kb[0][0].text, '🚲 Scott Aspect 950 azul con por…'); // la más reciente primero
  assert.strictEqual(kb[1][0].text, '🚲 Trek');
  kb.forEach(f => assert(Buffer.byteLength(f[0].callback_data) <= 64));
  const antes = ss.getSheetByName('Arreglos').getLastRow();
  assert.match(last(btn('bici_0')), /^Trabajo realizado/);
  msg('Centrado de rueda'); msg('5000');
  const g = btn('foto_no');
  assert.match(last(g), /Listo, guardado/);
  assert.strictEqual(ss.getSheetByName('Arreglos').getLastRow(), antes + 1);
  assert.strictEqual(ultimaFila()[4], BICI_LARGA);
  assert.strictEqual(ultimaFila()[5], 'Centrado de rueda');
  assert(!g.some(c => c.method === 'QRSERVER' || c.method === 'sendPhoto')); // guardar NO manda QR
});

t('bicis repetidas con distinta mayúscula o espacios: aparecen una sola vez', () => {
  const hoja = ss.getSheetByName('Arreglos');
  [[201, 'venzo  29 NEGRA'], [202, 'Trek'], [203, 'Venzo 29 negra'], [204, 'TREK']].forEach(([id, bici]) => {
    hoja.appendRow([id, '01/09/2026 10:00', '1166660000', 'Beto', bici, 'Service', 1000, '', '']);
  });
  msg('Nuevo');
  const r = msg('1166660000');
  assert.match(last(r), /Beto, 4 arreglo\(s\)/);
  // se muestra el texto del arreglo más reciente de cada bici
  assert.deepStrictEqual(botones(r).map(f => f[0].text), ['🚲 TREK', '🚲 Venzo 29 negra', '➕ Otra bici', '✖ Cancelar']);
  btn('agregar_no');
});

t('cliente con más de 4 bicis: ofrece solo las 4 más recientes', () => {
  const hoja = ss.getSheetByName('Arreglos');
  ['A', 'B', 'C', 'D', 'E', 'F'].forEach((letra, i) => {
    hoja.appendRow([301 + i, '01/09/2026 10:00', '1188880000', 'Caro', 'Bici ' + letra, 'Service', 1000, '', '']);
  });
  msg('Nuevo');
  const r = msg('1188880000');
  assert.deepStrictEqual(botones(r).map(f => f[0].text),
    ['🚲 Bici F', '🚲 Bici E', '🚲 Bici D', '🚲 Bici C', '➕ Otra bici', '✖ Cancelar']);
  assert.match(last(btn('bici_3')), /^Trabajo realizado/);
  msg('Frenos'); msg('2000'); btn('foto_no');
  assert.strictEqual(ultimaFila()[4], 'Bici C');
  assert.strictEqual(ultimaFila()[3], 'Caro');
});

t('Consultar un celular sin arreglos', () => {
  msg('Consultar');
  assert.match(last(msg('1100000000')), /No encontré arreglos/);
});

t('mensaje repetido por Telegram se procesa una sola vez', () => {
  calls.length = 0;
  const u = { update_id: 9999, message: { chat: { id: USER, type: 'private' }, from: { id: USER }, text: '/start' } };
  post(u); const n = calls.length; post(u);
  assert.strictEqual(calls.length, n);
});

t('foto sin estar cargando: no se guarda nada', () => {
  const antes = ss.getSheetByName('Arreglos').getLastRow();
  const r = msg(undefined, { photo: [{ file_id: 'x' }] });
  assert.match(last(r), /Elegí una opción/);
  assert.strictEqual(ss.getSheetByName('Arreglos').getLastRow(), antes);
});

t('botón tocado por alguien no autorizado', () => {
  calls.length = 0;
  post({ update_id: ++updId, callback_query: { id: 'z', from: { id: 777 }, data: 'qr_1', message: { message_id: 1, chat: { id: 777 } } } });
  assert.strictEqual(calls.length, 1);
  assert.strictEqual(calls[0].payload.text, 'No autorizado');
});

t('conectarTelegram arma la URL con el secreto', () => {
  calls.length = 0;
  ctx.conectarTelegram();
  const c = calls.find(x => x.method === 'setWebhook');
  assert.strictEqual(c.payload.url, 'https://script.google.com/macros/s/ABC/exec?secret=' + props.WEBHOOK_SECRET);
  assert.deepStrictEqual(c.payload.allowed_updates, ['message', 'callback_query']);
});

t('estadoTelegram no muestra el secreto', () => {
  ctx.__logs.length = 0;
  ctx.estadoTelegram();
  assert(!ctx.__logs.join(' ').includes('secret=abc'));
  assert(ctx.__logs[0].includes('@Historial_Control_bot'));
});

t('el token nunca aparece en los logs', () => {
  assert(!JSON.stringify(ctx.__logs).includes('TEST_TOKEN'));
  assert(!JSON.stringify(ctx.__errors).includes('TEST_TOKEN'));
});

console.log(`\n${ok} pruebas OK`);
