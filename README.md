# Bot Historial de arreglos

Bot de Telegram para registrar y consultar los arreglos de una bicicletería.
Cada arreglo se guarda en una hoja de Google Sheets, con foto opcional de la boleta
en Google Drive, y se puede generar un QR con el detalle del trabajo.

Corre en **Google Apps Script**: no necesita servidor ni una compu prendida, y es gratis.

## Qué hace

- **Nuevo:** pide celular, cliente, bici, trabajo y precio, uno por uno. La fecha se pone sola.
  Al final pregunta si querés adjuntar la foto de la boleta (Sí / No).
- **Cliente que ya existe:** si el celular ya está guardado, no vuelve a pedir el nombre y ofrece las
  bicis de ese cliente como botones (hasta 4, las más recientes primero), además de "Otra bici" y
  "Cancelar". Al elegir una, pasa directo al trabajo.
- **Consultar:** pide el celular y muestra los arreglos de ese cliente, del más nuevo al más viejo,
  con su foto si la tiene.
- **QR:** al consultar, ofrece generar un único QR con todo el historial del cliente: nombre del negocio
  y una línea por arreglo con número, fecha, bici y trabajo. No lleva celular, nombre del cliente ni
  precio, porque cualquiera que lo escanee puede leerlo. Si el historial es muy largo, entran los
  arreglos más recientes y se avisa cuántos quedaron afuera.
- Botones dentro del chat: no hace falta escribir comandos.
- Solo responde a los usuarios autorizados en la hoja.

## Cómo funciona

```
Telegram  ──webhook──▶  Apps Script (doPost)  ──▶  Google Sheets (pestaña Arreglos)
                               │
                               ├──▶  Google Drive (carpeta de boletas)
                               └──▶  api.qrserver.com (imagen del QR)
```

- La conversación de cada carga se guarda temporalmente en `CacheService`.
- Los IDs de arreglo son correlativos y se asignan con `LockService` para que no se repitan.
- Telegram puede reenviar el mismo mensaje: se procesa una sola vez (por `update_id`).
- El webhook lleva una clave secreta en la URL; los pedidos sin esa clave se ignoran.
- Los botones llevan solo un índice, no el texto: la lista de opciones queda en la caché junto con el
  ID del mensaje, así un botón de una pregunta vieja no puede elegir algo de otra lista.
- Los textos que escribe el usuario se guardan protegidos para que la hoja no los tome como fórmulas.

## Estructura

```
bot-historial-arreglos/
├── src/
│   ├── Code.js            # todo el código del bot
│   └── appsscript.json    # configuración del proyecto de Apps Script
├── test/
│   └── simulador.test.js  # simula Telegram, Sheets y Drive para probar sin internet
├── .clasp.json.example    # modelo para vincular la carpeta con un proyecto de Apps Script
└── package.json
```

Las pruebas no necesitan instalar nada: con Node.js alcanza (`npm test`).

## Configuración (en Google)

**Hoja "Historial de arreglos", pestaña Configuración:**
- `Nombre del negocio`: aparece arriba en el QR.
- `Usuarios autorizados`: IDs de Telegram separados por coma. Si alguien no autorizado le escribe
  al bot, el bot le muestra su ID.

**Apps Script → Configuración del proyecto → Propiedades del script:**
- `TELEGRAM_TOKEN`: el token de @BotFather. Nunca va en el código ni en GitHub.
- `FOLDER_ID`, `WEBHOOK_SECRET`: los crea `configurarTodo()`.
- `WEBAPP_URL`: la URL `/exec` de la implementación.

## Trabajar desde VS Code

Requisitos: [Node.js](https://nodejs.org) y [clasp](https://github.com/google/clasp), la herramienta oficial de Google.

```bash
npm install -g @google/clasp
clasp login          # abre el navegador para iniciar sesión con tu cuenta de Google
```

La primera vez, activá la "Google Apps Script API" en https://script.google.com/home/usersettings.

Después copiá `.clasp.json.example` como `.clasp.json` y poné el ID de tu proyecto de Apps Script
(está en Configuración del proyecto → ID de la secuencia de comandos). Ese archivo no se sube al repositorio.

Flujo de trabajo:

```bash
npm test             # corre las pruebas con el simulador
npm run subir        # sube src/ al proyecto de Apps Script (clasp push)
npm run bajar        # trae los cambios hechos en el editor web (clasp pull)
```

**Importante:** subir el código no actualiza el bot. Después de `npm run subir`, en Apps Script:
**Implementar → Gestionar implementaciones → ✏️ → Versión: Versión nueva → Implementar.**
Así la URL no cambia y no hay que reconectar Telegram.

Lo mismo se puede hacer desde la terminal, publicando sobre la implementación que ya existe:

```bash
clasp list-deployments                                   # muestra el ID de la implementación
clasp create-deployment --deploymentId <ID> --description "qué cambió"
```

## Funciones para correr a mano (desde el editor de Apps Script)

| Función | Para qué |
|---|---|
| `configurarTodo()` | Crea las pestañas, columnas y la carpeta de boletas. Se puede correr más de una vez. |
| `conectarTelegram()` | Conecta el bot con la aplicación web. También vacía mensajes trabados. |
| `estadoTelegram()` | Muestra si la conexión anda y el último error de Telegram. |

## Problemas conocidos y soluciones

- **El bot no responde y `estadoTelegram()` dice `302 Found`:** pasa si `doPost` responde con
  `ContentService`, que hace una redirección que Telegram toma como error. Por eso se responde con
  `HtmlService`. Después de corregirlo, correr `conectarTelegram()` para vaciar la cola.
- **Cambié el código y el bot sigue igual:** falta publicar una versión nueva (ver arriba).

## Tecnologías

Google Apps Script (JavaScript V8) · Google Sheets · Google Drive · Telegram Bot API · Node.js (pruebas) · clasp
