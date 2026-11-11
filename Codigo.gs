
/**
 * ZEBRA PRICE DESK - ANALYTICS API
 * Intcomex | Zebra
 *
 * Validación de usuarios:
 * Columna A: Username
 * Columna B: Nombre
 * Columna C: Apellido
 * Columna D: Rol
 *
 * Eventos permitidos:
 * - quote_session
 * - quote_item
 *
 * IMPORTANTE:
 * La autenticación por username no demuestra la identidad
 * de quien utiliza la aplicación. La comprobación de roles
 * no debe considerarse seguridad robusta.
 */

// =====================================================
// CONFIGURACIÓN
// =====================================================

const USERS_SPREADSHEET_ID =
  '1z_-VWCunuX4fNa6FEBQ2IT4bmtWAdC6VVKSc2ztL7IE';

const ANALYTICS_FOLDER_ID =
  '1DegMUDB_Wkhmh9bxV7DpE3wF_P5lgXLc';

const ANALYTICS_FILE_NAME =
  'Zebra Price Desk - Dashboard Analytics';

const ANALYTICS_SPREADSHEET_PROPERTY =
  'ZPD_ANALYTICS_SPREADSHEET_ID';

const ANALYTICS_TAB_NAME = 'Eventos';

// La hoja de usuarios utiliza su primera pestaña.
const USERS_TAB_INDEX = 0;

const MAX_EVENTS_PER_REQUEST = 500;
const MAX_READ_EVENTS = 20000;

const ANALYTICS_HEADERS = [
  'timestamp',
  'eventType',
  'quoteId',
  'username',
  'userName',
  'sku',
  'skuCount',
  'partnerName',
  'partnerId',
  'partnerLevel',
  'country',
  'region',
  'basePrice',
  'finalPrice',
  'discountApplied',
  'promoCode',
  'accessAllowed',
  'accessMessage',
  'status'
];


// =====================================================
// VALIDACIÓN DE USUARIOS
// =====================================================

/**
 * Busca el username en la hoja de usuarios.
 * Solamente los usuarios registrados pueden ingresar.
 */
function validateUser_(username) {

  const cleanUsername = normalizeUsername_(username);

  if (!cleanUsername) {
    throw new Error('Debes proporcionar un username.');
  }

  const spreadsheet = SpreadsheetApp.openById(
    USERS_SPREADSHEET_ID
  );

  const usersSheet = spreadsheet.getSheets()[USERS_TAB_INDEX];

  if (!usersSheet) {
    throw new Error('No se encontró la hoja de usuarios.');
  }

  const lastRow = usersSheet.getLastRow();

  if (lastRow < 1) {
    throw new Error('La hoja de usuarios está vacía.');
  }

  // Leer toda la información de las columnas A, B, C y D.
  const rows = usersSheet
    .getRange(1, 1, lastRow, 4)
    .getDisplayValues();

  // Buscar el username en cualquier fila de la columna A.
  for (let i = 0; i < rows.length; i++) {

    const rowUsername = normalizeUsername_(rows[i][0]);

    if (!rowUsername) continue;

    if (rowUsername === cleanUsername) {

      const firstName = String(rows[i][1] || '').trim();
      const lastName = String(rows[i][2] || '').trim();
      const role = normalizeRole_(rows[i][3]);

      return {
        username: String(rows[i][0]).trim(),
        name: `${firstName} ${lastName}`.trim() || String(rows[i][0]).trim(),
        role: role,
        isAdmin: role === 'ADMINISTRADOR'
      };
    }
  }

  throw new Error('Usuario no registrado o sin autorización.');
}


/**
 * Normaliza usernames para comparar sin distinguir mayúsculas.
 */
function normalizeUsername_(value) {
  return String(value || '').trim().toUpperCase();
}


/**
 * Normaliza roles para aceptar Administrador,
 * ADMINISTRADOR o administrador.
 */
function normalizeRole_(value) {

  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase();
}


// =====================================================
// REGISTRO DE EVENTOS
// =====================================================

/**
 * Recibe eventos enviados desde Zebra Price Desk.
 *
 * Formato esperado:
 *
 * {
 *   "username": "USUARIO",
 *   "events": [
 *     {
 *       "eventType": "quote_session",
 *       "quoteId": "12345"
 *     }
 *   ]
 * }
 */

function doPost(e) {
  try {
    let payload = {};

    const rawBody = (e && e.postData && e.postData.contents) || '{}';

    // Leer solicitudes JSON o formularios HTML.
    try {
      payload = JSON.parse(rawBody);
    } catch (parseError) {
      const formPayload = e && e.parameter && e.parameter.payload;

      if (!formPayload) {
        throw parseError;
      }

      payload = JSON.parse(formPayload);
    }

    // Enrutar las operaciones administrativas de promociones.
    if (String(payload.action || '').indexOf('promotion_') === 0) {
      const response = handlePromotionAction_(payload);

      // Guardar temporalmente el resultado para recuperarlo desde el navegador.
      const requestId = String(payload.requestId || '');

      if (/^[a-f0-9-]{20,100}$/i.test(requestId)) {
        CacheService.getScriptCache().put(
          'ZPD_PROMO_RESULT_' + requestId,
          response.getContent(),
          120
        );
      }

      return response;
    }

    // Validar que el usuario exista.
    const user = validateUser_(payload.username);

    // Validar que los eventos tengan el formato esperado.
    const events = Array.isArray(payload.events)
      ? payload.events
      : [];

    const sheet = getAnalyticsSheet_();
    const rows = [];

    events.forEach(event => {
      if (!event) return;

      const eventType = String(event.eventType || '');

      if (!['quote_session', 'quote_item'].includes(eventType)) {
        return;
      }

      rows.push([
        String(event.timestamp || new Date().toISOString()),
        eventType,
        safeText_(event.quoteId),
        user.username,
        user.name,
        safeText_(event.sku),
        numericOrBlank_(event.skuCount),
        safeText_(event.partnerName),
        safeText_(event.partnerId),
        safeText_(event.partnerLevel),
        safeText_(event.country),
        safeText_(event.region),
        numericOrBlank_(event.basePrice),
        numericOrBlank_(event.finalPrice),
        numericOrBlank_(event.discountApplied),
        safeText_(event.promoCode),

        event.accessAllowed === true
          ? 'TRUE'
          : event.accessAllowed === false
            ? 'FALSE'
            : '',

        safeText_(event.accessMessage),
        safeText_(event.status)
      ]);
    });

    if (rows.length) {
      const startRow = sheet.getLastRow() + 1;

      sheet.getRange(
        startRow,
        1,
        rows.length,
        ANALYTICS_HEADERS.length
      ).setValues(rows);
    }

    return json_({
      ok: true,
      saved: rows.length,
      username: user.username
    });

  } catch (error) {
    console.error(error);

    return json_({
      ok: false,
      error: String(error.message || error)
    });
  }
}


// =====================================================
// CONSULTA DEL DASHBOARD
// =====================================================

/**
 * Consulta los eventos del dashboard.
 *
 * Parámetros:
 * action=read
 * username=USERNAME
 *
 * También admite JSONP mediante callback.
 *
 * IMPORTANTE:
 * El username puede falsificarse porque no hay contraseña
 * ni una identidad autenticada.
 */
function doGet(e) {

  const params = (e && e.parameter) || {};

  const callback = String(params.callback || '');

  try {

    // Recuperar el resultado temporal de una operación de promociones.
if (params.action === 'promotion_result') {
  const requestId = String(params.requestId || '');

  if (!/^[a-f0-9-]{20,100}$/i.test(requestId)) {
    throw new Error('Identificador de solicitud no válido.');
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'ZPD_PROMO_RESULT_' + requestId;
  const cachedResult = cache.get(cacheKey);

  if (!cachedResult) {
    throw new Error('Resultado no disponible o solicitud expirada.');
  }

  cache.remove(cacheKey);

  return jsonp_(JSON.parse(cachedResult), callback);
}

    if (params.action !== 'read') {
      throw new Error('Acción no válida.');
    }

    const user = validateUser_(params.username);

    if (!user.isAdmin) {
      throw new Error('Acceso administrativo no autorizado.');
    }

    const sheet = getAnalyticsSheet_();

    const values = sheet.getDataRange().getValues();

    const headers = values.shift() || ANALYTICS_HEADERS;

    const events = values
      .filter(row => row.some(value => value !== ''))
      .map(row => {

        const item = {};

        headers.forEach((header, index) => {

          const value = row[index];

          item[header] = value instanceof Date
            ? value.toISOString()
            : value;

        });

        item.skuCount = Number(item.skuCount) || 0;

        item.basePrice = item.basePrice === ''
          ? ''
          : Number(item.basePrice);

        item.finalPrice = item.finalPrice === ''
          ? ''
          : Number(item.finalPrice);

        item.discountApplied = item.discountApplied === ''
          ? 0
          : Number(item.discountApplied);

        item.accessAllowed =
          item.accessAllowed === 'TRUE'
            ? true
            : item.accessAllowed === 'FALSE'
              ? false
              : null;

        return item;

      });

    return jsonp_({
      ok: true,
      username: user.username,
      events: events.slice(-MAX_READ_EVENTS)
    }, callback);

  } catch (error) {

    return jsonp_({
      ok: false,
      error: String(error.message || error),
      events: []
    }, callback);

  }

}


// =====================================================
// CREACIÓN Y CONFIGURACIÓN DE LA BASE CENTRAL
// =====================================================

/**
 * Ejecutar una sola vez desde el editor.
 *
 * Si ya existe la hoja central, la reutiliza.
 * Si no existe, la crea dentro de la carpeta configurada.
 */
function setupAnalyticsStorage() {

  const sheet = getAnalyticsSheet_();

  const spreadsheet = sheet.getParent();

  const url = spreadsheet.getUrl();

  Logger.log('Base central creada/lista: ' + url);

  return url;

}


/**
 * Obtiene la hoja central de analítica.
 */
function getAnalyticsSheet_() {

  const properties = PropertiesService.getScriptProperties();

  let spreadsheetId = properties.getProperty(
    ANALYTICS_SPREADSHEET_PROPERTY
  );

  let spreadsheet = null;

  if (spreadsheetId) {

    try {

      spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    } catch (error) {

      properties.deleteProperty(
        ANALYTICS_SPREADSHEET_PROPERTY
      );

      spreadsheetId = '';

    }

  }

  if (!spreadsheet) {

    const folder = DriveApp.getFolderById(
      ANALYTICS_FOLDER_ID
    );

    spreadsheet = SpreadsheetApp.create(
      ANALYTICS_FILE_NAME
    );

    const file = DriveApp.getFileById(
      spreadsheet.getId()
    );

    file.moveTo(folder);

    properties.setProperty(
      ANALYTICS_SPREADSHEET_PROPERTY,
      spreadsheet.getId()
    );

  }

  let sheet = spreadsheet.getSheetByName(
    ANALYTICS_TAB_NAME
  );

  if (!sheet) {

    sheet = spreadsheet.insertSheet(
      ANALYTICS_TAB_NAME
    );

  }

  if (sheet.getLastRow() === 0) {

    sheet.getRange(
      1,
      1,
      1,
      ANALYTICS_HEADERS.length
    ).setValues([ANALYTICS_HEADERS]);

    sheet.setFrozenRows(1);

    sheet.getRange(
      1,
      1,
      1,
      ANALYTICS_HEADERS.length
    ).setFontWeight('bold');

  }

  return sheet;

}


// =====================================================
// FUNCIONES AUXILIARES
// =====================================================

/**
 * Evita interpretar texto recibido como fórmula
 * al guardarlo en Google Sheets.
 */
function safeText_(value) {

  const text = String(
    value == null ? '' : value
  ).slice(0, 500);

  return /^[=+\-@]/.test(text)
    ? "'" + text
    : text;

}


/**
 * Convierte valores numéricos o deja la celda vacía.
 */
function numericOrBlank_(value) {

  if (
    value === '' ||
    value === null ||
    value === undefined ||
    !Number.isFinite(Number(value))
  ) {
    return '';
  }

  return Number(value);

}


/**
 * Respuesta JSON.
 */
function json_(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

}


/**
 * Respuesta JSONP para consultas desde el navegador.
 */
function jsonp_(data, callback) {

  const safeCallback =
    /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)
      ? callback
      : '';

  if (!safeCallback) {
    return json_(data);
  }

  return ContentService
    .createTextOutput(
      `${safeCallback}(${JSON.stringify(data)});`
    )
    .setMimeType(ContentService.MimeType.JAVASCRIPT);

}


// =====================================================
// ADMINISTRACIÓN DE PROMOCIONES
// =====================================================

const PROMOTIONS_FILE_NAME = 'Zebra Price Desk - Promociones';
const PROMOTIONS_FILE_PROPERTY = 'ZPD_PROMOTIONS_SPREADSHEET_ID';
const PROMOTIONS_TAB_NAME = 'Promociones';
const PROMOTIONS_HEADERS = [
  'ID',
  'Codigo',
  'Region',
  'TipoCoincidencia',
  'PatronSKU',
  'TipoDescuento',
  'ValorDescuento',
  'FechaInicio',
  'FechaFin',
  'Activa',
  'ExcluirDNP',
  'Observaciones'
];

/**
 * Crea una hoja independiente con la matriz inicial de promociones.
 * Ejecutar manualmente una sola vez desde el editor de Apps Script.
 * Si ya existe, devuelve la hoja creada anteriormente sin duplicarla.
 *
 * Las fechas se dejan vacías porque deben confirmarse según la vigencia
 * comercial oficial de cada promoción antes de activar la validación temporal.
 */
function setupPromotionsStorage() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty(PROMOTIONS_FILE_PROPERTY);

  if (existingId) {
    const existing = SpreadsheetApp.openById(existingId);
    const existingSheet = existing.getSheetByName(PROMOTIONS_TAB_NAME);
    if (existingSheet) {
      Logger.log('La hoja de promociones ya existe: ' + existing.getUrl());
      Logger.log('Spreadsheet ID: ' + existing.getId());
      Logger.log('GID: ' + existingSheet.getSheetId());
      return {
        url: existing.getUrl(),
        spreadsheetId: existing.getId(),
        gid: existingSheet.getSheetId(),
        created: false
      };
    }
  }

  const spreadsheet = SpreadsheetApp.create(PROMOTIONS_FILE_NAME);
  const sheet = spreadsheet.getSheets()[0];
  sheet.setName(PROMOTIONS_TAB_NAME);

  const rows = [];
  const add = (code, region, match, sku, type, value, excludeDnp, notes) => {
    rows.push([
      Utilities.getUuid(),
      code,
      region,
      match,
      sku,
      type,
      value,
      '',
      '',
      true,
      Boolean(excludeDnp),
      notes || ''
    ]);
  };

  add('MC34N2605', 'NOLA', 'PREFIJO', 'MC34', 'FIJO', 100, false, 'Cualquier SKU MC34; no aplica a Colombia.');
  [
    'MC3401-0G1D43SS-A6',
    'MC3401-0G1J54SS-A6',
    'MC3401-0G1M54SS-A6'
  ].forEach(sku => add('MC34COL2609', 'COLOMBIA', 'EXACTO', sku, 'FIJO', 165, false));
  [
    'MC3401-0G1D43SS-A6',
    'MC3401-0G1J54SS-A6'
  ].forEach(sku => add('MC34S2609', 'SOLA', 'EXACTO', sku, 'FIJO', 100, false));
  [
    'WLMT0-T22',
    'WCMTA-T27',
    'WCMTB-T27'
  ].forEach(prefix => add('TC2LCOL2601', 'COLOMBIA', 'PREFIJO', prefix, 'FIJO', 30, false));
  add('ZD220T3HWA26', 'COLOMBIA,NOLA', 'PREFIJO', 'ZD220', 'FIJO', 10, false);
  [
    'ZD4A04',
    'ZD4AH4'
  ].forEach(prefix => add('ZD421T3HWA26', 'COLOMBIA,NOLA', 'PREFIJO', prefix, 'FIJO', 15, false));
  add('ZT231T3HWA26', 'COLOMBIA,NOLA', 'PREFIJO', 'ZT231', 'FIJO', 30, false);
  add('ZT411T3HWA26', 'NOLA', 'PREFIJO', 'ZT411', 'FIJO', 45, true, 'Excluir cualquier SKU presente en DNP.');
  add('ZT411COL2601', 'COLOMBIA', 'EXACTO', 'ZT41142-T010000Z', 'FIJO', 52, false);
  [
    'ZC32-000CQ00LA00',
    'ZC31-000CQ00LA00'
  ].forEach(sku => add('ZC300QNSMB2409', 'COLOMBIA,SOLA,NOLA', 'EXACTO', sku, 'FIJO', 25, false));
  add('DS2208NS2601', 'COLOMBIA,SOLA,NOLA', 'PREFIJO', 'DS2208', 'FIJO', 10, false);
  add('DS2278NS2601', 'COLOMBIA,SOLA,NOLA', 'PREFIJO', 'DS2278', 'FIJO', 16.20, false);
  add('PVCZCNSM2609', 'COLOMBIA,SOLA,NOLA', 'EXACTO', '104523-111', 'PORCENTAJE', 3, false);
  [
    'DS4608-SR00007ZZWW',
    'DS4608-SR7U2100AZW',
    'DS4608-SR7U2100SGW',
    'DS4678-SR7U2100SFW'
  ].forEach(sku => add('DS46X8N2609', 'COLOMBIA,NOLA', 'EXACTO', sku, 'FIJO', 20, false));

  sheet.getRange(1, 1, 1, PROMOTIONS_HEADERS.length).setValues([PROMOTIONS_HEADERS]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, PROMOTIONS_HEADERS.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, PROMOTIONS_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#17365D')
    .setFontColor('#FFFFFF');
  sheet.getRange(1, 1, Math.max(rows.length + 1, 2), PROMOTIONS_HEADERS.length).createFilter();
  sheet.autoResizeColumns(1, PROMOTIONS_HEADERS.length);
  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 190);
  sheet.setColumnWidth(5, 230);
  sheet.setColumnWidth(12, 320);
  sheet.getRange(2, 8, Math.max(rows.length, 1), 2).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 7, Math.max(rows.length, 1), 1).setNumberFormat('0.00');

  properties.setProperty(PROMOTIONS_FILE_PROPERTY, spreadsheet.getId());

  try {
    DriveApp.getFileById(spreadsheet.getId())
      .moveTo(DriveApp.getFolderById(ANALYTICS_FOLDER_ID));
  } catch (error) {
    Logger.log('No se pudo mover a la carpeta de analítica; quedó en Mi unidad. ' + error.message);
  }

  Logger.log('HOJA DE PROMOCIONES CREADA');
  Logger.log('URL: ' + spreadsheet.getUrl());
  Logger.log('Spreadsheet ID: ' + spreadsheet.getId());
  Logger.log('GID: ' + sheet.getSheetId());
  Logger.log('Reglas cargadas: ' + rows.length);
  Logger.log('Nota: las fechas de vigencia están vacías y deben completarse.');

  return {
    url: spreadsheet.getUrl(),
    spreadsheetId: spreadsheet.getId(),
    gid: sheet.getSheetId(),
    rows: rows.length,
    created: true
  };
}


// =====================================================
// PRUEBAS DESDE EL EDITOR
// =====================================================

/**
 * Prueba si un username existe y consulta su rol.
 *
 * Cambia el valor por un username real antes de ejecutar.
 */
function testUserValidation() {
  const username = 'xcbnpere';

  Logger.log('Username enviado: [' + username + ']');

  const user = validateUser_(username);

  Logger.log('Usuario encontrado: ' + JSON.stringify(user));
}


/**
 * Comprueba que la hoja central esté disponible.
 */
function testAnalyticsConnection() {

  const sheet = getAnalyticsSheet_();

  Logger.log('Hoja: ' + sheet.getName());

  Logger.log('Filas registradas: ' + sheet.getLastRow());

  Logger.log('URL: ' + sheet.getParent().getUrl());

}


function diagnosticarUsuario() {
  const username = 'xcbnpere'.trim().toUpperCase();

  const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();

  const usuarios = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat()
    : [];

  const posicion = usuarios.findIndex(
    usuario => String(usuario).trim().toUpperCase() === username
  );

  Logger.log('Archivo: ' + ss.getName());
  Logger.log('Pestaña: ' + sheet.getName());
  Logger.log('Última fila: ' + lastRow);
  Logger.log('Usuario buscado: [' + username + ']');

  if (posicion >= 0) {
    Logger.log('RESULTADO: Usuario encontrado en la fila ' + (posicion + 2));
  } else {
    Logger.log('RESULTADO: Usuario NO encontrado en la columna A');
  }
}


function testAnalyticsPost() {

  const url = 'https://script.google.com/macros/s/AKfycbyAsqGY809-ijPAuRpH50lmVrJxV_EQ-_xn4PX8YKeW8pboBz-0LFFJdfkBuH5M36H2yA/exec';

  const payload = {
    username: 'xcbnpere',
    events: [
      {
        eventType: 'quote_session',
        quoteId: 'TEST-' + new Date().getTime(),
        country: 'Colombia',
        status: 'TEST'
      }
    ]
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  Logger.log('Código HTTP: ' + response.getResponseCode());
  Logger.log('Respuesta API: ' + response.getContentText());

}
