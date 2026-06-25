// ============================================================
// DASHBOARD CONTRATISTAS — LOMA NEGRA ZAPALA
// Google Apps Script — Backend
// ============================================================

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

const SHEETS = {
  partes:      'Partes',
  indicadores: 'Indicadores',
  personal:    'Personal',
};

const HEADERS = {
  partes: ['empresa','fecha','pres','prev','hh','ota','otc','apt','vcp','cha','epp','abl','alv','rcc','rce','timestamp'],
  indicadores: ['mes','anio','hh','acc','dp','freq','grav','timestamp'],
  personal: ['empresa','id','nombre','dni','ingreso','rol','art','med','alt','conf','herr','obs','timestamp'],
};

// ============================================================
// ENTRY POINT — maneja GET y POST
// ============================================================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const body   = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
  const action  = params.action || body.action;
  const empresa = params.empresa || body.empresa;
  const pin     = params.pin || body.pin;

  const cors = ContentService.createTextOutput();
  cors.setMimeType(ContentService.MimeType.JSON);

  try {
    // Autenticación
    if (!autenticar(empresa, pin)) {
      cors.setContent(JSON.stringify({ ok: false, error: 'PIN incorrecto' }));
      return cors;
    }

    let resultado;

    switch (action) {

      // PARTES ------------------------------------------------
      case 'getPartes':
        resultado = getPartes(empresa);
        break;
      case 'addParte':
        resultado = addParte(empresa, body.parte);
        break;

      // INDICADORES -------------------------------------------
      case 'getIndicadores':
        resultado = getIndicadores();
        break;
      case 'setIndicador':
        resultado = setIndicador(body.indicador);
        break;

      // PERSONAL ----------------------------------------------
      case 'getPersonal':
        resultado = getPersonal(empresa);
        break;
      case 'addPersonal':
        resultado = addPersonal(empresa, body.trabajador);
        break;
      case 'updatePersonal':
        resultado = updatePersonal(empresa, body.trabajador);
        break;
      case 'deletePersonal':
        resultado = deletePersonal(empresa, body.id);
        break;

      default:
        resultado = { ok: false, error: 'Acción desconocida: ' + action };
    }

    cors.setContent(JSON.stringify(resultado));
  } catch (err) {
    cors.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return cors;
}

// ============================================================
// AUTENTICACIÓN
// ============================================================

const PINS = {
  'INSAI':     '1111',
  'EUROCLEAN': '2222',
  'ADMIN':     '9999',
};

function autenticar(empresa, pin) {
  if (!empresa || !pin) return false;
  return PINS[empresa] === pin;
}

// ============================================================
// HELPERS DE SHEETS
// ============================================================

function getSheet(nombre) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.appendRow(HEADERS[nombre.toLowerCase()] || []);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function objectToRow(headers, obj) {
  return headers.map(h => obj[h] !== undefined ? obj[h] : '');
}

// ============================================================
// PARTES DIARIOS
// ============================================================

function getPartes(empresa) {
  const sheet = getSheet(SHEETS.partes);
  const todos = sheetToObjects(sheet);
  const filtrados = todos.filter(p => p.empresa === empresa);
  return { ok: true, data: filtrados };
}

function addParte(empresa, parte) {
  if (!parte) return { ok: false, error: 'Sin datos de parte' };
  const sheet = getSheet(SHEETS.partes);
  parte.empresa = empresa;
  parte.timestamp = new Date().toISOString();
  sheet.appendRow(objectToRow(HEADERS.partes, parte));
  return { ok: true };
}

// ============================================================
// INDICADORES (compartidos, sin filtro por empresa)
// ============================================================

function getIndicadores() {
  const sheet = getSheet(SHEETS.indicadores);
  return { ok: true, data: sheetToObjects(sheet) };
}

function setIndicador(ind) {
  if (!ind) return { ok: false, error: 'Sin datos de indicador' };
  const sheet = getSheet(SHEETS.indicadores);
  const todos = sheet.getDataRange().getValues();
  const headers = todos[0];
  const mesCol  = headers.indexOf('mes');
  const anioCol = headers.indexOf('anio');

  // Buscar fila existente para ese mes/año
  let filaExistente = -1;
  for (let i = 1; i < todos.length; i++) {
    if (Number(todos[i][mesCol]) === Number(ind.mes) && Number(todos[i][anioCol]) === Number(ind.anio)) {
      filaExistente = i + 1; // 1-indexed para Sheets
      break;
    }
  }

  ind.timestamp = new Date().toISOString();
  const row = objectToRow(HEADERS.indicadores, ind);

  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { ok: true };
}

// ============================================================
// PERSONAL
// ============================================================

function getPersonal(empresa) {
  const sheet = getSheet(SHEETS.personal);
  const todos = sheetToObjects(sheet);
  return { ok: true, data: todos.filter(w => w.empresa === empresa && w.id !== '') };
}

function addPersonal(empresa, trabajador) {
  if (!trabajador) return { ok: false, error: 'Sin datos de trabajador' };
  const sheet = getSheet(SHEETS.personal);
  trabajador.empresa = empresa;
  trabajador.id = Utilities.getUuid();
  trabajador.timestamp = new Date().toISOString();
  sheet.appendRow(objectToRow(HEADERS.personal, trabajador));
  return { ok: true, id: trabajador.id };
}

function updatePersonal(empresa, trabajador) {
  if (!trabajador || !trabajador.id) return { ok: false, error: 'Sin ID de trabajador' };
  const sheet = getSheet(SHEETS.personal);
  const todos = sheet.getDataRange().getValues();
  const headers = todos[0];
  const idCol = headers.indexOf('id');
  const empCol = headers.indexOf('empresa');

  for (let i = 1; i < todos.length; i++) {
    if (todos[i][idCol] === trabajador.id && todos[i][empCol] === empresa) {
      trabajador.empresa = empresa;
      trabajador.timestamp = new Date().toISOString();
      const row = objectToRow(HEADERS.personal, trabajador);
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Trabajador no encontrado' };
}

function deletePersonal(empresa, id) {
  if (!id) return { ok: false, error: 'Sin ID' };
  const sheet = getSheet(SHEETS.personal);
  const todos = sheet.getDataRange().getValues();
  const headers = todos[0];
  const idCol  = headers.indexOf('id');
  const empCol = headers.indexOf('empresa');

  for (let i = todos.length - 1; i >= 1; i--) {
    if (todos[i][idCol] === id && todos[i][empCol] === empresa) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Trabajador no encontrado' };
}

// ============================================================
// SETUP INICIAL — ejecutar una sola vez para crear las hojas
// ============================================================

function setupSheets() {
  Object.values(SHEETS).forEach(nombre => getSheet(nombre));
  SpreadsheetApp.getUi().alert('✓ Hojas creadas correctamente: Partes, Indicadores, Personal');
}
