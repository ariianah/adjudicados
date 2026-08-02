/**
 * Backend de Google Apps Script para el Panel de Adjudicaciones.
 * Guarda pares clave/valor (mismas claves que usa la app: "invitados",
 * "adjudicados", "planes", "valoresUnidades") en una hoja llamada "Storage".
 *
 * INSTALACIÓN:
 * 1. Crear una Google Sheet nueva (puede estar vacía).
 * 2. Extensiones > Apps Script.
 * 3. Borrar el contenido de Code.gs y pegar este archivo completo.
 * 4. Guardar. Implementar > Nueva implementación > Tipo: Aplicación web.
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 5. Copiar la URL que te da ("Web app URL"). Esa es tu VITE_SHEETS_API_URL.
 */

const SHEET_NAME = 'Storage';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['key', 'value', 'updatedAt']);
  }
  return sheet;
}

function findRow_(sheet, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return i + 1; // fila 1-indexada
  }
  return -1;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = (e.parameter.action || 'get');
  const sheet = getSheet_();

  if (action === 'get') {
    const key = e.parameter.key;
    const row = findRow_(sheet, key);
    if (row === -1) return jsonResponse_({ ok: false, error: 'not_found' });
    const value = sheet.getRange(row, 2).getValue();
    return jsonResponse_({ ok: true, key: key, value: value });
  }

  if (action === 'list') {
    const prefix = e.parameter.prefix || '';
    const data = sheet.getDataRange().getValues();
    const keys = data.slice(1)
      .map(function (r) { return r[0]; })
      .filter(function (k) { return String(k).indexOf(prefix) === 0; });
    return jsonResponse_({ ok: true, keys: keys });
  }

  return jsonResponse_({ ok: false, error: 'unknown_action' });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  const sheet = getSheet_();

  if (action === 'set') {
    const key = body.key;
    const value = body.value;
    const row = findRow_(sheet, key);
    const now = new Date().toISOString();
    if (row === -1) {
      sheet.appendRow([key, value, now]);
    } else {
      sheet.getRange(row, 2, 1, 2).setValues([[value, now]]);
    }
    return jsonResponse_({ ok: true, key: key, value: value });
  }

  if (action === 'delete') {
    const key = body.key;
    const row = findRow_(sheet, key);
    if (row !== -1) sheet.deleteRow(row);
    return jsonResponse_({ ok: true, key: key, deleted: true });
  }

  return jsonResponse_({ ok: false, error: 'unknown_action' });
}
