// Google Apps Script para guardar las confirmaciones en la planilla.
// Se pega en la planilla: Extensiones > Apps Script.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Fecha y hora', 'Nombre', 'Dispositivo']);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    var data = JSON.parse(e.postData.contents || '{}');
    var nombre = String(data.nombre || '').trim().slice(0, 60);
    if (nombre) {
      sheet.appendRow([new Date(), nombre, String(data.dispositivo || '')]);
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
