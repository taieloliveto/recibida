// Google Apps Script para guardar las confirmaciones en la planilla.
// Recibida (acto + huevada) -> primera pestaña. Fiesta -> pestaña "Fiesta".
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents || '{}');
    var nombre = String(data.nombre || '').trim().slice(0, 60);
    if (!nombre) return ContentService.createTextOutput('sin nombre');
    var disp = String(data.dispositivo || '');
    if (data.tipo === 'fiesta') {
      var f = ss.getSheetByName('Fiesta') || ss.insertSheet('Fiesta');
      if (f.getLastRow() === 0) f.appendRow(['Fecha y hora', 'Nombre', '¿Viene?', 'Dispositivo']);
      f.appendRow([new Date(), nombre, String(data.respuesta || ''), disp]);
    } else {
      var s = ss.getSheets()[0];
      if (s.getLastRow() === 0) s.appendRow(['Fecha y hora', 'Nombre', 'Dispositivo']);
      s.appendRow([new Date(), nombre, disp]);
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
