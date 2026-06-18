// ============================================================
// GOOGLE APPS SCRIPT — Backend Formulario de Evaluacion
// ============================================================
// INSTRUCCIONES:
//   1. Abre tu Google Sheet -> Extensiones -> Apps Script
//   2. Borra todo el codigo por defecto y pega este archivo
//   3. Guarda (Ctrl+S) -> ponle nombre "Backend Formulario"
//   4. Clic en Implementar -> Nueva implementacion
//   5. Tipo: Aplicacion web | Ejecutar como: Yo | Acceso: Cualquier persona
//   6. Clic en Implementar -> Autoriza los permisos
//   7. COPIA la URL que aparece -> pegala en index.html y resultados.html
// ============================================================

const HEADERS = [
  'Timestamp', 'Grupo', 'Evaluador', 'Fecha',
  'T1_Entorno', 'T2_Actores', 'T3_FactoresEntrada', 'T4_FactoresPerdida',
  'T5_Medidas', 'T6_EquidadSocial', 'T7_Sostenibilidad', 'T8_BeneficioEcon',
  'T9_Coherencia', 'T10_Prompts',
  'F1_Estructura', 'F2_Claridad', 'F3_DisenoVisual',
  'F4_RecursosVisuales', 'F5_CoherenciaEstilo', 'F6_Ortografia',
  'PromedioTecnico', 'PromedioForma', 'NotaFinal'
];

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const data = JSON.parse(e.postData.contents);
    sheet.appendRow([
      new Date().toLocaleString('es-CL'),
      data.grupo, data.evaluador, data.fecha,
      ...data.tecnico,
      ...data.forma,
      data.avgTecnico, data.avgForma, data.notaFinal
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = sheet.getDataRange().getValues();
  let rows = [];
  if (values.length >= 2) {
    const headers = values[0];
    rows = values.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
  }
  const json     = JSON.stringify(rows);
  const callback = e && e.parameter && e.parameter.callback;
  return ContentService
    .createTextOutput(callback ? callback + '(' + json + ')' : json)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
