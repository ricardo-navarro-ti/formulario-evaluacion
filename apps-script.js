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
  'Timestamp', 'Grupo', 'Evaluador',
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
    
    // VALIDACIONES
    if (!data.grupo || data.grupo.trim() === '' || data.grupo.includes('Selecciona')) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Falta completar el campo "Nombre del Grupo".' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (!data.evaluador || data.evaluador.trim() === '' || data.evaluador.includes('Selecciona')) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Falta completar el campo "Evaluador/a".' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // SIEMPRE buscar si ya existe evaluación de este evaluador para este grupo
    const lastRow = sheet.getLastRow();
    let existingRow = null;
    
    if (lastRow > 1) {
      // Leer SOLO las primeras 3 columnas para buscar más rápido
      const searchRange = sheet.getRange(2, 1, lastRow - 1, 3);
      const searchValues = searchRange.getValues();
      
      const grupoLower = data.grupo.trim().toLowerCase();
      const evaluadorLower = data.evaluador.trim().toLowerCase();
      
      for (let i = 0; i < searchValues.length; i++) {
        const row = searchValues[i];
        const grupoExistente = String(row[1] || '').trim().toLowerCase();
        const evaluadorExistente = String(row[2] || '').trim().toLowerCase();
        
        if (grupoExistente === grupoLower && evaluadorExistente === evaluadorLower) {
          existingRow = i + 2; // +2 porque fila 1 es header y el array empieza en 0
          break;
        }
      }
    }
    
    const rowData = [
      new Date().toLocaleString('es-CL'),
      data.grupo, data.evaluador,
      ...data.tecnico,
      ...data.forma,
      data.avgTecnico, data.avgForma, data.notaFinal
    ];
    
    // Si existe, ACTUALIZAR en lugar de crear
    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, HEADERS.length).setValues([rowData]);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', message: 'Evaluación actualizada.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Si NO existe, crear nuevo registro
    sheet.appendRow(rowData);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Evaluación creada.' }))
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
