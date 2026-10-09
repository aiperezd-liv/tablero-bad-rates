/**
 * BAD RATES DASHBOARD - BACKEND GOOGLE APPS SCRIPT
 * Proyecto: Riesgos Negocios Financieros
 */

// Configuración de nombres de las pestañas en tu Google Sheet
const SHEET_EXT_TOTAL = "Ext_Total_Adj";
const SHEET_E_CC = "E_CC";

/**
 * Función principal llamada desde el frontend mediante google.script.run
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheetTotal = ss.getSheetByName(SHEET_EXT_TOTAL);
    const sheetECC = ss.getSheetByName(SHEET_E_CC);

    if (!sheetTotal) throw new Error("No se encontró la pestaña '" + SHEET_EXT_TOTAL + "'");
    if (!sheetECC) throw new Error("No se encontró la pestaña '" + SHEET_E_CC + "'");

    const dataTotal = extractSheetData_(sheetTotal);
    const dataECC = extractSheetData_(sheetECC);

    return {
      success: true,
      dataTotal: dataTotal,
      dataECC: dataECC,
      timestamp: new Date().toISOString()
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * Función auxiliar para convertir filas de hoja de cálculo a un arreglo de Objetos JSON
 */
function extractSheetData_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return [];

  const rangeValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = rangeValues[0].map(h => String(h).trim().toUpperCase());

  const result = [];

  for (let r = 1; r < rangeValues.length; r++) {
    const row = rangeValues[r];
    const rowObj = {};
    let hasData = false;

    for (let c = 0; c < headers.length; c++) {
      const header = headers[c];
      let val = row[c];

      if (val !== "" && val !== null && val !== undefined) {
        hasData = true;
      }

      // Tratamiento especial para la fecha de COSECHA
      if (header === 'COSECHA' && val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-01");
      } else if (typeof val === 'number') {
        val = Number(val);
      }

      rowObj[header] = val;
    }

    if (hasData) {
      result.push(rowObj);
    }
  }

  return result;
}

/**
 * Permite servir la aplicación como una Web App independiente si se requiere
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Bad Rates Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
