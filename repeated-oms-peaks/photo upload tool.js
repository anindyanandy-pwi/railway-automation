// Code.gs - put this in the Apps Script editor
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Drive Formula')
    .addItem('Open','showFormulaSidebar')
    .addToUi();
}

function showFormulaSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Drive → Sheet Formula')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Called from the sidebar to insert the generated formula into the currently active cell.
 * Returns an object {success: true/false, message: "..."}
 * Requires users to have edit access and the script to be authorized (first time).
 */
function insertFormulaToActiveCell(formulaText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getActiveSheet();
    const range = sh.getActiveRange();
    if (!range) return { success:false, message: 'No active cell selected.' };
    range.setValue(formulaText);
    return { success:true, message: 'Formula inserted.' };
  } catch (err) {
    return { success:false, message: err.message };
  }
}
