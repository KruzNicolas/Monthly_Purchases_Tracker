/**
 * ----------------------------------------------------------------------------
 * MAIN CONTROLLER
 * Entry point for the application.
 * ----------------------------------------------------------------------------
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Purchase Tracker")
    .addItem("➕ Add New Purchase", "showFormSidebar")
    .addToUi();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("form")
    .setTitle("Expense Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- EXPOSED REPORTING FUNCTIONS ---

/**
 * Opens the form in a modal dialog.
 */
function showFormSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("form")
    .setTitle("Expense Tracker")
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "New Purchase");
}

/**
 * Adds purchase data to the DB_Transactions sheet.
 * @param {Array<Object>} purchaseData - Array of product objects.
 * @returns {string} Success message.
 */
function addPurchase(purchaseData) {
  if (!purchaseData || purchaseData.length === 0) {
    throw new Error("No data provided.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    throw new Error(`Sheet "${CONFIG.SHEETS.TRANSACTIONS}" not found.`);
  }

  const batchId = Utils.generateBatchId(purchaseData[0].category);

  const rows = purchaseData.map((item) => [
    batchId,
    item.date, // The sheet usually parses "YYYY-MM-DD" strings correctly
    item.category,
    item.store,
    item.product,
    item.quantity,
    item.price,
    "", // Total is computed by the sheet
  ]);

  const lastRow = Utils.getRealLastRow(sheet);
  const targetRange = sheet.getRange(
    lastRow + 1,
    1,
    rows.length,
    rows[0].length
  );

  targetRange.setValues(rows);

  // Force changes to be written
  SpreadsheetApp.flush();

  return `✅ Successfully added ${rows.length} item(s)!`;
}
