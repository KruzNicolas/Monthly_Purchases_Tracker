/**
 * ----------------------------------------------------------------------------
 * MAIN CONTROLLER
 * Entry point for the application.
 * ----------------------------------------------------------------------------
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Charts & Reports")
    .addItem("➕ Open Sidebar Form", "showFormSidebar")
    .addSeparator()
    .addItem("Generate Weekly Chart 📉 (Current Sheet)", "generateWeeklyChart")
    .addItem("Generate Store Chart 🏪 (Current Sheet)", "generateStoreChart")
    .addSeparator()
    .addItem(
      "🚀 Generate YEARLY SUMMARY (Dashboard)",
      "generateYearlyOverviewChart"
    )
    .addSeparator()
    .addItem("📧 Send Report via Email", "showEmailDialog")
    .addToUi();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("form")
    .setTitle("Expense Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- EXPOSED REPORTING FUNCTIONS ---

/**
 * Opens the form in the sidebar.
 */
function showFormSidebar() {
  const html =
    HtmlService.createHtmlOutputFromFile("form").setTitle("Expense Tracker");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens a dialog with a link to the Web App.
 * This is more reliable than auto-redirecting, which can be blocked by browsers.
 */
/**
 * Opens a dialog with a link to the Web App.
 * Uses the URL defined in Script Properties (Environment Variable: FORM_URL).
 */
function openWebApp() {
  const url = PropertiesService.getScriptProperties().getProperty("FORM_URL");

  if (!url) {
    SpreadsheetApp.getUi().alert(
      "⚠️ FORM_URL not found in Script Properties!\n\n1. Go to Project Settings > Script Properties.\n2. Add a property named 'FORM_URL'.\n3. Paste your Web App URL as the value."
    );
    return;
  }

  const html = `
    <script>
      window.open('${url}', '_blank');
      google.script.host.close();
    </script>
    <p>Redirecting...</p>
  `;

  const userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(200)
    .setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(userInterface, "Opening...");
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

  return `✅ Successfully added ${rows.length} items to "${CONFIG.SHEETS.TRANSACTIONS}"!`;
}
