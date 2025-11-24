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

/**
 * Adds purchases from the form.
 * @param {Array} purchases
 */
function addPurchase(purchases) {
  if (!Array.isArray(purchases)) throw new Error("Invalid data received.");

  purchases.forEach((p) => {
    const { date, store, product, quantity, price } = p;
    if (!date || !store || !product)
      throw new Error("Missing required fields.");

    const parsedDate = new Date(date);
    if (isNaN(parsedDate)) throw new Error(`Invalid date: ${date}`);

    // 1. Get Sheet
    const sheet = SheetLogic.getMonthlySheet(parsedDate);

    // 2. Determine Week
    const weekNum = Utils.getWeekOfMonth(parsedDate);
    const weekRange = Utils.getWeekDateRange(parsedDate);
    const weekLabel = `Week ${weekNum} (${weekRange})`;

    // 3. Ensure Structure
    SheetLogic.ensureWeekExists(sheet, weekLabel);
    const block = SheetLogic.getWeekBlock(sheet, weekLabel);

    // 4. Prepare Data
    // Formula: Quantity * Price
    const totalFormula = `=E${block.totalRow}*F${block.totalRow}`;
    const rowData = [
      Utils.formatDate(parsedDate, "yyyy-MM-dd"),
      product,
      store,
      Number(quantity),
      Number(price),
      totalFormula,
    ];

    // 5. Add Row
    SheetLogic.addTransactionRow(sheet, block, rowData);

    // 6. Update Totals
    // Re-fetch block because rows shifted
    const updatedBlock = SheetLogic.getWeekBlock(sheet, weekLabel);
    SheetLogic.updateWeekTotal(sheet, updatedBlock);
    SheetLogic.updateMonthTotal(sheet);
  });

  return "✅ Purchase(s) added successfully!";
}

// --- EXPOSED REPORTING FUNCTIONS ---

function generateWeeklyChart() {
  Reporting.generateWeeklyChart();
}

function generateStoreChart() {
  Reporting.generateStoreChart();
}

function generateYearlyOverviewChart() {
  Reporting.generateYearlyOverviewChart();
}

function showEmailDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const sheetNames = [];

  allSheets.forEach((s) => {
    const name = s.getName();
    if (name !== CONFIG.SHEETS.HELPER && name !== CONFIG.SHEETS.DASHBOARD) {
      sheetNames.push(name);
    }
  });

  let html = `
    <style>
      body { font-family: 'Segoe UI', sans-serif; padding: 15px; color: #333; }
      h3 { margin-top: 0; color: ${CONFIG.COLORS.HEADER_BG}; }
      label { font-weight: bold; font-size: 0.9em; }
      select, input { width: 100%; padding: 10px; margin: 5px 0 20px 0; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
      button { background-color: ${CONFIG.COLORS.HEADER_BG}; color: white; padding: 12px 20px; border: none; border-radius: 4px; cursor: pointer; width: 100%; font-size: 1em; }
      button:hover { opacity: 0.9; }
      .note { font-size: 0.8em; color: #666; margin-top: 10px; text-align: center; }
    </style>
    <h3>📤 Send Monthly Report</h3>
    <form id="reportForm" onsubmit="event.preventDefault(); send()">
      <label for="sheetName">1. Select Month to report:</label>
      <select name="sheetName" id="sheetName">
  `;

  sheetNames.forEach((name) => {
    html += `<option value="${name}">${name}</option>`;
  });

  html += `
      </select>
      
      <label for="email">2. Destination Email:</label>
      <input type="email" name="email" id="email" placeholder="name@example.com" required>
      
      <button type="button" onclick="send()">✉️ Send Report</button>
      <div class="note">The email will include the weekly and store breakdown.</div>
    </form>
    <script>
      function send() {
        const btn = document.querySelector('button');
        const form = document.getElementById('reportForm');
        const email = form.email.value;
        if(!email) { alert('Please enter a valid email.'); return; }
        
        btn.innerText = 'Sending... ⏳';
        btn.disabled = true;
        
        google.script.run
          .withSuccessHandler(function() { google.script.host.close(); })
          .withFailureHandler(function(e) { alert('Error: ' + e); btn.disabled = false; btn.innerText = 'Try Again'; })
          .processEmailReport(form);
      }
    </script>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(350)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Report Assistant");
}

function processEmailReport(formObject) {
  Reporting.sendEmailReport(formObject);
}

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
