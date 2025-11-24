/**
 * ----------------------------------------------------------------------------
 * REPORTING LOGIC
 * Handles Charts and Emails.
 * ----------------------------------------------------------------------------
 */

const Reporting = {
  /**
   * Clears charts in a specific range.
   */
  clearChartsInRange(sheet, rangeA1) {
    const range = sheet.getRange(rangeA1);
    const charts = sheet.getCharts();

    charts.forEach((chart) => {
      const pos = chart.getContainerInfo();
      const row = pos.getAnchorRow();
      const col = pos.getAnchorColumn();

      if (
        range.getRow() <= row &&
        row <= range.getLastRow() &&
        range.getColumn() <= col &&
        col <= range.getLastColumn()
      ) {
        sheet.removeChart(chart);
      }
    });
  },

  /**
   * Generates the Weekly Chart.
   */
  generateWeeklyChart() {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (
      sheet.getName() === CONFIG.SHEETS.DASHBOARD ||
      sheet.getName() === CONFIG.SHEETS.HELPER
    ) {
      SpreadsheetApp.getUi().alert("⚠️ Please go to a MONTH sheet.");
      return;
    }

    const helperSheet = Utils.getOrCreateSheet(CONFIG.SHEETS.HELPER, true);
    const lastRow = sheet.getLastRow();
    const colB = sheet.getRange(1, 2, lastRow).getValues().flat(); // Get the data from monthly sheet
    const colG = sheet.getRange(1, 7, lastRow).getValues().flat(); // Get the data from monthly sheet

    const weekNames = [];
    const weekTotals = [];

    for (let i = 0; i < colB.length; i++) {
      // Iterate through column B to find week totals data
      if (
        // If the cell matches the week total label push the week name and total to respective arrays
        typeof colB[i] === "string" &&
        colB[i].trim().toUpperCase() === CONFIG.TEXT.WEEK_TOTAL_LABEL
      ) {
        weekNames.push("Week " + (weekTotals.length + 1));
        weekTotals.push(Utils.parseCurrency(colG[i]));
      }
    }

    if (weekTotals.length === 0) {
      SpreadsheetApp.getUi().alert("No data found.");
      return;
    }

    this.clearChartsInRange(sheet, "I5:P18");

    helperSheet.getRange("A:B").clearContent();
    helperSheet.getRange("A1:B1").setValues([["Week", "Total"]]);

    const data = weekNames.map((name, i) => [name, weekTotals[i]]);
    helperSheet.getRange(2, 1, data.length, 2).setValues(data);
    helperSheet.getRange(2, 2, data.length, 1).setNumberFormat("#,##0");

    const chart = sheet
      .newChart()
      .asLineChart()
      .setPosition(5, 9, 0, 0)
      .addRange(helperSheet.getRange(2, 1, weekTotals.length, 2))
      .setOption("title", "Weekly Expenses (Total) 📉🔥")
      .setOption("vAxis", { title: "Amount (COP)", format: "#,##0" })
      .setOption("legend", { position: "none" })
      .setOption("pointSize", 7)
      .build();

    sheet.insertChart(chart);
  },

  /**
   * Generates the Store Chart.
   */
  generateStoreChart() {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (
      sheet.getName() === CONFIG.SHEETS.DASHBOARD ||
      sheet.getName() === CONFIG.SHEETS.HELPER
    ) {
      SpreadsheetApp.getUi().alert("⚠️ Please go to a MONTH sheet.");
      return;
    }

    const helperSheet = Utils.getOrCreateSheet(CONFIG.SHEETS.HELPER, true);
    const lastRow = sheet.getLastRow();
    const colD = sheet.getRange(1, 4, lastRow).getValues().flat();
    const colG = sheet.getRange(1, 7, lastRow).getValues().flat();

    const storeTotals = {};
    for (let i = 0; i < colD.length; i++) {
      const store = colD[i];
      const total = Utils.parseCurrency(colG[i]);
      if (typeof store === "string" && store.trim() !== "" && total > 0) {
        storeTotals[store] = (storeTotals[store] || 0) + total;
      }
    }

    this.clearChartsInRange(sheet, "I25:P38");

    helperSheet.getRange("C:D").clearContent();
    helperSheet.getRange("C1:D1").setValues([["Store", "Total"]]);

    const data = Object.entries(storeTotals);
    if (data.length > 0) {
      helperSheet.getRange(2, 3, data.length, 2).setValues(data);
      helperSheet.getRange(2, 4, data.length, 1).setNumberFormat("#,##0");

      const chart = sheet
        .newChart()
        .asColumnChart()
        .setPosition(25, 9, 0, 0)
        .addRange(helperSheet.getRange(2, 3, data.length, 2))
        .setOption("title", "Expenses by Store 🏪💸")
        .setOption("vAxis", { format: "#,##0" })
        .setOption("legend", { position: "none" })
        .build();

      sheet.insertChart(chart);
    }
  },

  /**
   * Generates Yearly Overview.
   */
  generateYearlyOverviewChart() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const helperSheet = Utils.getOrCreateSheet(CONFIG.SHEETS.HELPER, true);
    const dashboardSheet = Utils.getOrCreateSheet(CONFIG.SHEETS.DASHBOARD);

    ss.setActiveSheet(dashboardSheet);
    ss.moveActiveSheet(1);

    const monthData = [];
    ss.getSheets().forEach((sheet) => {
      const name = sheet.getName();
      if (name !== CONFIG.SHEETS.HELPER && name !== CONFIG.SHEETS.DASHBOARD) {
        const tf = sheet.createTextFinder(CONFIG.TEXT.MONTH_TOTAL_LABEL);
        const match = tf.findNext();
        if (match) {
          let val = match.getValue();
          let total = Utils.parseCurrency(val);
          if (total === 0)
            total = Utils.parseCurrency(match.offset(0, 1).getValue());
          monthData.push([name, total]);
        }
      }
    });

    if (monthData.length === 0) {
      SpreadsheetApp.getUi().alert("No data found.");
      return;
    }

    helperSheet.getRange("F:G").clearContent();
    helperSheet.getRange("F1:G1").setValues([["Month", "Total Spent"]]);
    helperSheet.getRange(2, 6, monthData.length, 2).setValues(monthData);
    helperSheet.getRange(2, 7, monthData.length, 1).setNumberFormat("#,##0");

    dashboardSheet.getCharts().forEach((c) => dashboardSheet.removeChart(c));

    const chart = dashboardSheet
      .newChart()
      .asLineChart()
      .setPosition(2, 2, 0, 0)
      .addRange(helperSheet.getRange(2, 6, monthData.length, 2))
      .setOption("title", "Monthly Expenses Overview (Full Year) 🚀")
      .setOption("vAxis", { title: "Amount (COP)", format: "#,##0" })
      .setOption("hAxis", { title: "Month" })
      .setOption("pointSize", 10)
      .setOption("legend", { position: "none" })
      .build();

    dashboardSheet.insertChart(chart);
  },

  /**
   * Sends the email report.
   */
  sendEmailReport(formObject) {
    const { sheetName, email } = formObject;
    if (!sheetName || !email) throw new Error("Missing data.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet not found.");

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) throw new Error("Sheet is empty.");

    // Get Month Total
    let monthTotal = 0;
    const j2Val = sheet.getRange("J2").getValue();
    monthTotal = typeof j2Val === "number" ? j2Val : Utils.parseCurrency(j2Val);

    if (monthTotal === 0) {
      const tf = sheet.createTextFinder(CONFIG.TEXT.MONTH_TOTAL_LABEL);
      const match = tf.findNext();
      if (match) {
        let val = Utils.parseCurrency(match.getValue());
        if (val === 0) val = Utils.parseCurrency(match.offset(0, 1).getValue());
        monthTotal = val;
      }
    }

    // Get Data
    const colB = sheet.getRange(1, 2, lastRow).getValues().flat();
    const colD = sheet.getRange(1, 4, lastRow).getValues().flat();
    const colG = sheet.getRange(1, 7, lastRow).getValues().flat();

    const weeklyData = [];
    const storeTotals = {};

    for (let i = 0; i < lastRow; i++) {
      const valB = colB[i];
      const valD = colD[i];
      const valG = colG[i];
      const amount = Utils.parseCurrency(valG);

      if (
        typeof valB === "string" &&
        valB.trim().toUpperCase() === CONFIG.TEXT.WEEK_TOTAL_LABEL
      ) {
        weeklyData.push({
          week: "Week " + (weeklyData.length + 1),
          total: amount,
        });
      }

      if (
        typeof valD === "string" &&
        valD.trim() !== "" &&
        valD !== "Store" &&
        amount > 0
      ) {
        storeTotals[valD] = (storeTotals[valD] || 0) + amount;
      }
    }

    // Build HTML
    const fmt = (n) => "$" + n.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ".");

    let html = `
      <div style="font-family: sans-serif; color: #333; max-width: 600px; border: 1px solid #e0e0e0; border-radius: 8px;">
        <div style="background-color: ${
          CONFIG.COLORS.HEADER_BG
        }; color: white; padding: 20px; text-align: center;">
          <h1 style="margin:0;">📊 Expense Report</h1>
          <p>${sheetName}</p>
        </div>
        <div style="padding: 20px;">
          <div style="background: #f1f8e9; padding: 15px; text-align: center; margin-bottom: 20px;">
            <span style="display:block; font-weight:bold; color: ${
              CONFIG.COLORS.HEADER_BG
            }">TOTAL SPENT</span>
            <span style="font-size: 24px; font-weight:bold;">${fmt(
              monthTotal
            )}</span>
          </div>
          <h3>📅 Weekly</h3>
          <ul>${weeklyData
            .map((w) => `<li><b>${w.week}:</b> ${fmt(w.total)}</li>`)
            .join("")}</ul>
          <h3>🏪 Stores</h3>
          <ul>${Object.entries(storeTotals)
            .map(([k, v]) => `<li><b>${k}:</b> ${fmt(v)}</li>`)
            .join("")}</ul>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: email,
      subject: `Expense Report: ${sheetName} (${fmt(monthTotal)})`,
      htmlBody: html,
    });
  },
};
