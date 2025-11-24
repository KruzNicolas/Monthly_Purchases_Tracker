/**
 * ----------------------------------------------------------------------------
 * SHEET LOGIC
 * Handles all direct interactions with the Monthly Sheets.
 * ----------------------------------------------------------------------------
 */

const SheetLogic = {
  /**
   * Gets or creates the monthly sheet based on the date.
   * @param {Date} date
   * @return {GoogleAppsScript.Spreadsheet.Sheet}
   */
  getMonthlySheet(date) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const monthOnly = Utils.formatDate(date, "MMMM");
    const monthWithYear = Utils.formatDate(date, "MMMM yyyy");

    let sheet = ss.getSheetByName(monthOnly);
    if (!sheet) sheet = ss.insertSheet(monthOnly);

    // Set Title
    sheet
      .getRange("B1")
      .setValue(monthWithYear)
      .setFontWeight("bold")
      .setFontSize(13)
      .setFontColor(CONFIG.COLORS.HEADER_BORDER)
      .setBackground(null);

    return sheet;
  },

  /**
   * Ensures the week block exists.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @param {string} weekLabel
   */
  ensureWeekExists(sheet, weekLabel) {
    const lastRow = sheet.getLastRow();
    const colB =
      lastRow > 0 ? sheet.getRange(1, 2, lastRow).getValues().flat() : [];

    if (colB.indexOf(weekLabel) !== -1) return;

    const startRow = lastRow === 0 ? 3 : lastRow + 2;

    // Week Label
    sheet
      .getRange(`B${startRow}`)
      .setValue(weekLabel)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor(CONFIG.COLORS.LABEL_TEXT);

    // Header
    sheet
      .getRange(`B${startRow + 1}:G${startRow + 1}`)
      .setValues([["Date", "Product", "Store", "Quantity", "Price", "Total"]])
      .setFontWeight("bold")
      .setBackground(CONFIG.COLORS.HEADER_BG)
      .setFontColor(CONFIG.COLORS.HEADER_TEXT)
      .setHorizontalAlignment("center")
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        CONFIG.COLORS.HEADER_BORDER,
        SpreadsheetApp.BorderStyle.SOLID
      );

    // Week Total Row (Initial)
    const totalRow = startRow + 2;
    sheet
      .getRange(`B${totalRow}`)
      .setValue(CONFIG.TEXT.WEEK_TOTAL_LABEL)
      .setFontWeight("bold")
      .setFontColor(CONFIG.COLORS.WEEK_TOTAL_TEXT)
      .setBackground(CONFIG.COLORS.WEEK_TOTAL_BG);

    sheet
      .getRange(`G${totalRow}`)
      .setValue(0)
      .setNumberFormat(CONFIG.FORMATS.CURRENCY)
      .setFontWeight("bold")
      .setFontColor(CONFIG.COLORS.WEEK_TOTAL_TEXT)
      .setBackground(CONFIG.COLORS.WEEK_TOTAL_BG);

    // Spacer
    sheet.insertRowsAfter(totalRow, 1);
  },

  /**
   * Finds the coordinates of a week block.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @param {string} weekLabel
   */
  getWeekBlock(sheet, weekLabel) {
    const lastRow = sheet.getLastRow();
    const colB = sheet.getRange(1, 2, lastRow).getValues().flat();

    const headerIndex = colB.indexOf(weekLabel);
    if (headerIndex === -1) return null;

    const weekLabelRow = headerIndex + 1;
    const headerRow = weekLabelRow + 1;
    let row = headerRow + 1;
    let totalRow = null;

    while (row <= lastRow) {
      const cellVal = sheet.getRange(row, 2).getValue();
      if (
        typeof cellVal === "string" &&
        cellVal.startsWith(CONFIG.TEXT.WEEK_TOTAL_LABEL)
      ) {
        totalRow = row;
        break;
      }
      row++;
    }

    return { weekLabelRow, headerRow, firstDataRow: headerRow + 1, totalRow };
  },

  /**
   * Adds a transaction row and updates formulas.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @param {Object} block
   * @param {Array} rowData
   */
  addTransactionRow(sheet, block, rowData) {
    // Insert before Total
    sheet.insertRowBefore(block.totalRow);

    // Reset styles
    sheet
      .getRange(`B${block.totalRow}:G${block.totalRow}`)
      .setFontWeight("normal")
      .setBackground(null)
      .setFontColor(CONFIG.COLORS.ROW_TEXT)
      .setBorder(false, false, false, false, false, false);

    // Set Data & Style
    sheet
      .getRange(`B${block.totalRow}:G${block.totalRow}`)
      .setValues([rowData])
      .setBackground(CONFIG.COLORS.ROW_BG)
      .setFontColor(CONFIG.COLORS.ROW_TEXT)
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        CONFIG.COLORS.ROW_BORDER,
        SpreadsheetApp.BorderStyle.SOLID
      );

    // Format Currency
    sheet
      .getRange(`F${block.totalRow}:G${block.totalRow}`)
      .setNumberFormat(CONFIG.FORMATS.CURRENCY);

    // Clean spacer row
    sheet
      .getRange(`B${block.totalRow + 1}:G${block.totalRow + 1}`)
      .setBackground(null);
  },

  /**
   * Updates the Week Total formula.
   */
  updateWeekTotal(sheet, block) {
    const { firstDataRow, totalRow } = block;
    const formula = `=SUM(G${firstDataRow}:G${totalRow - 1})`;

    sheet
      .getRange(`B${totalRow}:G${totalRow}`)
      .setBackground(CONFIG.COLORS.WEEK_TOTAL_BG)
      .setFontColor(CONFIG.COLORS.WEEK_TOTAL_TEXT)
      .setFontWeight("bold");

    sheet.getRange(totalRow, 2).setValue(CONFIG.TEXT.WEEK_TOTAL_LABEL);
    sheet
      .getRange(totalRow, 7)
      .setFormula(formula)
      .setNumberFormat(CONFIG.FORMATS.CURRENCY);

    // Clean next row
    sheet
      .getRange(`${totalRow + 1}:${totalRow + 1}`)
      .setBackground(null)
      .setFontColor(null);
  },

  /**
   * Updates the Month Total formula.
   */
  updateMonthTotal(sheet) {
    sheet
      .getRange("I2")
      .setValue(CONFIG.TEXT.MONTH_TOTAL_LABEL)
      .setFontWeight("bold")
      .setFontColor(CONFIG.COLORS.MONTH_TOTAL_TEXT)
      .setBackground(CONFIG.COLORS.MONTH_TOTAL_BG);

    sheet
      .getRange("J2")
      .setFormula(`=SUMIF(B:B, "${CONFIG.TEXT.WEEK_TOTAL_LABEL}", G:G)`)
      .setNumberFormat(CONFIG.FORMATS.CURRENCY)
      .setFontWeight("bold")
      .setFontColor(CONFIG.COLORS.MONTH_TOTAL_TEXT)
      .setBackground(CONFIG.COLORS.MONTH_TOTAL_BG);
  },
};
