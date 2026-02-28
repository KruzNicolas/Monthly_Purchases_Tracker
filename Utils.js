/**
 * ----------------------------------------------------------------------------
 * UTILITIES
 * Helper functions for data formatting and processing.
 * ----------------------------------------------------------------------------
 */

const Utils = {
  /**
   * Formats a date object into a string.
   * @param {Date} date - The date to format.
   * @param {string} format - The format string (e.g., "yyyyMMdd_HHmm").
   * @returns {string} The formatted date string.
   */
  formatDate: function (date, format = "yyyyMMdd_HHmm") {
    return Utilities.formatDate(date, CONFIG.TIMEZONE, format);
  },

  /**
   * Generates a unique batch ID based on date and category.
   * Example: 20240520_1430_Pha
   * @param {string} category - The category name.
   * @returns {string} The generated batch ID.
   */
  generateBatchId: function (category) {
    const now = new Date();
    const timestamp = this.formatDate(now, "yyyyMMdd_HHmm");
    const categoryCode = category ? category.substring(0, 3) : "Unk";
    return `${timestamp}_${categoryCode}`;
  },

  /**
   * Finds the last row with actual data in a specific column.
   * Useful when sheet.getLastRow() returns rows with only formatting.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @param {number} column - Column index (1-based).
   * @returns {number} The last row with data.
   */
  getRealLastRow: function (sheet, column = 1) {
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return 0;
    const values = sheet.getRange(1, column, lastRow, 1).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][0] !== "") {
        return i + 1;
      }
    }
    return 0;
  },
};
