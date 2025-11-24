/**
 * ----------------------------------------------------------------------------
 * UTILITIES
 * Helper functions for dates, parsing, and formatting.
 * ----------------------------------------------------------------------------
 */

const Utils = {
  /**
   * Parses a value into a number, removing currency symbols.
   * @param {string|number} value 
   * @return {number}
   */
  parseCurrency(value) {
    if (typeof value === "number") return value;
    if (typeof value === "string") {
      const cleaned = Number(value.replace(/[^\d.-]/g, ""));
      return isNaN(cleaned) ? 0 : cleaned;
    }
    return 0;
  },

  /**
   * Formats a date string.
   * @param {Date} date 
   * @param {string} format 
   * @return {string}
   */
  formatDate(date, format) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
  },

  /**
   * Calculates the week number of the month.
   * @param {Date} date 
   * @return {number}
   */
  getWeekOfMonth(date) {
    const firstDay = new Date(date.getFullYear(), date.getMonth(), 1);
    const dom = date.getDate();
    const dow = firstDay.getDay() || 7;
    return Math.ceil((dom + dow - 1) / 7);
  },

  /**
   * Returns the date range string for the week of the given date.
   * @param {Date} date 
   * @return {string}
   */
  getWeekDateRange(date) {
    const day = date.getDay() === 0 ? 7 : date.getDay();
    const start = new Date(date);
    start.setDate(date.getDate() - day + 1);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    
    const fmt = d => this.formatDate(d, CONFIG.FORMATS.DATE_HEADER);
    return `${fmt(start)} - ${fmt(end)}`;
  },

  /**
   * Helper to get or create a sheet.
   * @param {string} name 
   * @param {boolean} hide 
   * @return {GoogleAppsScript.Spreadsheet.Sheet}
   */
  getOrCreateSheet(name, hide = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (hide) sheet.hideSheet();
    }
    return sheet;
  }
};
