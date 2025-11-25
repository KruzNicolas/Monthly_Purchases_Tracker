# 📊 Monthly Purchases Tracker (Google Apps Script)

## 📝 Overview

A robust, modular Google Apps Script application designed to log daily expenses into month-based sheets. It automatically groups transactions by week, calculates totals dynamically, and generates visual reports.

## ✨ Features

- **Automated Organization:** Automatically creates monthly sheets and weekly blocks.
- **Dynamic Calculations:** Uses spreadsheet formulas for real-time updates.
- **Visual Reporting:** Generates Weekly, Store-based, and Yearly Overview charts.
- **Email Reports:** Sends HTML-formatted summaries via email.
- **Modern UI:** Includes a responsive Sidebar form and Web App with Dark Mode support.
- **Region Support:** Configurable Timezone handling to ensure accurate date logging regardless of server location.
- **Currency Support:** Formatted for COP (Colombian Peso).

## 🏗️ Project Architecture

The codebase is refactored into specialized modules to ensure maintainability and scalability (DRY principles).

### 1. `Code.js` (Main Controller)

Entry point functions for the Google App Script application.

- **`onOpen()`**: Adds a custom menu to the Google Sheets UI for easy access to the purchase form, charts and report email functionality.
- **`doGet(e)`**: Serves the HTML form for adding purchases when accessed via a web URL.
- **`addPurchase(purchases)`**: Validates the input purchase data, ensures the appropriate monthly sheet and week block exist, adds the purchase row, and recalculates totals.
- **`generateWeeklyChart()`**: Call the function from other module to create a weekly spending chart for the current month.
- **`generateStoreChart()`**: Call the function from other module to create a per-store spending chart for the current month.
- **`generateYearlyOverviewChart()`**: Call the function from other module to create a yearly spending overview chart.
- **`showEmailDialog()`**: Opens a dialog in the Google Sheets UI to allow the user to input parameters like month and email address for sending an email report.
- **`processEmailReport(formObject)`**: Call the function from other module to send an email report of the month specified in the form, with month, weekly and per-store spended totals.
- **`showFormSidebar()`**: Opens a sidebar in the Google Sheets UI to display the purchase input form.

### 2. `SheetLogic.js` (Domain Logic)

Using Namespace pattern or Singleton object to encapsulate the handles sheet manipulation functions.

- **`getMonthlySheet(date)`**: Retrieves the sheet for the given month, or creates it if doesn't exist.
- **`ensureWeekExists(sheet, weekLabel)`**: Ensures the week block for the given week label exists in the sheet or creates it.
- **`getWeekBlock(sheet, weekLabel)`**: Retrieves the row indices for the week block with the given label.
- **`addTransactionRow(sheet, block, rowData)`**: Adds a transaction row to the specified week block with all the necessary formatting.
- **`updateWeekTotal(sheet, block)`**: Recalculates and updates the total for the specified week block.
- **`updateMonthTotal(sheet)`**: Recalculates and updates the total for the month by summing all week totals.

### 3. `Reporting.js` (Reporting Service)

Using Namespace pattern or Singleton object to encapsulate the handles charts and emails functions.

- **`deleteChartByTitle(sheet, title)`**: Removes existing charts by title to prevent duplicates when rows shift.
- **`generateWeeklyChart()`**: Creates a weekly spending chart, intelligently parsing week labels from the sheet data.
- **`generateStoreChart()`**: Creates a per-store spending chart for the current month.
- **`generateYearlyOverviewChart()`**: Creates a yearly spending overview chart.
- **`sendEmailReport(formObject)`**: Sends an email report of the month you specify in the form, with month, weekly and per-store spended totals.

### 4. `Config.js` & `Utils.js` (Helpers)

- **`Config.js`**: Centralized configuration for Sheets, Colors, Formats, Timezone, and Texts.
- **`Utils.js`**: Helper functions for dates, parsing, and formatting.
  - **`parseCurrency(value)`**: Parses a currency string to a number.
  - **`parseDate(dateString)`**: Safely parses dates to the configured timezone to prevent day shifts.
  - **`formatDate(date, format)`**: Formats a Date object to a string using the configured timezone.
  - **`getWeekOfMonth(date)`**: Returns the week number of the month for a given date.
  - **`getWeekDateRange(date)`**: Returns the start and end dates of the week for a given date.
  - **`getOrCreateSheet(name, hide = false)`**: Retrieves or creates a sheet by name to avoid script errors.

## ⚙️ Setup & Configuration

1.  **Script Properties:**

    - Go to **Project Settings > Script Properties**.
    - Add `FORM_URL`: The URL of your deployed Web App (ends in `/exec`).

2.  **Configuration:**
    - Edit `Config.js` to change colors, currency formats, or sheet names without touching the logic.

## 🎨 Sheet Layout & Styling

The script enforces a strict layout to ensure formula integrity:

- **Header:** `#064E3B` (Dark Green)
- **Data Rows:** `#F6FFF6` (Light Green)
- **Week Total:** `#064E3B` (Bold White Text)
- **Month Total:** Located at `I2:J2` (Blue `#3453AC`)

## 🚀 Usage

1.  **Add Purchase:**

    - Open the Sidebar form from the menu `📊 Charts & Reports > ➕ Open Sidebar Form`.
    - Or use the Web App link if configured.
    - Fill in the Date, Store, and Products.
    - Click "Save Purchase".

2.  **Generate Charts:**

    - Navigate to a specific Month sheet.
    - Select `📊 Charts & Reports > Generate Weekly Chart` or `Generate Store Chart`.
    - Select `📊 Charts & Reports > Generate Yearly Overview` to track expenses across all months.
    - Charts will appear directly on the sheet.

3.  **Send Report:**
    - Select `📧 Send Report via Email` from the menu.
    - Choose the month and enter the recipient email.
    - An HTML report with totals and breakdowns will be sent.

---

_Generated for personal expense tracking._
