# 📊 Monthly Purchases Tracker (Google Apps Script)

## 📝 Overview

A streamlined Google Apps Script application designed to log daily expenses into a centralized database sheet (`DB_Transactions`). It provides a user-friendly interface via a Sidebar or Web App for quick data entry, ensuring all transactions are stored in a structured format for easy analysis.

## ✨ Features

- **Centralized Database:** All transactions are stored in a single `DB_Transactions` sheet.
- **User-Friendly Forms:** Add purchases via a Sidebar in Google Sheets or a standalone Web App.
- **Batch Processing:** Generates unique batch IDs for transaction groups based on date and category.
- **Configurable:** Easy configuration for categories and timezone.
- **Responsive UI:** Clean HTML form for data entry.

## 🏗️ Project Architecture

The codebase is refactored into specialized modules for simplicity and maintainability.

### 1. `Code.js` (Main Controller)

Entry point functions for the Google App Script application.

- **`onOpen()`**: Adds a custom menu to the Google Sheets UI for easy access to the purchase form and reporting options.
- **`doGet(e)`**: Serves the HTML form for adding purchases when accessed via a web URL.
- **`showFormSidebar()`**: Opens a sidebar in the Google Sheets UI to display the purchase input form.
- **`openWebApp()`**: Opens a dialog to redirect to the Web App (requires `FORM_URL` property).
- **`addPurchase(purchaseData)`**: Validates the input purchase data and appends it to the `DB_Transactions` sheet with a generated batch ID.

### 2. `Config.js` (Configuration)

Centralizes configuration constants.

- **`SHEETS`**: Defines sheet names (e.g., `DB_Transactions`).
- **`CATEGORIES`**: List of available expense categories.
- **`TIMEZONE`**: Configured timezone (e.g., "America/Bogota").

### 3. `Utilities.js` (Helpers)

Helper functions for data formatting and processing.

- **`Utils.formatDate(date, format)`**: Formats a Date object to a string using the configured timezone.
- **`Utils.generateBatchId(category)`**: Generates a unique batch ID (e.g., `20240520_1430_Gro`) for grouping transactions.
- **`Utils.getRealLastRow(sheet, column)`**: Finds the last row with actual data, ignoring empty formatted rows.

## ⚙️ Setup & Configuration

1.  **Sheet Setup:**

    - Create a sheet named `DB_Transactions` (or update `Config.js` to match your sheet name).

2.  **Script Properties (Optional for Web App):**

    - Go to **Project Settings > Script Properties**.
    - Add `FORM_URL`: The URL of your deployed Web App.

3.  **Configuration:**
    - Edit `Config.js` to modify categories or timezone settings.

## 🚀 Usage

1.  **Add Purchase:**

    - Open the Sidebar form from the menu `📊 Charts & Reports > ➕ Open Sidebar Form`.
    - Or use the Web App link if configured.
    - Fill in the Date, Store, and Products.
    - Click "Save Purchase".
    - The data will be appended to the `DB_Transactions` sheet.

2.  **Reporting (Menu Options):**
    - The `📊 Charts & Reports` menu includes options for generating charts and sending emails.

---

_Refactored for modularity and centralized data management._
