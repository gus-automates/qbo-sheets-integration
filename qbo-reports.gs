/**
 * ============================================================
 * FILE: qbo-reports.gs
 * ============================================================
 * This file contains all report and data functions.
 * Each function pulls a specific dataset from QuickBooks Online
 * and writes it into a named Google Sheet tab.
 *
 * Every function here relies on qboFetch_() from qbo-api-core.gs.
 * Make sure that file is included in the same Apps Script project.
 *
 * Reports available:
 *  1. pullCustomersToSheet()    → Customer list from QBO Query API
 *  2. pullPnLByCustomer()       → Profit & Loss broken down by Customer
 *  3. pullOpenPOsAndBills()     → Open Purchase Orders + linked Bills
 *  4. pullActualsPnL()          → Monthly P&L filtered by accounting class
 * ============================================================
 */


// ── REPORT 1: CUSTOMER LIST ────────────────────────────────────────────────────

/**
 * Pulls a list of all customers from QuickBooks Online
 * and writes them into a sheet named "QBO_CUSTOMERS".
 *
 * Uses the QBO Query API (SQL-like syntax called QQL).
 * Fields: Id, DisplayName, CompanyName, Email, Active status, Balance
 */
function pullCustomersToSheet() {
  // QQL query — similar to SQL. MAXRESULTS caps the response (QBO max is 1000).
  const query = `
    SELECT Id, DisplayName, CompanyName, PrimaryEmailAddr, Active, Balance
    FROM Customer
    MAXRESULTS 1000
  `;

  const json      = qboFetch_("query", { query });
  const customers = json?.QueryResponse?.Customer || [];

  // Map each customer object to a flat array matching our column order
  const rows = customers.map(c => ([
    c.Id || "",
    c.DisplayName || "",
    c.CompanyName || "",
    // Email is a nested object: { Address: "..." }
    c.PrimaryEmailAddr?.Address || "",
    // Active defaults to true if the field is missing
    c.Active !== false,
    typeof c.Balance === "number" ? c.Balance : "",
  ]));

  const header   = ["Id", "Display Name", "Company Name", "Email", "Active", "Balance"];
  const sheetName = "QBO_CUSTOMERS";

  writeToSheet_(sheetName, header, rows);
  SpreadsheetApp.getUi().alert(`Pulled ${rows.length} customers into "${sheetName}"`);
}


// ── REPORT 2: PROFIT & LOSS BY CUSTOMER ───────────────────────────────────────

/**
 * Pulls a Profit & Loss report summarized by Customer for a given date range.
 *
 * Uses the QBO Reports API — a separate system from the Query API.
 * The Reports API returns pre-aggregated financial statements with a
 * nested row structure (Sections > Data rows > Summaries).
 *
 * This function flattens that nested structure into rows a spreadsheet can display,
 * adding a "Path" column so you can see where each line sits in the P&L hierarchy
 * (e.g., "Income > Services > Consulting").
 */
function pullPnLByCustomer() {
  // Date range for the report — adjust as needed
  const params = {
    start_date:           "2025-01-01",
    end_date:             "2025-12-31",
    summarize_column_by:  "Customers",  // Creates one column per customer
    accounting_method:    "Accrual",    // or "Cash"
  };

  const report    = qboFetch_("reports/ProfitAndLoss", params);
  const sheetName = "PNL_BY_CUSTOMER";

  // Build header row from the report's column definitions
  const cols      = report?.Columns?.Column || [];
  const colTitles = cols.map(c => c.ColTitle || c.ColType || "");
  const header    = ["Path", "Row Type", ...colTitles];

  const rows = [];

  // The report's row structure is deeply nested.
  // We recursively walk it and flatten into a 2D array.
  flattenReportRows_(report?.Rows?.Row || [], [], rows);

  writeToSheet_(sheetName, header, rows);
  SpreadsheetApp.getUi().alert(`Wrote ${rows.length} rows into "${sheetName}"`);
}

/**
 * Recursively flattens the nested QBO report row structure.
 *
 * QBO report rows have three types:
 *  - Section:  A group heading (e.g. "Income", "Cost of Goods Sold")
 *  - Data:     An individual account line with values
 *  - Summary:  A subtotal row at the end of each section
 *
 * Each call processes one level of rows and recurses into child sections.
 *
 * @param {Array}  rows    - Array of row objects from the report JSON
 * @param {Array}  path    - Breadcrumb trail of parent section names (builds the "Path" column)
 * @param {Array}  out     - The output array we're appending flattened rows into
 */
function flattenReportRows_(rows, path, out) {
  rows.forEach(r => {
    const rowType    = r.type || "";
    const headerTitle = getFirstColValue_(r?.Header?.ColData);

    // When we enter a Section, extend the path breadcrumb
    const nextPath = (rowType === "Section" && headerTitle)
      ? [...path, headerTitle]
      : path;

    // Data rows: individual account lines
    if (Array.isArray(r.ColData) && r.ColData.length) {
      out.push([
        nextPath.join(" > "),
        "Data",
        ...r.ColData.map(cd => cd?.value ?? ""),
      ]);
    }

    // Summary rows: subtotals for each section
    if (Array.isArray(r?.Summary?.ColData) && r.Summary.ColData.length) {
      out.push([
        nextPath.join(" > "),
        "Summary",
        ...r.Summary.ColData.map(cd => cd?.value ?? ""),
      ]);
    }

    // Recurse into child rows if this is a Section
    if (Array.isArray(r?.Rows?.Row)) {
      flattenReportRows_(r.Rows.Row, nextPath, out);
    }
  });
}

/**
 * Safely extracts the first value from a ColData array.
 * ColData is how QBO report rows store cell values: [{ value: "..." }, ...]
 *
 * @param {Array} colDataArr
 * @returns {string}
 */
function getFirstColValue_(colDataArr) {
  if (!Array.isArray(colDataArr) || !colDataArr.length) return "";
  return colDataArr[0]?.value ?? "";
}


// ── REPORT 3: OPEN PURCHASE ORDERS & LINKED BILLS ─────────────────────────────

/**
 * Pulls all open Purchase Orders from QBO and writes them to columns A–F.
 * For each PO, it also fetches any linked Bills and writes those to columns I–K.
 *
 * This uses the QBO Query API to get POs, then a separate API call per Bill
 * to get detailed billing info (Bills aren't fully embedded in PO responses).
 *
 * Columns A–F (POs):    PO #, Date, Supplier, Net Amount, Gross Total, Open Balance
 * Columns I–K (Bills):  Bill #, PO #, Bill Net Amount
 */
function pullOpenPOsAndBills() {
  const ss        = SpreadsheetApp.getActive();
  const sheetName = "PO_TRACKER";
  const sh        = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Clear existing data before writing fresh results
  const lastRow  = Math.max(sh.getLastRow(), 1);
  sh.getRange(1, 1, lastRow, 6).clearContent();  // Columns A–F (POs)
  sh.getRange(1, 9, lastRow, 3).clearContent();  // Columns I–K (Bills)

  // ── Fetch all open POs ──────────────────────────────────────────────────────
  // QBO returns all POs; we filter client-side for Open status
  const poJson = qboFetch_("query", { query: "SELECT * FROM PurchaseOrder MAXRESULTS 1000" });
  const openPOs = (poJson?.QueryResponse?.PurchaseOrder || [])
    .filter(po => po.POStatus === "Open");

  // ── Build PO rows ───────────────────────────────────────────────────────────
  const poRows = openPOs.map(po => {
    const grossTotal = po.TotalAmt || 0;
    // QBO's TotalAmt includes tax. Subtract TotalTax to get the pre-tax (net) amount.
    const tax        = po?.TxnTaxDetail?.TotalTax || 0;
    const netAmount  = grossTotal - tax;

    return [
      po.DocNumber       || "",   // A - PO Number
      po.TxnDate         || "",   // B - Date
      po?.VendorRef?.name || "",  // C - Supplier name
      netAmount,                  // D - Net Amount (before tax)
      grossTotal,                 // E - Gross Total (including tax)
      "",                         // F - Open Balance (not exposed directly by QBO API; can be calculated externally)
    ];
  });

  // Sort by date, newest first
  poRows.sort((a, b) => new Date(b[1]) - new Date(a[1]));

  // ── Build Bill rows ─────────────────────────────────────────────────────────
  // Each PO may have linked transactions. We look for Bills specifically.
  const billRows = [];

  openPOs.forEach(po => {
    const poNumber    = po.DocNumber || "";
    // LinkedTxn is QBO's way of linking related transactions (e.g. a PO to a Bill)
    const linkedBills = (po.LinkedTxn || []).filter(t => t.TxnType === "Bill");

    linkedBills.forEach(link => {
      try {
        // Fetch the full Bill record using its transaction ID
        const billJson = qboFetch_(`bill/${link.TxnId}`);
        const bill     = billJson?.Bill;
        if (!bill) return;

        // The pre-tax amount is stored inside TaxLine details, not as a top-level field.
        // We sum NetAmountTaxable across all tax lines to get the total pre-tax amount.
        const taxLines  = bill?.TxnTaxDetail?.TaxLine || [];
        const netAmount = taxLines.reduce((sum, line) => {
          return sum + (line?.TaxLineDetail?.NetAmountTaxable || 0);
        }, 0);

        billRows.push([
          bill.DocNumber || "",  // I - Bill Number
          poNumber,              // J - Linked PO Number
          netAmount,             // K - Bill Net Amount (pre-tax)
        ]);
      } catch (e) {
        // Log individual failures without stopping the whole report
        Logger.log(`Could not fetch Bill ${link.TxnId}: ${e.message}`);
      }
    });
  });

  // ── Write to sheet ──────────────────────────────────────────────────────────
  const poHeader   = ["PO #", "Date", "Supplier", "Net Amount", "Gross Total", "Open Balance"];
  const billHeader = ["Bill #", "PO #", "Bill Net Amount"];

  sh.getRange(1, 1, 1, poHeader.length).setValues([poHeader]);
  if (poRows.length) sh.getRange(2, 1, poRows.length, 6).setValues(poRows);

  sh.getRange(1, 9, 1, billHeader.length).setValues([billHeader]);
  if (billRows.length) sh.getRange(2, 9, billRows.length, 3).setValues(billRows);

  SpreadsheetApp.getUi().alert(
    `Done. ${poRows.length} open POs in columns A–F, ${billRows.length} linked Bills in columns I–K.`
  );
}


// ── REPORT 4: MONTHLY ACTUALS P&L BY PROJECT ──────────────────────────────────

/**
 * Prompts the user for a month and year, then pulls the P&L for that period
 * filtered by a specific accounting class and broken out by project/customer.
 *
 * This is split into two functions:
 *  - pullActualsPnL()           → Handles the UI prompts and validation
 *  - fetchAndWriteActualsPnL_() → Does the API call and sheet writing
 *
 * The target sheet must already exist and be named in the format: Actuals-MarYY
 * (e.g. "Actuals-Mar26" for March 2026)
 *
 * Before writing data, it runs a classification check to ensure all transactions
 * have been assigned to an accounting class in QBO. If any are unclassified,
 * it stops and alerts the user to fix them first.
 */
function pullActualsPnL() {
  const ui = SpreadsheetApp.getUi();

  // ── Prompt for month ────────────────────────────────────────────────────────
  const monthResult = ui.prompt(
    "Pull Actuals P&L — Step 1 of 2",
    "Enter month number (1–12):",
    ui.ButtonSet.OK_CANCEL
  );
  if (monthResult.getSelectedButton() !== ui.Button.OK) return;

  const month = parseInt(monthResult.getResponseText().trim());
  if (isNaN(month) || month < 1 || month > 12) {
    ui.alert("Invalid month. Please enter a number between 1 and 12.");
    return;
  }

  // ── Prompt for year ─────────────────────────────────────────────────────────
  const yearResult = ui.prompt(
    "Pull Actuals P&L — Step 2 of 2",
    "Enter year (e.g. 2026):",
    ui.ButtonSet.OK_CANCEL
  );
  if (yearResult.getSelectedButton() !== ui.Button.OK) return;

  const year = parseInt(yearResult.getResponseText().trim());
  if (isNaN(year) || year < 2020) {
    ui.alert("Invalid year.");
    return;
  }

  fetchAndWriteActualsPnL_(month, year);
}

/**
 * Fetches and writes the monthly P&L report.
 * Called by pullActualsPnL() after user input is validated.
 *
 * @param {number} month - Month number (1–12)
 * @param {number} year  - 4-digit year (e.g. 2026)
 */
function fetchAndWriteActualsPnL_(month, year) {
  const ui = SpreadsheetApp.getUi();

  // Build the expected sheet name (e.g. "Actuals-Mar26")
  const monthNames  = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const monthLabel  = monthNames[month - 1];
  const yearLabel   = String(year).slice(2);       // "2026" → "26"
  const sheetName   = `Actuals-${monthLabel}${yearLabel}`;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);

  if (!sh) {
    ui.alert(`Sheet "${sheetName}" not found. Please create it first.`);
    return;
  }
  ss.setActiveSheet(sh);

  // Build date range for the full month
  const startDate = `${year}-${String(month).padStart(2, "0")}-01`;
  const lastDay   = new Date(year, month, 0).getDate(); // Day 0 of next month = last day of this month
  const endDate   = `${year}-${String(month).padStart(2, "0")}-${lastDay}`;

  // ── Pre-flight: verify all transactions are classified ──────────────────────
  // Stops here and alerts the user if any transactions are missing a class.
  // Replace the class IDs below with your own from QBO > Settings > Classes.
  const CLASS_ID_A = "YOUR_CLASS_ID_A"; // e.g. "Billable" class
  const CLASS_ID_B = "YOUR_CLASS_ID_B"; // e.g. "Overhead" class

  if (!checkUnclassifiedTransactions_(startDate, endDate, CLASS_ID_A, CLASS_ID_B)) return;

  // ── Fetch the P&L report filtered by class A ─────────────────────────────
  const report = qboFetch_("reports/ProfitAndLoss", {
    start_date:          startDate,
    end_date:            endDate,
    summarize_column_by: "Customers",
    accounting_method:   "Accrual",
    class:               CLASS_ID_A,
  });

  // ── Extract column structure ─────────────────────────────────────────────
  const allColumns  = report?.Columns?.Column || [];
  // Skip index 0 — it's always a blank "Account" label column, not a project
  const projectCols = allColumns
    .map((col, i) => ({ title: col.ColTitle || "", index: i }))
    .filter(col => col.index > 0);

  // ── Flatten report rows ──────────────────────────────────────────────────
  const dataRows = [];

  function flattenRows(rows) {
    rows.forEach(r => {
      const headerVal   = r?.Header?.ColData?.[0]?.value || "";
      const summaryVals = r?.Summary?.ColData || [];

      // Section header row (e.g. "Income", "COGS") — written as a label-only row
      if (headerVal) {
        dataRows.push([headerVal, ...projectCols.map(() => "")]);
      }

      // Data row — individual account line with values per project
      if (Array.isArray(r.ColData) && r.ColData.length) {
        const label  = r.ColData[0]?.value || "";
        const values = projectCols.map(col => {
          const raw = r.ColData[col.index]?.value || "";
          return raw === "" ? "" : parseFloat(raw) || 0;
        });
        dataRows.push([label, ...values]);
      }

      // Recurse into child rows
      if (Array.isArray(r?.Rows?.Row)) flattenRows(r.Rows.Row);

      // Summary row — subtotals per section
      if (Array.isArray(r?.Summary?.ColData) && r.Summary.ColData.length) {
        const summaryLabel = r.Summary.ColData[0]?.value || "Total";
        const values = projectCols.map(col => {
          const raw = summaryVals[col.index]?.value || "";
          return raw === "" ? "" : parseFloat(raw) || 0;
        });
        dataRows.push([summaryLabel, ...values]);
      }
    });
  }

  flattenRows(report?.Rows?.Row || []);

  // ── Write to sheet ───────────────────────────────────────────────────────
  // Data starts at row 6, column D — leaving space for any manual headers above
  const headerRow = ["Account", ...projectCols.map(col => col.title)];
  const startRow  = 6;
  const startCol  = 4;   // Column D
  const numCols   = headerRow.length;

  // Clear the old data range before writing
  const lastRow   = Math.max(sh.getLastRow(), startRow);
  const clearCols = Math.max(sh.getLastColumn() - startCol + 1, numCols);
  sh.getRange(startRow, startCol, lastRow - startRow + 1, clearCols).clearContent();

  // Write header and data
  sh.getRange(startRow,     startCol, 1,              numCols).setValues([headerRow]);
  sh.getRange(startRow + 1, startCol, dataRows.length, numCols).setValues(dataRows);

  // Bold the header row and the account label column
  sh.getRange(startRow,     startCol, 1,              numCols).setFontWeight("bold");
  sh.getRange(startRow + 1, startCol, dataRows.length, 1      ).setFontWeight("bold");

  ui.alert(
    `Pulled ${projectCols.length} projects and ${dataRows.length} rows into "${sheetName}"`
  );
}

/**
 * Checks whether all transactions in the date range are assigned to a class.
 *
 * How it works:
 *  - Fetches the total net income for the full period (no class filter)
 *  - Fetches net income filtered to class A and class B separately
 *  - If Class A + Class B ≠ Total, some transactions are unclassified
 *
 * @param {string} startDate  - Report start date (YYYY-MM-DD)
 * @param {string} endDate    - Report end date (YYYY-MM-DD)
 * @param {string} classIdA   - QBO class ID for your first class
 * @param {string} classIdB   - QBO class ID for your second class
 * @returns {boolean} true if all transactions are classified, false otherwise
 */
function checkUnclassifiedTransactions_(startDate, endDate, classIdA, classIdB) {
  const ui         = SpreadsheetApp.getUi();
  const baseParams = { start_date: startDate, end_date: endDate, accounting_method: "Accrual" };

  // Helper: extract the final net income value from a P&L report response
  function getNetIncome(report) {
    const sections  = report?.Rows?.Row || [];
    const profitRow = sections.find(
      r => r?.Summary?.ColData?.[0]?.value === "PROFIT"
    );
    const cols = profitRow?.Summary?.ColData || [];
    return parseFloat(cols[cols.length - 1]?.value || "0") || 0;
  }

  const totalNet    = getNetIncome(qboFetch_("reports/ProfitAndLoss", baseParams));
  const classANet   = getNetIncome(qboFetch_("reports/ProfitAndLoss", { ...baseParams, class: classIdA }));
  const classBNet   = getNetIncome(qboFetch_("reports/ProfitAndLoss", { ...baseParams, class: classIdB }));

  // Round to 2 decimal places to avoid floating point noise
  const unclassified = Math.round((totalNet - classANet - classBNet) * 100) / 100;

  const allGood = unclassified === 0;

  ui.alert(
    `Classification Check: ${startDate} → ${endDate}\n\n` +
    `Total Net Income:  $${totalNet}\n` +
    `Class A:           $${classANet}\n` +
    `Class B:           $${classBNet}\n` +
    `Unclassified:      $${unclassified}\n\n` +
    (allGood
      ? "✅ All transactions classified. Continuing..."
      : "⚠️ Unclassified transactions found — fix in QBO before proceeding.")
  );

  return allGood;
}


// ── SHARED UTILITY ─────────────────────────────────────────────────────────────

/**
 * Writes a header row and data rows to a named sheet.
 * Creates the sheet if it doesn't already exist, and clears it before writing.
 *
 * @param {string}   sheetName - Name of the target sheet tab
 * @param {Array}    header    - Array of column header strings
 * @param {Array[]}  rows      - 2D array of data rows
 */
function writeToSheet_(sheetName, header, rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  sh.clearContents();
  sh.getRange(1, 1, 1, header.length).setValues([header]);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
}
