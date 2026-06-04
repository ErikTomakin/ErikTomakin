// ============================================================
// START: FILE HEADER — GSHEET SHOP EXPENSE TRACKER
// ============================================================
// GSheet Shop — AI-Ready Expense Tracker for Gig Workers
// Colors: SJSU Gold & Dark Blue
// ============================================================
// END: FILE HEADER
// ============================================================


// ============================================================
// START: COLOR CONSTANTS
// ============================================================

var DARK_BLUE  = "#0055A2";
var GOLD       = "#E5A823";
var LIGHT_GOLD = "#FDF3D7";
var LIGHT_BLUE = "#E8F0FA";

// ============================================================
// END: COLOR CONSTANTS
// ============================================================


// ============================================================
// START: TRIGGERS — onOpen / onEdit
// ============================================================

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getUi()
    .createMenu("Expense Tracker")
    .addItem("Add Transaction", "addTransaction")
    .addSeparator()
    .addItem("1 — Create Tabs",         "runCreateSheets")
    .addItem("2 — Setup Transactions",  "runSetupTransactions")
    .addItem("3 — Setup Tab",           "runSetupTab")
    .addItem("4 — Rebuild Dashboard",   "runRebuildDashboard")
    .addItem("5 — Setup AI Prompts",    "runSetupPrompts")
    .addItem("6 — Setup Instructions",  "runSetupInstructions")
    .addItem("7 — Setup Reports",       "runSetupReports")
    .addItem("8 — Mobile Entry Tab",    "runSetupMobileEntry")
    .addSeparator()
    .addItem("Install Triggers (run once)", "installTriggers")
    .addItem("Refresh Dropdowns",       "refreshDropdowns")
    .addItem("Clear All Transactions",  "clearAllTransactions")
    .addToUi();

  checkIRSRate();
}

function installTriggers() {
  // Remove any existing handleEdit installable triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "handleEdit") ScriptApp.deleteTrigger(t);
  });

  // Install a new authorized onEdit trigger
  ScriptApp.newTrigger("handleEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert("Triggers installed!\n\nCheckboxes on Dashboard and Mobile Entry will now work.");
}

function handleEdit(e) {
  var sheet    = e.range.getSheet();
  var row      = e.range.getRow();
  var col      = e.range.getColumn();
  var val      = e.value;
  var isChecked = (val === true || val === "TRUE");

  // Dashboard: checkbox in A8 triggers addTransaction
  if (sheet.getName() === "Dashboard" && row === 8 && col === 1 && isChecked) {
    addTransaction();
    return;
  }

  // Mobile Entry: checkbox in A14 triggers addTransactionFromEntry_
  if (sheet.getName() === "Mobile Entry" && row === 14 && col === 1 && isChecked) {
    addTransactionFromEntry_();
    return;
  }

  // Transactions tab: warn before editing cols A-G unless Edit? checkbox (col H) is checked
  if (sheet.getName() === "Transactions" && col >= 1 && col <= 7 && row > 1) {
    var editCheckbox = sheet.getRange(row, 8).getValue();
    if (editCheckbox !== true) {
      var ui       = SpreadsheetApp.getUi();
      var response = ui.alert("Edit Saved Transaction?", "You're about to edit a saved transaction. Continue?", ui.ButtonSet.YES_NO);
      if (response !== ui.Button.YES) {
        if (e.oldValue !== undefined) {
          e.range.setValue(e.oldValue);
        } else {
          e.range.clearContent();
        }
      }
      return;
    }
  }

  // Transactions tab: unchecking Edit? re-locks the row
  if (sheet.getName() === "Transactions" && col === 8 && row > 1) {
    if (!isChecked) {
      var prot = sheet.getRange(row, 1, 1, 7).protect().setDescription("tx_row_" + row);
      prot.removeEditors(prot.getEditors());
    } else {
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var i = 0; i < protections.length; i++) {
        if (protections[i].getDescription() === "tx_row_" + row) {
          protections[i].remove();
          break;
        }
      }
    }
  }
}

// ============================================================
// END: TRIGGERS — onOpen / onEdit
// ============================================================


// ============================================================
// START: IRS RATE CHECK
// ============================================================

function checkIRSRate() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var setup = ss.getSheetByName("Setup");
  if (!setup) return;

  var data = setup.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).indexOf("Business Mileage") !== -1) {
      var lastUpdated = data[i][3];
      if (lastUpdated) {
        var year        = new Date(lastUpdated).getFullYear();
        var currentYear = new Date().getFullYear();
        if (year < currentYear) {
          setup.setTabColor("red");
          SpreadsheetApp.getUi().alert(
            "IRS Mileage Rate may be outdated!\n\n" +
            "Please visit irs.gov and update the Business Mileage rate\n" +
            "in the Setup tab for " + currentYear + "."
          );
        }
      }
      break;
    }
  }
}

// ============================================================
// END: IRS RATE CHECK
// ============================================================


// ============================================================
// START: DYNAMIC SETUP READERS — getIncomeSources / getCategories
// ============================================================

function getIncomeSources(ss) {
  var setup = ss.getSheetByName("Setup");
  if (!setup) return [];
  var data      = setup.getDataRange().getValues();
  var sources   = [];
  var inSection = false;
  for (var i = 0; i < data.length; i++) {
    var cell = String(data[i][0]).trim();
    if (cell === "INCOME SOURCES") { inSection = true; continue; }
    if (inSection) {
      if (cell !== "" && cell === cell.toUpperCase()) break; // hit next section header, stop
      if (cell.indexOf("\n") !== -1) break;                 // multiline = instruction note, stop
      if (cell !== "") sources.push(cell);                  // skip blanks, collect valid entries
    }
  }
  return sources;
}

function getCategories(ss) {
  var setup = ss.getSheetByName("Setup");
  if (!setup) return [];
  var data      = setup.getDataRange().getValues();
  var cats      = [];
  var inSection = false;
  for (var i = 0; i < data.length; i++) {
    var cell = String(data[i][0]).trim();
    if (cell === "INCOME & EXPENSE CATEGORIES") { inSection = true; continue; }
    if (inSection) {
      if (cell !== "" && cell === cell.toUpperCase()) break; // hit next section header, stop
      if (cell !== "") cats.push(cell);                      // skip blanks, collect valid entries
    }
  }
  return cats;
}

// ============================================================
// END: DYNAMIC SETUP READERS — getIncomeSources / getCategories
// ============================================================


// ============================================================
// START: REFRESH DROPDOWNS
// ============================================================

// Private: applies data validation rules only — no alert, no dashboard rebuild
function applyDropdownValidations_(ss, sources, cats) {
  var txSheet = ss.getSheetByName("Transactions");
  var dash    = ss.getSheetByName("Dashboard");
  if (!txSheet || !dash) return;

  var types    = ["Income", "Expense"];
  var lastRow  = 1000;
  var srcRule  = SpreadsheetApp.newDataValidation().requireValueInList(sources, true).setAllowInvalid(false).build();
  var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(types, true).setAllowInvalid(false).build();
  var catRule  = SpreadsheetApp.newDataValidation().requireValueInList(cats, true).setAllowInvalid(false).build();

  txSheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(srcRule);
  txSheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(typeRule);
  txSheet.getRange(2, 5, lastRow - 1, 1).setDataValidation(catRule);
  dash.getRange("B3").setDataValidation(srcRule);
  dash.getRange("C3").setDataValidation(typeRule);
  dash.getRange("A5").setDataValidation(catRule);

  // Quick Entry tab dropdowns (B4 = Source, B5 = Type, B7 = Category)
  var entry = ss.getSheetByName("Mobile Entry");
  if (entry) {
    entry.getRange("B4").setDataValidation(srcRule);
    entry.getRange("B5").setDataValidation(typeRule);
    entry.getRange("B7").setDataValidation(catRule);
  }
}

// Private: surgically rebuilds the Income by Source rows on the dashboard
function rebuildDashboardSourceTable_(ss, sources) {
  var dash = ss.getSheetByName("Dashboard");
  if (!dash) return;

  var DATA_START_ROW = 13; // first source data row (set by runRebuildDashboard)

  var lastRow = dash.getLastRow();
  if (lastRow < DATA_START_ROW) return; // dashboard not set up yet

  // Find the first TOTAL row at or after row 13
  var colAVals = dash.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  var firstTotalRow = null;
  for (var i = 0; i < colAVals.length; i++) {
    if (String(colAVals[i][0]).trim() === "TOTAL") {
      firstTotalRow = DATA_START_ROW + i;
      break;
    }
  }
  if (firstTotalRow === null) return; // dashboard not built properly

  var oldSourceCount = firstTotalRow - DATA_START_ROW;
  var newSourceCount = sources.length;

  // Adjust row count if needed — delete/insert from the end of the new data block
  if (newSourceCount < oldSourceCount) {
    dash.deleteRows(DATA_START_ROW + newSourceCount, oldSourceCount - newSourceCount);
  } else if (newSourceCount > oldSourceCount) {
    dash.insertRowsBefore(firstTotalRow, newSourceCount - oldSourceCount);
  }

  // Rewrite source data rows
  var dashRow = DATA_START_ROW;
  sources.forEach(function(src, idx) {
    var bg = idx % 2 === 0 ? LIGHT_BLUE : "white";
    dash.getRange(dashRow, 1, 1, 4).clearContent().clearFormat();
    dash.getRange(dashRow, 1).setValue(src).setBackground(bg);
    dash.getRange(dashRow, 2)
      .setFormula('=IFERROR(SUMPRODUCT((Transactions!B$2:B$1000="' + src + '")*(Transactions!C$2:C$1000="Income")*Transactions!D$2:D$1000),0)')
      .setBackground(bg).setNumberFormat("$#,##0.00");
    dash.getRange(dashRow, 3)
      .setFormula('=IFERROR(SUMPRODUCT((Transactions!B$2:B$1000="' + src + '")*(Transactions!C$2:C$1000="Expense")*ABS(Transactions!D$2:D$1000)),0)')
      .setBackground(bg).setNumberFormat("$#,##0.00");
    dash.getRange(dashRow, 4)
      .setFormula("=" + columnLetter(2, dashRow) + "-" + columnLetter(3, dashRow))
      .setBackground(bg).setNumberFormat("$#,##0.00");
    dashRow++;
  });

  // Rewrite TOTAL row
  var newTotalRow = DATA_START_ROW + newSourceCount;
  dash.getRange(newTotalRow, 1, 1, 4).clearContent().clearFormat();
  dash.getRange(newTotalRow, 1).setValue("TOTAL")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold");
  dash.getRange(newTotalRow, 2)
    .setFormula("=SUM(B" + DATA_START_ROW + ":B" + (newTotalRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  dash.getRange(newTotalRow, 3)
    .setFormula("=SUM(C" + DATA_START_ROW + ":C" + (newTotalRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  dash.getRange(newTotalRow, 4)
    .setFormula("=SUM(D" + DATA_START_ROW + ":D" + (newTotalRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
}

// Public: Expense Tracker → Refresh Dropdowns
function refreshDropdowns() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var txSheet = ss.getSheetByName("Transactions");
  var dash    = ss.getSheetByName("Dashboard");
  if (!txSheet || !dash) {
    SpreadsheetApp.getUi().alert("Tabs not found. Run setup steps first.");
    return;
  }

  var sources = getIncomeSources(ss);
  var cats    = getCategories(ss);

  if (sources.length === 0) sources = ["DoorDash", "Uber", "Freelance", "Notary", "Other"];
  if (cats.length === 0)    cats    = ["Delivery Income", "Fuel", "Other Expense"];

  rebuildDashboardSourceTable_(ss, sources);
  applyDropdownValidations_(ss, sources, cats);

  SpreadsheetApp.getUi().alert("Dropdowns and dashboard updated!");
}

// ============================================================
// END: REFRESH DROPDOWNS
// ============================================================


// ============================================================
// START: TRANSACTION LOGIC — addTransaction / getMileageRate / clearForm / clearAllTransactions
// ============================================================

function addTransaction() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");
  var tx   = ss.getSheetByName("Transactions");
  if (!dash || !tx) {
    SpreadsheetApp.getUi().alert("Required tabs not found.");
    return;
  }

  // Flash button green to show script is running
  ss.toast("Processing your transaction...", "Please wait", 10);
  dash.getRange("A8:D9").setBackground("#2E7D32");
  dash.getRange("B8:D8").setFontColor("white");
  SpreadsheetApp.flush();

  var date     = dash.getRange("A3").getValue();
  var source   = dash.getRange("B3").getValue();
  var type     = dash.getRange("C3").getValue();
  var amount   = dash.getRange("D3").getValue();
  var category = dash.getRange("A5").getValue();
  var miles    = dash.getRange("B5").getValue();
  var hours    = dash.getRange("C5").getValue();
  var notes    = dash.getRange("D5").getValue();

  if (!source || !type) {
    SpreadsheetApp.getUi().alert("Please fill in at least Income Source and Type before submitting.");
    return;
  }

  if (!amount && (!miles || miles <= 0)) {
    SpreadsheetApp.getUi().alert("Please enter an Amount or Miles Driven — nothing to record.");
    return;
  }

  // Default date to today if left blank
  if (!date) date = new Date();

  // Only post a main transaction row if there is an amount
  if (amount) {
    var nextRow = tx.getLastRow() + 1;
    tx.getRange(nextRow, 1, 1, 7).setValues([[date, source, type, amount, category, hours, notes]]);
    tx.getRange(nextRow, 8).insertCheckboxes();
    var prot1 = tx.getRange(nextRow, 1, 1, 7).protect().setDescription("tx_row_" + nextRow);
    prot1.removeEditors(prot1.getEditors());
  }

  // Auto-post mileage deduction row if miles were entered
  if (miles && miles > 0) {
    var rate = getMileageRate(ss);
    if (rate === null) {
      SpreadsheetApp.getUi().alert("WARNING: Could not find IRS mileage rate in Setup tab.\nMileage deduction was NOT recorded.");
    } else {
      var deduction = miles * rate;
      var mileRow   = tx.getLastRow() + 1;
      tx.getRange(mileRow, 1, 1, 7).setValues([
        [date, source, "Expense", -deduction, "Mileage Deduction", "", miles + " miles @ $" + rate]
      ]);
      tx.getRange(mileRow, 8).insertCheckboxes();
      var prot2 = tx.getRange(mileRow, 1, 1, 7).protect().setDescription("tx_row_" + mileRow);
      prot2.removeEditors(prot2.getEditors());
    }
  }

  clearForm();
  SpreadsheetApp.getUi().alert("Transaction added!");
}

// ------------------------------------------------------------

function getMileageRate(ss) {
  var setup = ss.getSheetByName("Setup");
  if (!setup) return null;
  var data = setup.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).indexOf("Business Mileage") !== -1) {
      var rate = parseFloat(data[i][1]);
      return isNaN(rate) ? null : rate;
    }
  }
  return null;
}

// ------------------------------------------------------------

function clearForm() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");
  if (!dash) return;

  dash.getRange("A3").setValue(new Date());
  dash.getRange("B3").clearContent();
  dash.getRange("C3").clearContent();
  dash.getRange("D3").clearContent();
  dash.getRange("A5").clearContent();
  dash.getRange("B5").clearContent();
  dash.getRange("C5").clearContent();
  dash.getRange("D5").clearContent();
  dash.getRange("A8").setValue(false);
  dash.getRange("A7:D9").setBackground(DARK_BLUE);
  dash.getRange("B8:D8").setFontColor(GOLD);
  SpreadsheetApp.flush();
}

// ------------------------------------------------------------

function clearAllTransactions() {
  var ui       = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Clear All Transactions",
    "This will permanently delete all transaction data.\nType YES to confirm:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  if (response.getResponseText().trim().toUpperCase() !== "YES") {
    ui.alert("Cancelled. No data was deleted.");
    return;
  }
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tx      = ss.getSheetByName("Transactions");
  if (!tx) return;

  // Remove all row protections before deleting
  var protections = tx.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    protections[i].remove();
  }

  var lastRow = tx.getLastRow();
  if (lastRow > 1) tx.deleteRows(2, lastRow - 1);
  ui.alert("All transactions cleared.");
}

// ------------------------------------------------------------

function addTransactionFromEntry_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var entry = ss.getSheetByName("Mobile Entry");
  var tx    = ss.getSheetByName("Transactions");
  if (!entry || !tx) {
    SpreadsheetApp.getUi().alert("Required tabs not found.");
    return;
  }

  ss.toast("Processing your transaction...", "Please wait", 10);
  entry.getRange("A13:C15").setBackground("#2E7D32");
  entry.getRange("B14:C14").setFontColor("white");
  SpreadsheetApp.flush();

  var date     = entry.getRange("B3").getValue();
  var source   = entry.getRange("B4").getValue();
  var type     = entry.getRange("B5").getValue();
  var amount   = entry.getRange("B6").getValue();
  var category = entry.getRange("B7").getValue();
  var miles    = entry.getRange("B8").getValue();
  var hours    = entry.getRange("B9").getValue();
  var notes    = entry.getRange("B10").getValue();

  if (!source || !type) {
    SpreadsheetApp.getUi().alert("Please fill in at least Income Source and Type before submitting.");
    clearEntryForm_();
    return;
  }
  if (!amount && (!miles || miles <= 0)) {
    SpreadsheetApp.getUi().alert("Please enter an Amount or Miles Driven — nothing to record.");
    clearEntryForm_();
    return;
  }
  if (!date) date = new Date();

  if (amount) {
    var nextRow = tx.getLastRow() + 1;
    tx.getRange(nextRow, 1, 1, 7).setValues([[date, source, type, amount, category, hours, notes]]);
    tx.getRange(nextRow, 8).insertCheckboxes();
    var prot1 = tx.getRange(nextRow, 1, 1, 7).protect().setDescription("tx_row_" + nextRow);
    prot1.removeEditors(prot1.getEditors());
  }

  if (miles && miles > 0) {
    var rate = getMileageRate(ss);
    if (rate === null) {
      SpreadsheetApp.getUi().alert("WARNING: Could not find IRS mileage rate in Setup tab.\nMileage deduction was NOT recorded.");
    } else {
      var deduction = miles * rate;
      var mileRow   = tx.getLastRow() + 1;
      tx.getRange(mileRow, 1, 1, 7).setValues([
        [date, source, "Expense", -deduction, "Mileage Deduction", "", miles + " miles @ $" + rate]
      ]);
      tx.getRange(mileRow, 8).insertCheckboxes();
      var prot2 = tx.getRange(mileRow, 1, 1, 7).protect().setDescription("tx_row_" + mileRow);
      prot2.removeEditors(prot2.getEditors());
    }
  }

  clearEntryForm_();
  SpreadsheetApp.getUi().alert("Transaction added!");
}

// ------------------------------------------------------------

function clearEntryForm_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var entry = ss.getSheetByName("Mobile Entry");
  if (!entry) return;
  entry.getRange("B3").setValue(new Date());
  entry.getRange("B4").clearContent();
  entry.getRange("B5").clearContent();
  entry.getRange("B6").clearContent();
  entry.getRange("B7").clearContent();
  entry.getRange("B8").clearContent();
  entry.getRange("B9").clearContent();
  entry.getRange("B10").clearContent();
  entry.getRange("A14").setValue(false);
  entry.getRange("A13:C15").setBackground(DARK_BLUE);
  entry.getRange("B14:C14").setFontColor(GOLD);
  SpreadsheetApp.flush();
}

// ============================================================
// END: TRANSACTION LOGIC — addTransaction / getMileageRate / clearForm / clearAllTransactions
// ============================================================


// ============================================================
// START: SETUP STEP 1 — CREATE TABS
// ============================================================

function runCreateSheets() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var tabNames = ["Dashboard", "Mobile Entry", "Transactions", "Setup", "AI Prompts", "Instructions", "Reports"];
  var colors   = [GOLD, DARK_BLUE, GOLD, DARK_BLUE, GOLD, DARK_BLUE, GOLD];

  // Remove legacy tabs
  ["Sources", "Expenses"].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  tabNames.forEach(function(name, idx) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.setTabColor(colors[idx]);
  });

  SpreadsheetApp.getUi().alert("Tabs created! Run Step 2 next.");
}

// ============================================================
// END: SETUP STEP 1 — CREATE TABS
// ============================================================


// ============================================================
// START: SETUP STEP 2 — TRANSACTIONS TAB
// ============================================================

function runSetupTransactions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tx = ss.getSheetByName("Transactions");
  if (!tx) { SpreadsheetApp.getUi().alert("Run Step 1 first."); return; }

  tx.clearContents();
  tx.clearFormats();

  var headers = ["Date", "Source", "Type", "Amount", "Category", "Hours", "Notes", "Edit?"];
  var hRange  = tx.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers])
    .setBackground(DARK_BLUE)
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Alternating row colors
  for (var r = 2; r <= 1000; r++) {
    tx.getRange(r, 1, 1, headers.length)
      .setBackground(r % 2 === 0 ? LIGHT_BLUE : "white");
  }

  // Column widths
  tx.setColumnWidth(1, 100);
  tx.setColumnWidth(2, 140);
  tx.setColumnWidth(3, 90);
  tx.setColumnWidth(4, 100);
  tx.setColumnWidth(5, 160);
  tx.setColumnWidth(6, 80);
  tx.setColumnWidth(7, 220);
  tx.setColumnWidth(8, 60);

  // Format date, currency, and hours
  tx.getRange(2, 1, 999, 1).setNumberFormat("MM/dd/yyyy");
  tx.getRange(2, 4, 999, 1).setNumberFormat("$#,##0.00");
  tx.getRange(2, 6, 999, 1).setNumberFormat("0.00");

  // Wrap text in Notes column (G)
  tx.getRange(2, 7, 999, 1).setWrap(true);

  // Hide all columns beyond H so users can't wander off the edge
  var maxCols = tx.getMaxColumns();
  if (maxCols > 8) tx.hideColumns(9, maxCols - 8);

  tx.setFrozenRows(1);

  // Apply filter arrows to header row (A–H) so Date, Source, Type, Category are sortable/filterable
  var existingFilter = tx.getFilter();
  if (existingFilter) existingFilter.remove();
  tx.getRange(1, 1, 1, 8).createFilter();

  SpreadsheetApp.getUi().alert("Transactions tab ready! Run Step 3 next.");
}

// ============================================================
// END: SETUP STEP 2 — TRANSACTIONS TAB
// ============================================================


// ============================================================
// START: SETUP STEP 3 — SETUP TAB
// ============================================================

function runSetupTab() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var setup = ss.getSheetByName("Setup");
  if (!setup) { SpreadsheetApp.getUi().alert("Run Step 1 first."); return; }

  setup.clearContents();
  setup.clearFormats();
  setup.clearNotes();
  setup.getRange(1, 1, setup.getMaxRows(), setup.getMaxColumns()).clearDataValidations();

  // Remove any existing protections on this sheet
  var existingProts = setup.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var p = 0; p < existingProts.length; p++) existingProts[p].remove();

  var row        = 1;
  var headerRows = []; // track rows to lock at the end

  // -- MAIN HEADER --
  setup.getRange(row, 1, 1, 4).merge()
    .setValue("SETUP & CONFIGURATION")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center");
  setup.setRowHeight(row, 40);
  headerRows.push(row);
  row++;

  // -- INCOME SOURCES SECTION --
  row++;
  setup.getRange(row, 1, 1, 4).merge()
    .setValue("INCOME SOURCES")
    .setBackground(GOLD).setFontColor(DARK_BLUE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  setup.setRowHeight(row, 32);
  headerRows.push(row);
  row++;

  setup.getRange(row, 1, 1, 4).setValues([["Source Name", "Description", "", ""]])
    .setBackground(DARK_BLUE).setFontColor("white").setFontWeight("bold");
  headerRows.push(row);
  row++;

  var incomeSources = [
    "DoorDash",
    "Uber Eats",
    "Instacart",
    "Uber",
    "Lyft",
    "Etsy",
    "Upwork",
    "Notary Work",
    "Other"
  ];
  incomeSources.forEach(function(src) {
    setup.getRange(row, 1).setValue(src);
    setup.getRange(row, 1, 1, 4).setBackground(LIGHT_GOLD);
    row++;
  });

  // Small spacer between last source and instruction note
  setup.setRowHeight(row, 8);
  row++;

  // Instruction note (locked)
  setup.getRange(row, 1, 1, 4).merge()
    .setValue(
      "HOW TO EDIT INCOME SOURCES:\n" +
      "• ADD: Type a new source in any empty row above, then run Expense Tracker → Refresh Dropdowns.\n" +
      "• REMOVE: Click the source name and press Delete to clear it, then run Refresh Dropdowns.\n" +
      "• RENAME: Click the source name and type the new name (e.g. Etsy → Mayhem Marketplace), then run Refresh Dropdowns.\n" +
      "Note: Renaming only affects new entries. Old transactions will still show the previous name."
    )
    .setBackground("#FFF9C4").setFontColor("#666666").setFontSize(9).setFontStyle("italic")
    .setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true);
  setup.setRowHeight(row, 80);
  headerRows.push(row);
  row++;

  // -- INCOME & EXPENSE CATEGORIES SECTION --
  row++;
  setup.getRange(row, 1, 1, 4).merge()
    .setValue("INCOME & EXPENSE CATEGORIES")
    .setBackground(GOLD).setFontColor(DARK_BLUE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  setup.setRowHeight(row, 32);
  headerRows.push(row);
  row++;

  setup.getRange(row, 1, 1, 4).setValues([["Category Name", "Type", "", ""]])
    .setBackground(DARK_BLUE).setFontColor("white").setFontWeight("bold");
  headerRows.push(row);
  row++;

  var categories = [
    ["Base",                  "Income"],
    ["Bonus",                 "Income"],
    ["Tips",                  "Income"],
    ["Sales",                 "Income"],
    ["Other Income",          "Income"],
    ["Mileage Deduction",     "Expense"],
    ["Fuel",                  "Expense"],
    ["Vehicle Maintenance",   "Expense"],
    ["Car Insurance",         "Expense"],
    ["Parking & Tolls",       "Expense"],
    ["Shipping",              "Expense"],
    ["Packaging",             "Expense"],
    ["Platform Fees",         "Expense"],
    ["Product Supplies",      "Expense"],
    ["Software Subscriptions","Expense"],
    ["Printing & Supplies",   "Expense"],
    ["Professional Fees",     "Expense"],
    ["Phone Bill",            "Expense"],
    ["Internet",              "Expense"],
    ["Home Office",           "Expense"],
    ["Advertising",           "Expense"],
    ["Recurring Expense",     "Expense"],
    ["Other Expense",         "Expense"]
  ];

  categories.forEach(function(cat) {
    setup.getRange(row, 1).setValue(cat[0]);
    setup.getRange(row, 2).setValue(cat[1]);
    var bg = cat[1] === "Income" ? LIGHT_GOLD : LIGHT_BLUE;
    setup.getRange(row, 1, 1, 4).setBackground(bg);
    row++;
  });

  // -- RECURRING EXPENSES SECTION --
  row++;
  setup.getRange(row, 1, 1, 4).merge()
    .setValue("RECURRING EXPENSES (Reminder List)")
    .setBackground(GOLD).setFontColor(DARK_BLUE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  setup.setRowHeight(row, 32);
  headerRows.push(row);
  row++;

  setup.getRange(row, 1, 1, 4).setValues([["Expense", "Amount", "Frequency", "Next Due Date"]])
    .setBackground(DARK_BLUE).setFontColor("white").setFontWeight("bold");
  headerRows.push(row);
  row++;

  var recurring = [
    ["Phone Bill",             "$100", "Monthly", ""],
    ["Car Insurance",          "$150", "Monthly", ""],
    ["Software Subscriptions", "$30",  "Monthly", ""],
    ["Internet",               "$60",  "Monthly", ""]
  ];
  recurring.forEach(function(rec) {
    setup.getRange(row, 1, 1, 4).setValues([rec]).setBackground(LIGHT_BLUE);
    row++;
  });

  // -- IRS MILEAGE RATE SECTION --
  row++;
  setup.getRange(row, 1, 1, 4).merge()
    .setValue("IRS MILEAGE RATE")
    .setBackground(GOLD).setFontColor(DARK_BLUE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  setup.setRowHeight(row, 32);
  headerRows.push(row);
  row++;

  setup.getRange(row, 1, 1, 4).setValues([["Description", "Rate (per mile)", "Effective Date", "Last Updated"]])
    .setBackground(DARK_BLUE).setFontColor("white").setFontWeight("bold");
  headerRows.push(row);
  row++;

  setup.getRange(row, 1, 1, 4).setValues([
    ["Business Mileage", 0.725, "01/01/2026", new Date()]
  ]).setBackground(LIGHT_GOLD);
  setup.getRange(row, 2).setNumberFormat("$0.000");

  // Column widths
  setup.setColumnWidth(1, 200);
  setup.setColumnWidth(2, 150);
  setup.setColumnWidth(3, 150);
  setup.setColumnWidth(4, 150);

  // Hide all columns beyond D
  var maxCols = setup.getMaxColumns();
  if (maxCols > 4) setup.hideColumns(5, maxCols - 4);

  // Lock all header/banner rows
  headerRows.forEach(function(r) {
    var prot = setup.getRange(r, 1, 1, 4).protect().setDescription("setup_header_" + r);
    prot.removeEditors(prot.getEditors());
  });

  var sources = getIncomeSources(ss);
  var cats    = getCategories(ss);
  if (sources.length === 0) sources = ["DoorDash", "Uber", "Freelance", "Notary", "Other"];
  if (cats.length === 0)    cats    = ["Delivery Income", "Fuel", "Other Expense"];
  applyDropdownValidations_(ss, sources, cats);
  SpreadsheetApp.getUi().alert("Setup tab ready! Run Step 4 next.");
}

// ============================================================
// END: SETUP STEP 3 — SETUP TAB
// ============================================================


// ============================================================
// START: SETUP STEP 4 — DASHBOARD
// ============================================================

function runRebuildDashboard() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");
  if (!dash) { SpreadsheetApp.getUi().alert("Run Step 1 first."); return; }

  dash.clearContents();
  dash.clearFormats();
  dash.clearNotes();
  try { dash.removeCheckboxes(); } catch(e) {}
  dash.getRange("A1:D20").clearDataValidations();

  var sources = getIncomeSources(ss);
  if (sources.length === 0) sources = ["DoorDash", "Uber", "Freelance", "Notary", "Other"];

  // -- TRANSACTION ENTRY FORM (ROWS 1-9) --

  // Row 1: Main header
  dash.getRange("A1:D1").merge()
    .setValue("ADD NEW TRANSACTION")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center");
  dash.setRowHeight(1, 40);

  // Row 2: Field labels (top row)
  dash.getRange("A2:D2").setValues([["Date", "Income Source", "Type", "Amount"]])
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Row 3: Input fields (top row)
  dash.getRange("A3").setValue(new Date()).setNumberFormat("MM/dd/yyyy")
    .setBackground("white").setBorder(true, true, true, true, false, false);
  dash.getRange("B3").setBackground("white").setBorder(true, true, true, true, false, false);
  dash.getRange("C3").setBackground("white").setBorder(true, true, true, true, false, false);
  dash.getRange("D3").setBackground("white").setNumberFormat("$#,##0.00")
    .setBorder(true, true, true, true, false, false);

  // Row 4: Field labels (bottom row)
  dash.getRange("A4:D4").setValues([["Category", "Miles Driven", "Hours Worked", "Notes"]])
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Row 5: Input fields (bottom row)
  dash.getRange("A5").setBackground("white").setBorder(true, true, true, true, false, false);
  dash.getRange("B5").setBackground("white").setNumberFormat("0.0")
    .setBorder(true, true, true, true, false, false);
  dash.getRange("C5").setBackground("white").setNumberFormat("0.0")
    .setBorder(true, true, true, true, false, false);
  dash.getRange("D5").setBackground("white").setBorder(true, true, true, true, false, false);

  // Row 6: Mileage auto-calc note
  dash.getRange("A6:D6").merge()
    .setValue("Mileage deduction auto-calculates using the IRS rate in the Setup tab.")
    .setBackground(LIGHT_GOLD).setFontColor("#666666").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dash.setRowHeight(6, 24);

  // Row 7: Top border of submit button
  dash.getRange("A7:D7").merge().setBackground(DARK_BLUE);
  dash.setRowHeight(7, 6);

  // Row 8: SUBMIT button — checkbox in A8 activates it, label in B8:D8 tells user what to do
  dash.getRange("A8").insertCheckboxes().setValue(false)
    .setBackground(DARK_BLUE);
  dash.getRange("B8:D8").merge()
    .setValue("CHECK THE BOX TO SUBMIT  ✔  (Allow 2-3 seconds)")
    .setBackground(DARK_BLUE).setFontColor(GOLD)
    .setFontWeight("bold").setFontSize(13)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dash.setRowHeight(8, 50);

  // Row 9: Bottom border of submit button
  dash.getRange("A9:D9").merge().setBackground(DARK_BLUE);
  dash.setRowHeight(9, 6);

  // Row 10: Spacer
  dash.setRowHeight(10, 16);

  // -- INCOME BY SOURCE TABLE (STARTS ROW 11) --

  var dashRow = 11;

  dash.getRange(dashRow, 1, 1, 4).merge()
    .setValue("INCOME BY SOURCE — YEAR TO DATE")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  dash.setRowHeight(dashRow, 32);
  dashRow++;

  dash.getRange(dashRow, 1, 1, 4).setValues([["Source", "Income", "Expenses", "Net"]])
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold")
    .setHorizontalAlignment("center");
  dashRow++;

  sources.forEach(function(src, idx) {
    var bg = idx % 2 === 0 ? LIGHT_BLUE : "white";
    dash.getRange(dashRow, 1).setValue(src).setBackground(bg);
    dash.getRange(dashRow, 2).setFormula(
      '=IFERROR(SUMPRODUCT((Transactions!B$2:B$1000="' + src + '")*(Transactions!C$2:C$1000="Income")*Transactions!D$2:D$1000),0)'
    ).setBackground(bg).setNumberFormat("$#,##0.00");
    dash.getRange(dashRow, 3).setFormula(
      '=IFERROR(SUMPRODUCT((Transactions!B$2:B$1000="' + src + '")*(Transactions!C$2:C$1000="Expense")*ABS(Transactions!D$2:D$1000)),0)'
    ).setBackground(bg).setNumberFormat("$#,##0.00");
    dash.getRange(dashRow, 4).setFormula(
      "=" + columnLetter(2, dashRow) + "-" + columnLetter(3, dashRow)
    ).setBackground(bg).setNumberFormat("$#,##0.00");
    dashRow++;
  });

  // Totals row for By Source
  var firstDataRow = dashRow - sources.length;
  dash.getRange(dashRow, 1).setValue("TOTAL")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold");
  dash.getRange(dashRow, 2)
    .setFormula("=SUM(B" + firstDataRow + ":B" + (dashRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  dash.getRange(dashRow, 3)
    .setFormula("=SUM(C" + firstDataRow + ":C" + (dashRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  dash.getRange(dashRow, 4)
    .setFormula("=SUM(D" + firstDataRow + ":D" + (dashRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  dashRow++;

  dashRow++; // Spacer row

  // -- MONTHLY BREAKDOWN TABLE --

  dash.getRange(dashRow, 1, 1, 4).merge()
    .setValue("MONTHLY BREAKDOWN — ALL 12 MONTHS")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  dash.setRowHeight(dashRow, 32);
  dashRow++;

  dash.getRange(dashRow, 1, 1, 4).setValues([["Month", "Income", "Expenses", "Net"]])
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold")
    .setHorizontalAlignment("center");
  dashRow++;

  var months = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  months.forEach(function(month, idx) {
    var monthNum = idx + 1;
    var bg = idx % 2 === 0 ? LIGHT_BLUE : "white";
    dash.getRange(dashRow, 1).setValue(month).setBackground(bg);
    dash.getRange(dashRow, 2).setFormula(
      '=IFERROR(SUMPRODUCT((MONTH(Transactions!A$2:A$1000)=' + monthNum + ')*(YEAR(Transactions!A$2:A$1000)=YEAR(TODAY()))*(Transactions!C$2:C$1000="Income")*Transactions!D$2:D$1000),0)'
    ).setBackground(bg).setNumberFormat("$#,##0.00");
    dash.getRange(dashRow, 3).setFormula(
      '=IFERROR(SUMPRODUCT((MONTH(Transactions!A$2:A$1000)=' + monthNum + ')*(YEAR(Transactions!A$2:A$1000)=YEAR(TODAY()))*(Transactions!C$2:C$1000="Expense")*ABS(Transactions!D$2:D$1000)),0)'
    ).setBackground(bg).setNumberFormat("$#,##0.00");
    dash.getRange(dashRow, 4).setFormula(
      "=" + columnLetter(2, dashRow) + "-" + columnLetter(3, dashRow)
    ).setBackground(bg).setNumberFormat("$#,##0.00");
    dashRow++;
  });

  // Totals row for Monthly
  var mFirstRow = dashRow - 12;
  dash.getRange(dashRow, 1).setValue("TOTAL")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold");
  dash.getRange(dashRow, 2)
    .setFormula("=SUM(B" + mFirstRow + ":B" + (dashRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  dash.getRange(dashRow, 3)
    .setFormula("=SUM(C" + mFirstRow + ":C" + (dashRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  dash.getRange(dashRow, 4)
    .setFormula("=SUM(D" + mFirstRow + ":D" + (dashRow - 1) + ")")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");

  // -- COLUMN WIDTHS & FINAL FORMATTING --

  dash.setColumnWidth(1, 160);
  dash.setColumnWidth(2, 140);
  dash.setColumnWidth(3, 140);
  dash.setColumnWidth(4, 140);

  // Hide all columns beyond D so users can't wander off the edge
  var maxCols = dash.getMaxColumns();
  if (maxCols > 4) dash.hideColumns(5, maxCols - 4);

  dash.setFrozenRows(1);

  var cats = getCategories(ss);
  if (cats.length === 0) cats = ["Delivery Income", "Fuel", "Other Expense"];
  applyDropdownValidations_(ss, sources, cats);
  SpreadsheetApp.getUi().alert("Dashboard rebuilt! Run Step 5 next.");
}

// ------------------------------------------------------------
// HELPER: columnLetter(col, row) — returns cell reference like "B14"
// ------------------------------------------------------------

function columnLetter(col, row) {
  var letters = ["", "A", "B", "C", "D", "E", "F", "G"];
  return letters[col] + row;
}

// ============================================================
// END: SETUP STEP 4 — DASHBOARD
// ============================================================


// ============================================================
// START: SETUP STEP 5 — AI PROMPTS TAB
// ============================================================

function runSetupPrompts() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var prompts = ss.getSheetByName("AI Prompts");
  if (!prompts) { SpreadsheetApp.getUi().alert("Run Step 1 first."); return; }

  prompts.clearContents();
  prompts.clearFormats();

  var row = 1;

  // -- HEADER --
  prompts.getRange(row, 1, 1, 2).merge()
    .setValue("AI PROMPTS — Copy these into ChatGPT, Claude, or any AI")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center");
  prompts.setRowHeight(row, 40);
  row++;

  // -- INSTRUCTIONS ROW --
  prompts.getRange(row, 1, 1, 2).merge()
    .setValue("Step 1: Export your data from the Transactions tab. Step 2: Copy a prompt below. Step 3: Paste both into any AI. Step 4: Paste the AI's output back here or into a note.")
    .setBackground(LIGHT_GOLD).setFontSize(10).setWrap(true)
    .setHorizontalAlignment("center");
  prompts.setRowHeight(row, 40);
  row++;

  // -- PROMPT DATA --
  var promptData = [
    {
      title: "PROMPT 1 — Monthly Tax Summary",
      prompt: "I am a gig worker with multiple income streams. Below is my transaction data for the month. Please summarize:\n" +
              "1. Total income by source\n" +
              "2. Total expenses by category\n" +
              "3. Net profit\n" +
              "4. Top 3 expense categories I could reduce\n" +
              "5. Estimated quarterly tax I should set aside (use 25-30% self-employment estimate)\n\n" +
              "[PASTE YOUR TRANSACTION DATA HERE]"
    },
    {
      title: "PROMPT 2 — Mileage & Deduction Review",
      prompt: "I am a self-employed gig worker. Here is my mileage and expense data. Please:\n" +
              "1. Confirm my mileage deduction total using the IRS standard rate\n" +
              "2. List all expense categories that are likely tax-deductible for gig workers\n" +
              "3. Flag any categories that may only be partially deductible (like phone or internet)\n" +
              "4. Suggest any deductions I may be missing based on my income sources\n\n" +
              "[PASTE YOUR TRANSACTION DATA HERE]"
    },
    {
      title: "PROMPT 3 — Annual Profit & Loss Summary",
      prompt: "I am a gig economy worker filing as self-employed. Below is my full-year transaction data. Please create:\n" +
              "1. A simple Profit & Loss summary by month\n" +
              "2. Annual totals for income, expenses, and net profit\n" +
              "3. Income breakdown by source (which platforms paid the most)\n" +
              "4. Top deductible expenses\n" +
              "5. Recommended Schedule C categories for tax filing\n\n" +
              "[PASTE YOUR TRANSACTION DATA HERE]"
    }
  ];

  promptData.forEach(function(p) {
    row++;
    prompts.getRange(row, 1, 1, 2).merge()
      .setValue(p.title)
      .setBackground(GOLD).setFontColor(DARK_BLUE)
      .setFontWeight("bold").setFontSize(12);
    prompts.setRowHeight(row, 32);
    row++;

    prompts.getRange(row, 1, 1, 2).merge()
      .setValue(p.prompt)
      .setBackground(LIGHT_BLUE).setFontSize(11)
      .setWrap(true).setVerticalAlignment("top");
    prompts.setRowHeight(row, 160);
    row++;
  });

  prompts.setColumnWidth(1, 500);
  prompts.setColumnWidth(2, 500);

  // Lock the tab — read only
  var promptProt = prompts.protect().setDescription("AI Prompts - Read Only");
  promptProt.removeEditors(promptProt.getEditors());

  SpreadsheetApp.getUi().alert("AI Prompts tab ready! Run Step 6 next.");
}

// ============================================================
// END: SETUP STEP 5 — AI PROMPTS TAB
// ============================================================


// ============================================================
// START: SETUP STEP 6 — INSTRUCTIONS TAB
// ============================================================

function runSetupInstructions() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var inst = ss.getSheetByName("Instructions");
  if (!inst) { SpreadsheetApp.getUi().alert("Run Step 1 first."); return; }

  inst.clearContents();
  inst.clearFormats();

  // Remove existing protections before re-applying
  var existingProts = inst.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var p = 0; p < existingProts.length; p++) existingProts[p].remove();

  var row = 1;

  // -- HEADER --
  inst.getRange(row, 1, 1, 2).merge()
    .setValue("HOW TO USE THIS EXPENSE TRACKER")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center");
  inst.setRowHeight(row, 50);
  row++;

  // -- INSTRUCTION SECTIONS --
  var sections = [
    {
      heading: "GETTING STARTED",
      body: "1. Go to the Setup tab and confirm your income sources and expense categories.\n" +
            "2. Update the IRS mileage rate each January (irs.gov). The Setup tab turns RED as a reminder.\n" +
            "3. Add your recurring expenses in the Recurring Expenses section of the Setup tab.\n" +
            "4. Run Expense Tracker → Refresh Dropdowns after any changes to Setup.\n" +
            "5. Run Expense Tracker → Install Triggers (run once) to activate the submit checkboxes."
    },
    {
      heading: "ADDING TRANSACTIONS — DASHBOARD",
      body: "1. Go to the Dashboard tab.\n" +
            "2. Fill in: Date, Income Source, Type (Income or Expense), and Amount.\n" +
            "3. Optionally fill in Category, Miles Driven, Hours Worked, and Notes.\n" +
            "4. Check the blue checkbox in the dark blue bar — it flashes green while processing.\n" +
            "5. Wait 2-3 seconds for the transaction to post.\n" +
            "6. If you entered mileage, a second row auto-posts your IRS mileage deduction.\n" +
            "7. You can also use Expense Tracker menu → Add Transaction."
    },
    {
      heading: "ADDING TRANSACTIONS — MOBILE ENTRY TAB",
      body: "The Mobile Entry tab is designed for phone and tablet use.\n\n" +
            "• All fields are stacked top to bottom in large 24pt text.\n" +
            "• Press Enter after each field to move to the next one.\n" +
            "• Required fields are marked with *\n" +
            "• Check the box at the bottom to post the entry — same 2-3 second processing time.\n\n" +
            "This tab works on the Google Sheets mobile app without needing desktop mode."
    },
    {
      heading: "INCOME SOURCES vs EXPENSE CATEGORIES",
      body: "• Income Sources = WHERE the money came from (DoorDash, Etsy, Upwork, etc.)\n" +
            "• Expense Categories = WHAT the money was spent on (Fuel, Phone Bill, etc.)\n" +
            "• Type 'Income' = money coming IN | Type 'Expense' = money going OUT\n" +
            "• Mileage Deduction posts as a NEGATIVE expense automatically.\n\n" +
            "To add, remove, or rename income sources:\n" +
            "• Go to the Setup tab and edit the source names under INCOME SOURCES.\n" +
            "• Run Expense Tracker → Refresh Dropdowns to update all menus."
    },
    {
      heading: "REPORTS TAB",
      body: "The Reports tab shows a live summary filtered by date range and income source.\n\n" +
            "• Set Start Date and End Date to filter by any time period.\n" +
            "• Choose a specific Income Source or leave it on All Sources.\n" +
            "• Section 1: Income & Expenses by Source\n" +
            "• Section 2: Expenses by Category\n" +
            "• Section 3: Overall Summary with net profit and estimated 28% tax set-aside.\n\n" +
            "All three sections update automatically when you change the filters."
    },
    {
      heading: "DASHBOARD SUMMARIES",
      body: "• BY SOURCE: Shows total income, expenses, and net per platform for the year.\n" +
            "• MONTHLY BREAKDOWN: Shows all 12 months of the current year.\n" +
            "• These update automatically as you add transactions.\n" +
            "• After adding or removing income sources, run Refresh Dropdowns to update the source table."
    },
    {
      heading: "USING AI (NO API NEEDED)",
      body: "1. Go to the AI Prompts tab and read the 3 prompts — decide which one fits your need.\n" +
            "2. Open ChatGPT, Claude.ai, Gemini, or any AI tool in a new browser tab.\n" +
            "3. Go back to the AI Prompts tab and copy your chosen prompt (Ctrl+C).\n" +
            "4. Switch to your AI tab and paste the prompt (Ctrl+V).\n" +
            "5. Go back to your spreadsheet and open the Transactions tab.\n" +
            "6. Copy your transaction data (Ctrl+A then Ctrl+C).\n" +
            "7. Switch back to your AI tab and paste the data below the prompt (Ctrl+V).\n" +
            "8. Send it — the AI will summarize your finances. No subscription or API key required."
    },
    {
      heading: "MENU SHORTCUTS",
      body: "Expense Tracker menu (top menu bar — not the file name):\n" +
            "• Add Transaction — posts the Dashboard form entry\n" +
            "• Refresh Dropdowns — updates all dropdowns and the dashboard source table\n" +
            "• Install Triggers (run once) — activates the submit checkboxes\n" +
            "• Clear All Transactions — wipes all rows (requires typing YES to confirm)\n" +
            "• Setup steps 1–8 — rebuilds individual tabs if needed"
    },
    {
      heading: "TIPS FOR GIG WORKERS",
      body: "• Log mileage the same day — it's easy to forget.\n" +
            "• Set aside 25-30% of net profit for quarterly taxes.\n" +
            "• Check the Reports tab and AI Prompts tab before every quarterly tax deadline.\n" +
            "• Keep receipts for all Expense entries — photos in Google Drive work great.\n" +
            "• Use 'Notary Work' or 'Freelance' as income sources for non-platform income.\n" +
            "• The IRS mileage rate changes every January — update it in the Setup tab."
    }
  ];

  // Print tip row at very top
  inst.getRange(1, 1).setValue("Print tip: fit to page width")
    .setBackground("#F0F0F0").setFontColor("#888888").setFontSize(9).setFontStyle("italic")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  inst.setRowHeight(1, 18);
  row = 2; // push all content down one row

  // Re-write header at row 2
  inst.getRange(row, 1).merge()
    .setValue("HOW TO USE THIS EXPENSE TRACKER")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center");
  inst.setRowHeight(row, 44);
  row++;

  sections.forEach(function(sec) {
    row++;
    inst.getRange(row, 1).merge()
      .setValue(sec.heading)
      .setBackground(GOLD).setFontColor(DARK_BLUE)
      .setFontWeight("bold").setFontSize(11);
    inst.setRowHeight(row, 28);
    row++;

    inst.getRange(row, 1).merge()
      .setValue(sec.body)
      .setBackground(LIGHT_BLUE).setFontSize(10)
      .setWrap(true).setVerticalAlignment("top");
    var lines = sec.body.split("\n").length;
    inst.setRowHeight(row, Math.max(60, lines * 18));
    row++;
  });

  // Single column at 680px — fits 8.5x11 portrait with normal margins
  inst.setColumnWidth(1, 680);

  // Hide all other columns
  var maxCols = inst.getMaxColumns();
  if (maxCols > 1) inst.hideColumns(2, maxCols - 1);

  inst.setFrozenRows(1);

  // Set portrait orientation and fit to page width for printing
  var ps = inst.getPageSetup();
  ps.setOrientation(SpreadsheetApp.PageOrientation.PORTRAIT);
  ps.setFitToWidth(1);
  inst.setPageSetup(ps);

  // Lock the tab — read only
  var instProt = inst.protect().setDescription("Instructions - Read Only");
  instProt.removeEditors(instProt.getEditors());

  SpreadsheetApp.getUi().alert("Instructions tab ready!\n\nTo print: File → Print → Portrait → Fit to page width → Print\n\nRun Step 7 — Setup Reports next.");
}

// ============================================================
// END: SETUP STEP 6 — INSTRUCTIONS TAB
// ============================================================


// ============================================================
// START: SETUP STEP 7 — REPORTS TAB
// ============================================================

function runSetupReports() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var reports = ss.getSheetByName("Reports");
  if (!reports) { SpreadsheetApp.getUi().alert("Run Step 1 first."); return; }

  reports.clearContents();
  reports.clearFormats();
  reports.clearNotes();
  reports.getRange(1, 1, reports.getMaxRows(), reports.getMaxColumns()).clearDataValidations();

  // Remove any existing sheet or range protections
  var sheetProts = reports.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var p = 0; p < sheetProts.length; p++) sheetProts[p].remove();
  var rangeProts = reports.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var p = 0; p < rangeProts.length; p++) rangeProts[p].remove();

  var sources = getIncomeSources(ss);
  if (sources.length === 0) sources = ["DoorDash", "Uber", "Freelance", "Notary", "Other"];

  // Read expense categories from Setup (only rows typed "Expense")
  var expenseCats = [];
  var setup = ss.getSheetByName("Setup");
  if (setup) {
    var setupData = setup.getDataRange().getValues();
    var inCats = false;
    for (var i = 0; i < setupData.length; i++) {
      var cName = String(setupData[i][0]).trim();
      var cType = String(setupData[i][1]).trim();
      if (cName === "INCOME & EXPENSE CATEGORIES") { inCats = true; continue; }
      if (inCats) {
        if (cName !== "" && cName === cName.toUpperCase()) break;
        if (cName.indexOf("\n") !== -1) break;
        if (cName !== "" && cType === "Expense") expenseCats.push(cName);
      }
    }
  }
  if (expenseCats.length === 0) expenseCats = ["Mileage Deduction", "Fuel", "Other Expense"];

  var row = 1;

  // -- MAIN HEADER --
  reports.getRange(row, 1, 1, 4).merge()
    .setValue("INCOME & EXPENSE REPORT")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center");
  reports.setRowHeight(row, 50);
  row++;

  // -- FILTER LABELS (row 2) --
  reports.getRange(row, 1, 1, 4)
    .setValues([["Start Date", "End Date", "Income Source", ""]])
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold")
    .setHorizontalAlignment("center");
  row++;

  // -- FILTER INPUTS (row 3) --
  var filterRow = row;
  var currentYear = new Date().getFullYear();
  reports.getRange(row, 1)
    .setValue(new Date(currentYear, 0, 1))
    .setNumberFormat("MM/dd/yyyy")
    .setBackground("white").setBorder(true, true, true, true, false, false)
    .setHorizontalAlignment("center");
  reports.getRange(row, 2)
    .setValue(new Date())
    .setNumberFormat("MM/dd/yyyy")
    .setBackground("white").setBorder(true, true, true, true, false, false)
    .setHorizontalAlignment("center");
  var srcOptions  = ["All Sources"].concat(sources);
  var srcDropRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(srcOptions, true).setAllowInvalid(false).build();
  reports.getRange(row, 3)
    .setValue("All Sources").setDataValidation(srcDropRule)
    .setBackground("white").setBorder(true, true, true, true, false, false)
    .setHorizontalAlignment("center");
  reports.getRange(row, 4).setBackground(LIGHT_GOLD);
  reports.setRowHeight(row, 36);
  row++;

  // -- FILTER TIP (row 4) --
  reports.getRange(row, 1, 1, 4).merge()
    .setValue("Change dates or income source above — all tables update automatically.")
    .setBackground(LIGHT_GOLD).setFontColor("#666666").setFontSize(9).setFontStyle("italic")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  reports.setRowHeight(row, 20);
  row++;

  row++; // spacer row 5

  // ============================================================
  // SECTION 1 — INCOME & EXPENSES BY SOURCE
  // ============================================================
  reports.getRange(row, 1, 1, 4).merge()
    .setValue("INCOME & EXPENSES BY SOURCE")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  reports.setRowHeight(row, 32);
  row++;

  reports.getRange(row, 1, 1, 4)
    .setValues([["Source", "Income", "Expenses", "Net"]])
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold")
    .setHorizontalAlignment("center");
  row++;

  var srcDataStart = row;
  var fr = filterRow; // shorthand for embedding in formula strings

  sources.forEach(function(src, idx) {
    var bg = idx % 2 === 0 ? LIGHT_BLUE : "white";
    reports.getRange(row, 1).setValue(src).setBackground(bg);

    // Income: all-sources path shows every row; specific-source path blanks non-matching rows
    var incF =
      '=IF($C$' + fr + '="All Sources",' +
        'IFERROR(SUMPRODUCT((Transactions!$A$2:$A$1000>=$A$' + fr + ')*(Transactions!$A$2:$A$1000<=$B$' + fr + ')*(Transactions!$B$2:$B$1000="' + src + '")*(Transactions!$C$2:$C$1000="Income")*Transactions!$D$2:$D$1000),0),' +
        'IF(A' + row + '=$C$' + fr + ',' +
          'IFERROR(SUMPRODUCT((Transactions!$A$2:$A$1000>=$A$' + fr + ')*(Transactions!$A$2:$A$1000<=$B$' + fr + ')*(Transactions!$B$2:$B$1000="' + src + '")*(Transactions!$C$2:$C$1000="Income")*Transactions!$D$2:$D$1000),0)' +
        ',""))';
    reports.getRange(row, 2).setFormula(incF).setBackground(bg).setNumberFormat("$#,##0.00");

    var expF =
      '=IF($C$' + fr + '="All Sources",' +
        'IFERROR(SUMPRODUCT((Transactions!$A$2:$A$1000>=$A$' + fr + ')*(Transactions!$A$2:$A$1000<=$B$' + fr + ')*(Transactions!$B$2:$B$1000="' + src + '")*(Transactions!$C$2:$C$1000="Expense")*ABS(Transactions!$D$2:$D$1000)),0),' +
        'IF(A' + row + '=$C$' + fr + ',' +
          'IFERROR(SUMPRODUCT((Transactions!$A$2:$A$1000>=$A$' + fr + ')*(Transactions!$A$2:$A$1000<=$B$' + fr + ')*(Transactions!$B$2:$B$1000="' + src + '")*(Transactions!$C$2:$C$1000="Expense")*ABS(Transactions!$D$2:$D$1000)),0)' +
        ',""))';
    reports.getRange(row, 3).setFormula(expF).setBackground(bg).setNumberFormat("$#,##0.00");

    // Net: blank if both Income and Expense cells are blank (specific-source, non-matching row)
    reports.getRange(row, 4)
      .setFormula('=IF(AND(B' + row + '="",C' + row + '=""),"",IFERROR(B' + row + ',0)-IFERROR(C' + row + ',0))')
      .setBackground(bg).setNumberFormat("$#,##0.00");

    row++;
  });

  // TOTAL row (section 1)
  var srcTotalRow = row;
  reports.getRange(row, 1).setValue("TOTAL")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold");
  reports.getRange(row, 2)
    .setFormula('=IFERROR(SUM(B' + srcDataStart + ':B' + (row - 1) + '),0)')
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  reports.getRange(row, 3)
    .setFormula('=IFERROR(SUM(C' + srcDataStart + ':C' + (row - 1) + '),0)')
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  reports.getRange(row, 4)
    .setFormula('=B' + row + '-C' + row)
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  row++;

  row++; // spacer

  // ============================================================
  // SECTION 2 — EXPENSES BY CATEGORY
  // ============================================================
  reports.getRange(row, 1, 1, 4).merge()
    .setValue("EXPENSES BY CATEGORY")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  reports.setRowHeight(row, 32);
  row++;

  reports.getRange(row, 1, 1, 4)
    .setValues([["Category", "Amount", "", ""]])
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold")
    .setHorizontalAlignment("center");
  row++;

  var catDataStart = row;
  expenseCats.forEach(function(cat, idx) {
    var bg = idx % 2 === 0 ? LIGHT_BLUE : "white";
    reports.getRange(row, 1).setValue(cat).setBackground(bg);
    reports.getRange(row, 3, 1, 2).setBackground(bg);

    var catF =
      '=IFERROR(SUMPRODUCT(' +
        '(Transactions!$A$2:$A$1000>=$A$' + fr + ')*' +
        '(Transactions!$A$2:$A$1000<=$B$' + fr + ')*' +
        'IF($C$' + fr + '="All Sources",1,(Transactions!$B$2:$B$1000=$C$' + fr + '))*' +
        '(Transactions!$C$2:$C$1000="Expense")*' +
        '(Transactions!$E$2:$E$1000="' + cat + '")*' +
        'ABS(Transactions!$D$2:$D$1000)' +
      '),0)';
    reports.getRange(row, 2).setFormula(catF).setBackground(bg).setNumberFormat("$#,##0.00");
    row++;
  });

  // TOTAL row (section 2)
  reports.getRange(row, 1).setValue("TOTAL EXPENSES")
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold");
  reports.getRange(row, 2)
    .setFormula('=IFERROR(SUM(B' + catDataStart + ':B' + (row - 1) + '),0)')
    .setBackground(GOLD).setFontColor(DARK_BLUE).setFontWeight("bold").setNumberFormat("$#,##0.00");
  reports.getRange(row, 3, 1, 2).setBackground(GOLD);
  row++;

  row++; // spacer

  // ============================================================
  // SECTION 3 — OVERALL SUMMARY
  // ============================================================
  reports.getRange(row, 1, 1, 4).merge()
    .setValue("OVERALL SUMMARY")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  reports.setRowHeight(row, 32);
  row++;

  var totalIncRow = row;
  reports.getRange(row, 1).setValue("Total Income").setFontWeight("bold").setBackground(LIGHT_GOLD);
  reports.getRange(row, 2)
    .setFormula('=IFERROR(SUM(B' + srcDataStart + ':B' + (srcTotalRow - 1) + '),0)')
    .setBackground(LIGHT_GOLD).setNumberFormat("$#,##0.00");
  reports.getRange(row, 3, 1, 2).setBackground(LIGHT_GOLD);
  row++;

  var totalExpRow = row;
  reports.getRange(row, 1).setValue("Total Expenses").setFontWeight("bold").setBackground(LIGHT_BLUE);
  reports.getRange(row, 2)
    .setFormula('=IFERROR(SUM(C' + srcDataStart + ':C' + (srcTotalRow - 1) + '),0)')
    .setBackground(LIGHT_BLUE).setNumberFormat("$#,##0.00");
  reports.getRange(row, 3, 1, 2).setBackground(LIGHT_BLUE);
  row++;

  var netRow = row;
  reports.getRange(row, 1).setValue("Net Profit").setFontWeight("bold").setBackground(LIGHT_GOLD);
  reports.getRange(row, 2)
    .setFormula('=B' + totalIncRow + '-B' + totalExpRow)
    .setBackground(LIGHT_GOLD).setNumberFormat("$#,##0.00");
  reports.getRange(row, 3, 1, 2).setBackground(LIGHT_GOLD);
  row++;

  reports.getRange(row, 1).setValue("Est. Tax Set-Aside (28%)").setFontWeight("bold").setBackground(LIGHT_BLUE);
  reports.getRange(row, 2)
    .setFormula('=MAX(0,B' + netRow + '*0.28)')
    .setBackground(LIGHT_BLUE).setNumberFormat("$#,##0.00");
  reports.getRange(row, 3, 1, 2).setBackground(LIGHT_BLUE);
  row++;

  // -- COLUMN WIDTHS --
  reports.setColumnWidth(1, 200);
  reports.setColumnWidth(2, 140);
  reports.setColumnWidth(3, 140);
  reports.setColumnWidth(4, 140);

  // Hide columns beyond D
  var maxCols = reports.getMaxColumns();
  if (maxCols > 4) reports.hideColumns(5, maxCols - 4);

  reports.setFrozenRows(1);

  // Lock the whole sheet, but leave the three filter input cells editable
  var prot = reports.protect().setDescription("Reports - layout locked");
  prot.setUnprotectedRanges([reports.getRange(filterRow, 1, 1, 3)]);
  prot.removeEditors(prot.getEditors());

  SpreadsheetApp.getUi().alert("Reports tab ready!\n\nUse the date and source filters in row 3 to slice the report.");
}

// ============================================================
// END: SETUP STEP 7 — REPORTS TAB
// ============================================================


// ============================================================
// START: SETUP STEP 8 — MOBILE ENTRY TAB
// ============================================================

function runSetupMobileEntry() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var entry = ss.getSheetByName("Mobile Entry");
  if (!entry) { SpreadsheetApp.getUi().alert("Run Step 1 first."); return; }

  entry.clearContents();
  entry.clearFormats();
  entry.clearNotes();
  entry.getRange(1, 1, entry.getMaxRows(), entry.getMaxColumns()).clearDataValidations();

  var existingProts = entry.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var p = 0; p < existingProts.length; p++) existingProts[p].remove();

  var row = 1;

  // -- HEADER --
  entry.getRange(row, 1, 1, 3).merge()
    .setValue("MOBILE ENTRY FORM")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(28)
    .setHorizontalAlignment("center");
  entry.setRowHeight(row, 70);
  row++;

  // -- SUBTITLE --
  entry.getRange(row, 1, 1, 3).merge()
    .setValue("Fill each field and press Enter to move to the next.  Fields marked * are required.")
    .setBackground(LIGHT_GOLD).setFontColor("#666666").setFontSize(12).setFontStyle("italic")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  entry.setRowHeight(row, 30);
  row++;

  // -- INPUT FIELDS --
  // row 3 = Date, 4 = Source, 5 = Type, 6 = Amount, 7 = Category, 8 = Miles, 9 = Hours, 10 = Notes
  var fields = [
    { label: "Date",          hint: "Leave blank to default to today",            required: false },
    { label: "Income Source", hint: "Required — select your platform or source",  required: true  },
    { label: "Type",          hint: "Required — Income or Expense",               required: true  },
    { label: "Amount",        hint: "Dollar amount (required unless mileage only)",required: false },
    { label: "Category",      hint: "Select the income or expense category",      required: false },
    { label: "Miles Driven",  hint: "Optional — IRS deduction auto-calculates",   required: false },
    { label: "Hours Worked",  hint: "Optional — for tracking time",               required: false },
    { label: "Notes",         hint: "Optional — any extra detail",                required: false }
  ];

  fields.forEach(function(field, idx) {
    var bg = idx % 2 === 0 ? LIGHT_BLUE : "white";
    entry.setRowHeight(row, 60);

    entry.getRange(row, 1)
      .setValue(field.required ? field.label + "  *" : field.label)
      .setBackground(bg).setFontWeight("bold").setFontSize(24)
      .setHorizontalAlignment("right").setVerticalAlignment("middle");

    entry.getRange(row, 2)
      .setBackground("white")
      .setBorder(true, true, true, true, false, false)
      .setVerticalAlignment("middle").setFontSize(24);

    entry.getRange(row, 3)
      .setValue(field.hint)
      .setBackground(bg).setFontColor("#888888").setFontSize(11).setFontStyle("italic")
      .setVerticalAlignment("middle").setWrap(true);

    row++;
  });

  // Number formats for input column
  entry.getRange("B3").setValue(new Date()).setNumberFormat("MM/dd/yyyy");
  entry.getRange("B6").setNumberFormat("$#,##0.00");
  entry.getRange("B8").setNumberFormat("0.0");
  entry.getRange("B9").setNumberFormat("0.00");
  entry.getRange("B10").setWrap(true);

  // Dropdowns
  var sources = getIncomeSources(ss);
  var cats    = getCategories(ss);
  var types   = ["Income", "Expense"];
  if (sources.length === 0) sources = ["DoorDash", "Uber", "Freelance", "Notary", "Other"];
  if (cats.length === 0)    cats    = ["Delivery Income", "Fuel", "Other Expense"];

  entry.getRange("B4")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(sources, true).setAllowInvalid(false).build());
  entry.getRange("B5")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(types, true).setAllowInvalid(false).build());
  entry.getRange("B7")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(cats, true).setAllowInvalid(false).build());

  // -- SUBMIT BUTTON (rows 11 = spacer, 12 = top border, 13 = button, 14 = bottom border) --
  row++; // row 11 spacer
  entry.setRowHeight(row, 12);
  row++;

  // Row 12: top border
  entry.getRange(row, 1, 1, 3).merge().setBackground(DARK_BLUE);
  entry.setRowHeight(row, 6);
  row++;

  // Row 13: submit checkbox + label
  entry.getRange(row, 1).insertCheckboxes().setValue(false).setBackground(DARK_BLUE);
  entry.getRange(row, 2, 1, 2).merge()
    .setValue("CHECK THE BOX TO SUBMIT  ✔  (Allow 2-3 seconds)")
    .setBackground(DARK_BLUE).setFontColor(GOLD)
    .setFontWeight("bold").setFontSize(20)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  entry.setRowHeight(row, 70);
  row++;

  // Row 14: bottom border
  entry.getRange(row, 1, 1, 3).merge().setBackground(DARK_BLUE);
  entry.setRowHeight(row, 6);

  // -- COLUMN WIDTHS --
  entry.setColumnWidth(1, 220);
  entry.setColumnWidth(2, 280);
  entry.setColumnWidth(3, 300);

  // Hide columns beyond C
  var maxCols = entry.getMaxColumns();
  if (maxCols > 3) entry.hideColumns(4, maxCols - 3);

  entry.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("Quick Entry tab ready!\n\nPress Enter after each field to move down. Check the box at the bottom to submit.");
}

// ============================================================
// END: SETUP STEP 8 — MOBILE ENTRY TAB
// ============================================================
