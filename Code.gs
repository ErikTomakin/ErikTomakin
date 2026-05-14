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
    .addSeparator()
    .addItem("Refresh Dropdowns",       "refreshDropdowns")
    .addItem("Clear All Transactions",  "clearAllTransactions")
    .addToUi();

  checkIRSRate();
}

function onEdit(e) {
  var sheet = e.range.getSheet();
  var row   = e.range.getRow();
  var col   = e.range.getColumn();

  // Transactions tab: column H checkbox toggles row protection on/off
  if (sheet.getName() === "Transactions" && col === 8 && row > 1) {
    if (e.value === true) {
      // Checked — remove protection so user can edit or delete this row
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var i = 0; i < protections.length; i++) {
        if (protections[i].getDescription() === "tx_row_" + row) {
          protections[i].remove();
          break;
        }
      }
    } else {
      // Unchecked — re-lock this row
      var protection = sheet.getRange(row, 1, 1, 7).protect().setDescription("tx_row_" + row);
      protection.removeEditors(protection.getEditors());
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
      if (cell === "" || cell.toUpperCase() === cell && cell !== "") break;
      if (cell) sources.push(cell);
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
    if (cell === "EXPENSE CATEGORIES") { inSection = true; continue; }
    if (inSection) {
      if (cell === "" || (cell === cell.toUpperCase() && cell !== "" && cats.length > 0)) break;
      if (cell) cats.push(cell);
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
  var types   = ["Income", "Expense"];

  if (sources.length === 0) sources = ["DoorDash", "Uber", "Freelance", "Notary", "Other"];
  if (cats.length === 0)    cats    = ["Delivery Income", "Fuel", "Other Expense"];

  var lastRow  = 1000;
  var srcRule  = SpreadsheetApp.newDataValidation().requireValueInList(sources, true).setAllowInvalid(false).build();
  var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(types, true).setAllowInvalid(false).build();
  var catRule  = SpreadsheetApp.newDataValidation().requireValueInList(cats, true).setAllowInvalid(false).build();

  // Transactions tab — columns B (source), C (type), E (category)
  txSheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(srcRule);
  txSheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(typeRule);
  txSheet.getRange(2, 5, lastRow - 1, 1).setDataValidation(catRule);

  // Dashboard entry form
  dash.getRange("B3").setDataValidation(srcRule);
  dash.getRange("C3").setDataValidation(typeRule);
  dash.getRange("A5").setDataValidation(catRule);

  SpreadsheetApp.getUi().alert("Dropdowns refreshed!");
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
  // Drawing button resets itself — nothing to clear here
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

// ============================================================
// END: TRANSACTION LOGIC — addTransaction / getMileageRate / clearForm / clearAllTransactions
// ============================================================


// ============================================================
// START: SETUP STEP 1 — CREATE TABS
// ============================================================

function runCreateSheets() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var tabNames = ["Dashboard", "Transactions", "Setup", "AI Prompts", "Instructions"];
  var colors   = [GOLD, DARK_BLUE, GOLD, DARK_BLUE, GOLD];

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

  // Format date and currency
  tx.getRange(2, 1, 999, 1).setNumberFormat("MM/dd/yyyy");
  tx.getRange(2, 4, 999, 1).setNumberFormat("$#,##0.00");

  // Wrap text in Notes column (G)
  tx.getRange(2, 7, 999, 1).setWrap(true);

  tx.setFrozenRows(1);
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

  var row = 1;

  // -- MAIN HEADER --
  setup.getRange(row, 1, 1, 4).merge()
    .setValue("SETUP & CONFIGURATION")
    .setBackground(DARK_BLUE).setFontColor("white")
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center");
  setup.setRowHeight(row, 40);
  row++;

  // -- INCOME SOURCES SECTION --
  row++;
  setup.getRange(row, 1, 1, 4).merge()
    .setValue("INCOME SOURCES")
    .setBackground(GOLD).setFontColor(DARK_BLUE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  setup.setRowHeight(row, 32);
  row++;

  setup.getRange(row, 1, 1, 4).setValues([["Source Name", "Description", "", ""]])
    .setBackground(DARK_BLUE).setFontColor("white").setFontWeight("bold");
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

  // -- EXPENSE CATEGORIES SECTION --
  row++;
  setup.getRange(row, 1, 1, 4).merge()
    .setValue("EXPENSE CATEGORIES")
    .setBackground(GOLD).setFontColor(DARK_BLUE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center");
  setup.setRowHeight(row, 32);
  row++;

  setup.getRange(row, 1, 1, 4).setValues([["Category Name", "Type", "", ""]])
    .setBackground(DARK_BLUE).setFontColor("white").setFontWeight("bold");
  row++;

  var categories = [
    ["Delivery Income",        "Income"],
    ["Ride Income",            "Income"],
    ["Freelance Income",       "Income"],
    ["Notary Income",          "Income"],
    ["Product Sales",          "Income"],
    ["Affiliate Income",       "Income"],
    ["Subscription Income",    "Income"],
    ["Other Income",           "Income"],
    ["Mileage Deduction",      "Expense"],
    ["Fuel",                   "Expense"],
    ["Vehicle Maintenance",    "Expense"],
    ["Car Insurance",          "Expense"],
    ["Parking & Tolls",        "Expense"],
    ["Shipping",               "Expense"],
    ["Packaging",              "Expense"],
    ["Platform Fees",          "Expense"],
    ["Product Supplies",       "Expense"],
    ["Software Subscriptions", "Expense"],
    ["Printing & Supplies",    "Expense"],
    ["Professional Fees",      "Expense"],
    ["Phone Bill",             "Expense"],
    ["Internet",               "Expense"],
    ["Home Office",            "Expense"],
    ["Advertising",            "Expense"],
    ["Recurring Expense",      "Expense"],
    ["Other Expense",          "Expense"]
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
  row++;

  setup.getRange(row, 1, 1, 4).setValues([["Expense", "Amount", "Frequency", "Next Due Date"]])
    .setBackground(DARK_BLUE).setFontColor("white").setFontWeight("bold");
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
  row++;

  setup.getRange(row, 1, 1, 4).setValues([["Description", "Rate (per mile)", "Effective Date", "Last Updated"]])
    .setBackground(DARK_BLUE).setFontColor("white").setFontWeight("bold");
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

  refreshDropdowns();
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

  // Rows 7-9: Space for the SUBMIT Drawing button (added manually after setup)
  dash.getRange("A7:D9").clearContent().clearFormat();
  dash.getRange("A7:D7").merge()
    .setValue("↓  Insert your SUBMIT button here (see Instructions tab)")
    .setBackground(LIGHT_GOLD).setFontColor("#888888").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dash.setRowHeight(7, 16);
  dash.setRowHeight(8, 50);
  dash.setRowHeight(9, 10);

  // Row 10: Spacer
  dash.setRowHeight(10, 20);

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

  refreshDropdowns();
  SpreadsheetApp.getUi().alert("Dashboard rebuilt!\n\nNext: manually add the SUBMIT button (see Instructions tab), then run Step 5.");
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
            "4. Run 'Refresh Dropdowns' from the Expense Tracker menu after any changes to Setup."
    },
    {
      heading: "ADDING THE SUBMIT BUTTON (one-time setup)",
      body: "1. On the Dashboard tab, click Insert → Drawing.\n" +
            "2. Click the Shapes tool → Shapes → Rounded Rectangle.\n" +
            "3. Draw a wide pill shape. Fill color: silver (#D4D4D4). Border: none or dark blue.\n" +
            "4. Double-click the shape and type: SUBMIT\n" +
            "5. Make the text bold, dark blue (#0055A2), font size 16.\n" +
            "6. Click Save & Close.\n" +
            "7. Right-click the floating button → Assign Script → type: addTransaction → OK.\n" +
            "8. Drag the button to sit over rows 8–9 on the Dashboard."
    },
    {
      heading: "ADDING TRANSACTIONS",
      body: "1. Go to the Dashboard tab.\n" +
            "2. Fill in: Date, Income Source, Type (Income or Expense), and Amount.\n" +
            "3. Fill in Category, Miles Driven (optional), Hours Worked (optional), and Notes.\n" +
            "4. Click the silver SUBMIT button.\n" +
            "5. If you entered mileage, a second row will auto-post your mileage deduction.\n" +
            "6. You can also use Expense Tracker menu → Add Transaction."
    },
    {
      heading: "INCOME SOURCES vs EXPENSE CATEGORIES",
      body: "• Income Sources = WHERE the money came from (DoorDash, Etsy, Upwork, etc.)\n" +
            "• Expense Categories = WHAT the money was spent on (Fuel, Phone Bill, etc.)\n" +
            "• Type 'Income' = money coming IN | Type 'Expense' = money going OUT\n" +
            "• Mileage Deduction posts as a NEGATIVE expense automatically."
    },
    {
      heading: "DASHBOARD SUMMARIES",
      body: "• BY SOURCE: Shows total income, expenses, and net per platform for the year.\n" +
            "• MONTHLY BREAKDOWN: Shows all 12 months of the current year.\n" +
            "• These update automatically as you add transactions."
    },
    {
      heading: "USING AI (NO API NEEDED)",
      body: "1. Go to the Transactions tab and copy your data (Ctrl+A, Ctrl+C).\n" +
            "2. Go to the AI Prompts tab and copy one of the 3 prompts.\n" +
            "3. Open ChatGPT, Claude, Gemini, or any AI tool.\n" +
            "4. Paste the prompt first, then paste your data below it.\n" +
            "5. The AI will analyze your finances — no subscription or API key required."
    },
    {
      heading: "MENU SHORTCUTS",
      body: "Expense Tracker menu (top of screen):\n" +
            "• Add Transaction — submits your form entry\n" +
            "• Refresh Dropdowns — updates dropdowns after you edit Setup\n" +
            "• Clear All Transactions — wipes all rows (requires YES confirmation)\n" +
            "• Setup steps 1–6 — rebuilds individual tabs if needed"
    },
    {
      heading: "TIPS FOR GIG WORKERS",
      body: "• Log mileage the same day — it's easy to forget.\n" +
            "• Set aside 25-30% of net profit for quarterly taxes.\n" +
            "• Check the AI Prompts tab before every quarterly tax deadline.\n" +
            "• Keep receipts for all Expense entries — photos in Google Drive work great.\n" +
            "• Use 'Notary Work' or 'Freelance Income' for non-platform income."
    }
  ];

  sections.forEach(function(sec) {
    row++;
    inst.getRange(row, 1, 1, 2).merge()
      .setValue(sec.heading)
      .setBackground(GOLD).setFontColor(DARK_BLUE)
      .setFontWeight("bold").setFontSize(12);
    inst.setRowHeight(row, 32);
    row++;

    inst.getRange(row, 1, 1, 2).merge()
      .setValue(sec.body)
      .setBackground(LIGHT_BLUE).setFontSize(11)
      .setWrap(true).setVerticalAlignment("top");
    var lines = sec.body.split("\n").length;
    inst.setRowHeight(row, Math.max(80, lines * 22));
    row++;
  });

  inst.setColumnWidth(1, 700);
  inst.setColumnWidth(2, 700);

  SpreadsheetApp.getUi().alert("Setup complete! Your expense tracker is ready.");
}

// ============================================================
// END: SETUP STEP 6 — INSTRUCTIONS TAB
// ============================================================
