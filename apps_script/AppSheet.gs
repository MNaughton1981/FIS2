/**
 * AppSheet.gs — builds normalized "feeder" tables for an AppSheet front-end.
 *
 * WHY THIS EXISTS
 *   AppSheet reads row-per-record tables, but this workbook is a wide month-matrix
 *   (categories × 12 months). This script generates clean App_* tabs — one row per
 *   record, with live formulas pulling from your existing tabs — so AppSheet always
 *   sees current numbers. Re-run after a layout change or when you want a refresh.
 *
 * TABLES BUILT (bind these in AppSheet; suggested KEY column in brackets)
 *   App_CashFlow  [Month]            Month, Income, Expenses, Net, SavingsRate
 *   App_Spending  [ID]              ID, Month, Category, Amount   (unpivot of CC Detail)
 *   App_Accounts  [Account]          Account, Balance              (latest month-end)
 *   App_Budget    [Lane]             Lane, Budget, Spent, Rollover, Status (guardrail)
 *   App_KPIs      [Key]              Key, Label, Value, Unit       (dashboard cards)
 *
 * MENU
 *   This file defines NO onOpen() (Categorize.gs owns the single menu). To add a
 *   button, paste this line into the menu in Categorize.gs's onOpen():
 *       .addItem("Build / Refresh AppSheet Tables", "buildAppSheetTables")
 *   Or just run buildAppSheetTables once from the Apps Script editor.
 */

const APP_MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function buildAppSheetTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const built = [];
  const tryBuild = (fn, label) => { try { fn(ss); built.push(label); } catch (e) { Logger.log(label + " failed: " + e.message); } };

  tryBuild(appCashFlow_, "App_CashFlow");
  tryBuild(appSpending_, "App_Spending");
  tryBuild(appAccounts_, "App_Accounts");
  tryBuild(appBudget_,   "App_Budget");
  tryBuild(appKpis_,     "App_KPIs");

  SpreadsheetApp.getUi().alert(
    "AppSheet tables refreshed:\n\n  " + built.join("\n  ") +
    "\n\nBind these in AppSheet (appsheet.com). They hold live formulas, so re-run " +
    "this after big layout changes."
  );
}

/* ------------------------------ helpers -------------------------------- */

function appFresh_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  else sh.clear();
  sh.setTabColor("674EA7");
  return sh;
}

// First row in column A whose label contains substr (case-insensitive). -1 if none.
function appRow_(sheet, substr) {
  if (!sheet) return -1;
  const vals = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  const n = String(substr).toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase().indexOf(n) !== -1) return i + 1;
  }
  return -1;
}

function appColL_(c) {
  let s = "";
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - m - 1) / 26); }
  return s;
}

function lastNonBlank_(tab, row) {
  // formula: last non-blank value in B:M of `row` on `tab`
  return "=IFERROR(LOOKUP(2,1/('" + tab + "'!B" + row + ":M" + row + "<>\"\"),'" + tab + "'!B" + row + ":M" + row + "),0)";
}

/* ------------------------------ builders ------------------------------- */

function appCashFlow_(ss) {
  const is = ss.getSheetByName("Income Statement");
  if (!is) throw new Error("Income Statement not found");
  const incR = appRow_(is, "total income");
  const expR = appRow_(is, "total expenses");
  const netR = appRow_(is, "net cash flow");
  const sh = appFresh_(ss, "App_CashFlow");
  const rows = [["Month", "Income", "Expenses", "Net", "SavingsRate"]];
  for (let m = 0; m < 12; m++) {
    const L = appColL_(2 + m);
    rows.push([
      APP_MONTHS[m],
      "='Income Statement'!" + L + incR,
      "='Income Statement'!" + L + expR,
      "='Income Statement'!" + L + netR,
      "=IFERROR('Income Statement'!" + L + netR + "/'Income Statement'!" + L + incR + ',"")'
    ]);
  }
  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  sh.getRange(1, 1, 1, 5).setFontWeight("bold");
}

function appSpending_(ss) {
  const cc = ss.getSheetByName("CC Detail");
  if (!cc) throw new Error("CC Detail not found");
  const sh = appFresh_(ss, "App_Spending");
  const rows = [["ID", "Month", "Category", "Amount"]];
  for (let r = 5; r <= 17; r++) {              // CC_CATS live in rows 5..17
    const cat = cc.getRange(r, 1).getValue();
    if (!cat) continue;
    for (let m = 0; m < 12; m++) {
      const L = appColL_(2 + m);
      rows.push([APP_MONTHS[m] + " | " + cat, APP_MONTHS[m], cat, "=N('CC Detail'!" + L + r + ")"]);
    }
  }
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.getRange(1, 1, 1, 4).setFontWeight("bold");
}

function appAccounts_(ss) {
  const bal = ss.getSheetByName("Balances");
  if (!bal) throw new Error("Balances not found");
  const sh = appFresh_(ss, "App_Accounts");
  const rows = [["Account", "Balance"]];
  for (let r = 5; r <= 16; r++) {              // asset rows (above Total assets)
    const name = bal.getRange(r, 1).getValue();
    if (!name) continue;
    rows.push([name, lastNonBlank_("Balances", r)]);
  }
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight("bold");
}

function appBudget_(ss) {
  const g = ss.getSheetByName("Spending Guardrail");
  if (!g) throw new Error("Spending Guardrail not found");
  const sh = appFresh_(ss, "App_Budget");
  const rows = [["Lane", "Budget", "Spent", "Rollover", "Status"]];
  for (let r = 14; r <= 17; r++) {             // standings panel: Matt/Sue/Household/Combined
    const lane = g.getRange(r, 1).getValue();
    if (!lane) continue;
    rows.push([
      lane,
      "='Spending Guardrail'!B" + r,
      "='Spending Guardrail'!C" + r,
      "='Spending Guardrail'!D" + r,
      "='Spending Guardrail'!E" + r
    ]);
  }
  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  sh.getRange(1, 1, 1, 5).setFontWeight("bold");
}

function appKpis_(ss) {
  const fd = ss.getSheetByName("FIRE Dashboard");
  const bal = ss.getSheetByName("Balances");
  const pen = ss.getSheetByName("Pension - Sue");
  const g = ss.getSheetByName("Spending Guardrail");
  const sh = appFresh_(ss, "App_KPIs");

  const fdv = (sub) => { const r = appRow_(fd, sub); return r > 0 ? "='FIRE Dashboard'!B" + r : ""; };
  const assetR = appRow_(bal, "total assets");
  const penR = appRow_(pen, "annual pension");

  const rows = [["Key", "Label", "Value", "Unit"]];
  rows.push(["net_worth",   "Net worth",                 assetR > 0 ? lastNonBlank_("Balances", assetR) : "", "$"]);
  rows.push(["invested",    "Total invested",            fdv("total invested"), "$"]);
  rows.push(["savings_rate","Savings rate",              fdv("savings rate"),   "%"]);
  rows.push(["fi_number",   "FI number",                 fdv("gross fi number"),"$"]);
  rows.push(["years_to_fi", "Years to FI",               fdv("years to reach gross"), "yrs"]);
  rows.push(["pension",     "Sue pension (annual)",      penR > 0 ? "='Pension - Sue'!B" + penR : "", "$"]);
  rows.push(["guardrail",   "Guardrail rollover (combined)", g ? "='Spending Guardrail'!D17" : "", "$"]);
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.getRange(1, 1, 1, 4).setFontWeight("bold");
}
