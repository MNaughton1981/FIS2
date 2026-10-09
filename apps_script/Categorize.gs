/**
 * Categorize.gs — transaction categorizer & roll-up for the Family Income Statement.
 *
 * WHAT CHANGED vs the old v4 script (and why)
 *   • Rows are resolved BY LABEL, not hard-coded numbers. The v4 script hard-coded
 *     e.g. "Schwab investment": 28, "Fidelity investment": 29, "Federal taxes": 56,
 *     "Electric": 34. When the workbook moved to v5 a fixed row was removed (the old
 *     529/Schwab split merged), shifting everything below up by one — so those numbers
 *     now point at the WRONG rows (Fidelity would land on the Total row, etc.). Looking
 *     rows up by their label makes the script immune to future layout shifts.
 *   • Utilities (Eversource) are protected: EVERSOURCE is intercepted as __MANUAL__ so
 *     gas + electric are NEVER auto-written — you enter them by hand and they survive
 *     every import.
 *   • Any variable category that isn't a real CC Detail row (e.g. a stray "Entertainment"
 *     rule) is folded into "Other / misc" instead of being silently dropped.
 *
 * MENU NOTE
 *   This file defines the single onOpen() and a merged "📊 Family Finances" menu that
 *   also lists the MonthlyEntry.gs items. DELETE the onOpen() function inside
 *   MonthlyEntry.gs (keep its other functions) so there is only one onOpen in the project.
 */

const CC_CATS = [
  "Groceries", "Gas / Vehicle", "Dining out", "Subscriptions",
  "Kids / Activities", "Travel", "Personal Care", "Pharmacy / Health",
  "Home / Hardware", "Amazon (misc)", "Landscaping",
  "Cash / allowances", "Other / misc"
];

// Token suffix (from Rules) -> a distinctive lowercase substring of the Income
// Statement row label. Rows are found by scanning column A, so exact row numbers
// don't matter and can shift without breaking anything.
const INCOME_LABEL = {
  "Sue":            "sue — net pay",
  "Matt":           "matt — net pay",
  "Interest":       "interest earned",
  "FSA reimb":      "fsa dep care reimbursement",
  "Tuition reimb":  "tuition reimbursement",
  "Tax refunds":    "tax refunds",
  "CC rewards":     "cc rewards",
  "Misc refunds":   "cc rewards",           // same combined row
};

const FIXED_LABEL = {
  "Mortgage":            "mortgage (rocket",
  "Car (Sienna)":        "sienna",
  "Car (RAV4)":          "rav4",
  "Childcare":           "childcare",
  "Tuition":             "tuition (northeastern",
  "Life ins (AmGen)":    "american general",
  "Life ins (ALIC)":     "— alic",
  "Verizon":             "verizon",
  "T-Mobile":            "t-mobile",
  "Schwab investment":   "schwab investment",
  "Fidelity investment": "fidelity investment",
  "Federal taxes":       "federal taxes owed",
  // NOTE: "Electric" intentionally omitted — Eversource is manual (see categorize()).
};

/* --------------------------- IMPORT SOURCES ---------------------------- *
 * One line per account/card. To add a card: create a matching "Imports - <X>"
 * tab (paste the raw CSV; data starts row 5; the LAST column is the auto
 * "→ Category"), add a line below, and re-run "Categorize & Roll Up".
 * `format` points at a FORMATS entry that maps the raw CSV columns.
 * -------------------------------------------------------------------- */
const IMPORT_TABS = [
  { name: "Imports - Checking", format: "capone_bank" },
  { name: "Imports - Savings",  format: "capone_bank" },
  { name: "Imports - Barclay",  format: "barclay" },
  // Phasing out Barclay? Add the replacement card here and make the tab:
  // { name: "Imports - Chase",    format: "chase" },
  // { name: "Imports - Amex",     format: "amex" },
  // { name: "Imports - Discover", format: "discover" },
];

// Column maps are 0-based indices into each pasted row. Everything normalizes to
// a signed `amount`: +in (money in) / -out (spend). Amount types:
//   signed      {col}                        raw value already +in / -out
//   spendPos    {col}                        raw value is +spend (typical card) -> negated
//   debitcredit {amtCol,typeCol,creditWord}  magnitude + a Debit/Credit column
//   splitcols   {debitCol,creditCol}         separate Debit & Credit columns
// VERIFY a new card's column order against its downloaded CSV header and adjust.
const FORMATS = {
  // Capital One BANK: A=acct B=desc C=date D=type(Debit/Credit) E=amount
  capone_bank: { dateCol: 2, descCol: 1, amount: { type: "debitcredit", amtCol: 4, typeCol: 3, creditWord: "credit" } },
  // Barclaycard: A=date B=desc C=category D=amount (charges already negative)
  barclay:     { dateCol: 0, descCol: 1, amount: { type: "signed", col: 3 } },
  // Chase: Trans Date, Post Date, Description, Category, Type, Amount, Memo (Amount: charge -, credit +)
  chase:       { dateCol: 0, descCol: 2, amount: { type: "signed", col: 5 } },
  // Amex: Date, Description, Amount (Amount: charge +)
  amex:        { dateCol: 0, descCol: 1, amount: { type: "spendPos", col: 2 } },
  // Discover: Trans Date, Post Date, Description, Amount, Category (Amount: charge +)
  discover:    { dateCol: 0, descCol: 2, amount: { type: "spendPos", col: 3 } },
  // Citi: Status, Date, Description, Debit, Credit
  citi:        { dateCol: 1, descCol: 2, amount: { type: "splitcols", debitCol: 3, creditCol: 4 } },
  // Capital One CREDIT CARD: Trans Date, Posted Date, Card No, Description, Category, Debit, Credit
  capone_card: { dateCol: 0, descCol: 3, amount: { type: "splitcols", debitCol: 5, creditCol: 6 } },
};

// Pull {date, desc, amount(+in/-out)} from a raw row using a FORMATS entry.
function extractRow_(row, fmt) {
  const A = fmt.amount;
  let amount;
  if (A.type === "spendPos")          amount = -Number(row[A.col]);
  else if (A.type === "debitcredit")  amount = (String(row[A.typeCol]).toLowerCase() === A.creditWord) ? Number(row[A.amtCol]) : -Number(row[A.amtCol]);
  else if (A.type === "splitcols")    amount = (Number(row[A.creditCol]) || 0) - (Number(row[A.debitCol]) || 0);
  else                                amount = Number(row[A.col]);   // "signed"
  if (isNaN(amount)) amount = 0;
  return { date: row[fmt.dateCol], desc: row[fmt.descCol], amount: amount };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Family Finances")
    .addItem("Categorize & Roll Up", "categorizeAll")
    .addItem("Generate / Refresh Charts", "generateCharts")
    .addSeparator()
    .addItem("Post Monthly Entry", "postMonthlyEntry")          // from MonthlyEntry.gs
    .addItem("Create Monthly Form", "createMonthlyForm")        // from MonthlyEntry.gs
    .addItem("Clear Monthly Entry inputs", "clearMonthlyEntry") // from MonthlyEntry.gs
    .addSeparator()
    .addItem("Clear All Imports", "clearAllImports")
    .addItem("Clear CC Detail", "clearCCDetail")
    .addSeparator()
    .addItem("Roll Over to Next Year", "rollOverYear")
    .addToUi();
}

/* ------------------------------ helpers -------------------------------- */

function columnLetter(col) {
  let s = "";
  while (col > 0) { const m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = Math.floor((col - m - 1) / 26); }
  return s;
}

// Find the first row in column A whose label contains `substr` (case-insensitive).
function findRowContains_(sheet, substr) {
  const vals = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  const needle = String(substr).toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase().indexOf(needle) !== -1) return i + 1;
  }
  return -1;
}

function loadRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Rules");
  if (!sheet) throw new Error("Rules tab not found");
  const last = sheet.getLastRow();
  const data = sheet.getRange(5, 1, last - 4, 2).getValues();
  return data
    .filter(r => r[0] && r[1])
    .map(r => ({ pattern: String(r[0]), regex: new RegExp(String(r[0]), "i"), category: String(r[1]).trim() }));
}

function categorize(description, rules) {
  if (!description) return { category: "Other / misc", matched: false };

  // Gas + electric bill together as EVERSOURCE WEB_PAY and can't be split reliably,
  // so they are entered MANUALLY on the Income Statement. Intercept here (before the
  // rules loop) so the importer never buckets or overwrites those rows.
  if (/EVERSOURCE/i.test(description)) return { category: "__MANUAL__", matched: true };

  for (const rule of rules) {
    if (rule.regex.test(description)) return { category: rule.category, matched: true };
  }
  return { category: "Other / misc", matched: false };
}

function getMonth(dateValue) {
  if (dateValue instanceof Date) return dateValue.getMonth();
  const d = new Date(dateValue);
  if (isNaN(d.getTime())) return null;
  return d.getMonth();
}

/* ------------------------------ main ----------------------------------- */

function categorizeAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let rules;
  try { rules = loadRules(); } catch (e) { ui.alert("Error loading rules: " + e.message); return; }

  const monthlyVar = {}, monthlyInc = {}, monthlyFix = {};
  function bucket(obj, month, key, val) { if (!obj[month]) obj[month] = {}; obj[month][key] = (obj[month][key] || 0) + val; }

  function processBatched(sheetName, format) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 0;
    const fmt = FORMATS[format];
    if (!fmt) { Logger.log("Unknown format '" + format + "' for " + sheetName); return 0; }
    const last = sheet.getLastRow();
    if (last < 5) return 0;

    const numRows = last - 4;
    const numCols = sheet.getLastColumn();
    const data = sheet.getRange(5, 1, numRows, numCols).getValues();
    const categories = [];
    const backgrounds = [];
    let processedCount = 0;

    data.forEach((row) => {
      const blank = () => { categories.push([""]); backgrounds.push(new Array(numCols).fill("#ffffff")); };

      const ex = extractRow_(row, fmt);
      const date = ex.date, desc = ex.desc, amount = ex.amount;
      if (!date || !desc) { blank(); return; }

      const month = getMonth(date);
      if (month === null) { blank(); return; }

      const result = categorize(desc, rules);
      const category = result.category;

      if (category === "__TRANSFER__" || category === "__MANUAL__") {
        // ignore — transfers, and manually-entered categories (utilities)
      } else if (category.indexOf("__INCOME__:") === 0) {
        bucket(monthlyInc, month, category.replace("__INCOME__:", ""), amount);
      } else if (category.indexOf("__FIXED__:") === 0) {
        bucket(monthlyFix, month, category.replace("__FIXED__:", ""), Math.abs(amount));
      } else {
        // Variable spending. Fold any category that isn't a real CC Detail row
        // (e.g. a stray "Entertainment" rule) into "Other / misc" so it isn't lost.
        const vc = (CC_CATS.indexOf(category) >= 0) ? category : "Other / misc";
        bucket(monthlyVar, month, vc, -amount);
      }

      categories.push([category]);
      backgrounds.push(new Array(numCols).fill(result.matched ? "#ffffff" : "#FFF2CC"));
      processedCount++;
    });

    sheet.getRange(5, numCols, numRows, 1).setValues(categories);
    sheet.getRange(5, 1, numRows, numCols).setBackgrounds(backgrounds);
    return processedCount;
  }

  const counts = IMPORT_TABS.map(t => ({ name: t.name, n: processBatched(t.name, t.format) }));

  // CC Detail grid (variable spending) — rows 5.. in CC_CATS order.
  const ccSheet = ss.getSheetByName("CC Detail");
  const ccGrid = CC_CATS.map((cat) => {
    const row = new Array(12).fill("");
    for (let m = 0; m < 12; m++) { const v = (monthlyVar[m] && monthlyVar[m][cat]) || 0; if (v !== 0) row[m] = v; }
    return row;
  });
  ccSheet.getRange(5, 2, CC_CATS.length, 12).setValues(ccGrid);

  // Income & fixed rows on the Income Statement — resolved BY LABEL, so a shifted
  // layout can't misplace them. Totals/utilities rows are never touched.
  const is = ss.getSheetByName("Income Statement");
  writeCategoryRows_(is, monthlyInc, INCOME_LABEL);
  writeCategoryRows_(is, monthlyFix, FIXED_LABEL);

  const summary = counts.map(c => "  " + c.name.replace("Imports - ", "") + ": " + c.n + " rows").join("\n");
  ui.alert(
    "Categorize complete.\n\n" + summary + "\n\n" +
    "Utilities (Eversource) were left for manual entry. Yellow rows in Imports = uncategorized."
  );
}

// For each sub-category present in bucketObj, find its Income Statement row by label
// and overlay the monthly values (preserving months you have no data for).
function writeCategoryRows_(is, bucketObj, labelMap) {
  const subs = {};
  Object.keys(bucketObj).forEach(m => Object.keys(bucketObj[m]).forEach(s => { subs[s] = true; }));
  const missing = [];
  Object.keys(subs).forEach(sub => {
    const key = labelMap[sub];
    const row = key ? findRowContains_(is, key) : -1;
    if (row < 0) { missing.push(sub); return; }
    const existing = is.getRange(row, 2, 1, 12).getValues()[0];
    for (let m = 0; m < 12; m++) {
      if (bucketObj[m] && bucketObj[m][sub] !== undefined) existing[m] = bucketObj[m][sub];
    }
    is.getRange(row, 2, 1, 12).setValues([existing]);
  });
  if (missing.length) Logger.log("No Income Statement row found for: " + missing.join(", "));
}

/* ------------------------------ charts --------------------------------- */

function generateCharts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const charts = ss.getSheetByName("Charts");
  if (!charts) { SpreadsheetApp.getUi().alert("Charts tab not found"); return; }
  charts.getCharts().forEach(c => charts.removeChart(c));

  const is = ss.getSheetByName("Income Statement");
  const bal = ss.getSheetByName("Balances");
  const solar = ss.getSheetByName("Solar ROI");
  const ccDetail = ss.getSheetByName("CC Detail");

  const incRow = findRowContains_(is, "total income");
  const expRow = findRowContains_(is, "total expenses");     // matches only "TOTAL EXPENSES"
  const netRow = findRowContains_(is, "net cash flow");
  const assetRow = findRowContains_(bal, "total assets");
  const cumRow = findRowContains_(solar, "cumulative savings");
  const capRow = findRowContains_(solar, "net capital cost");

  charts.getRange("A4:N50").clearContent().clearFormat();
  charts.getRange("A4").setValue("Cash Flow Summary (auto)").setFontWeight("bold");
  charts.getRange("A5").setValue("Month");
  charts.getRange("A6").setValue("Total Income");
  charts.getRange("A7").setValue("Total Expenses");
  charts.getRange("A8").setValue("Net Cash Flow");

  const monthLabels = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  for (let i = 0; i < 12; i++) {
    const L = columnLetter(2 + i);
    charts.getRange(5, 2 + i).setValue(monthLabels[i]);
    charts.getRange(6, 2 + i).setFormula(`='Income Statement'!${L}${incRow}`);
    charts.getRange(7, 2 + i).setFormula(`='Income Statement'!${L}${expRow}`);
    charts.getRange(8, 2 + i).setFormula(`='Income Statement'!${L}${netRow}`);
  }

  charts.insertChart(charts.newChart().asComboChart()
    .addRange(charts.getRange("A5:M8"))
    .setOption("title", "Monthly Cash Flow")
    .setOption("series", { 0: { type: "bars", color: "#2E7D32" }, 1: { type: "bars", color: "#C00000" }, 2: { type: "line", color: "#1F4E78", lineWidth: 3 } })
    .setOption("legend", { position: "bottom" })
    .setPosition(11, 1, 0, 0).build());

  charts.insertChart(charts.newChart().asColumnChart()
    .addRange(ccDetail.getRange("A4:M17"))
    .setOption("title", "Variable Spending by Category").setOption("isStacked", true)
    .setOption("legend", { position: "right" })
    .setPosition(30, 1, 0, 0).build());

  charts.getRange("A28").setValue("Solar Cumulative Savings vs. Capital").setFontWeight("bold");
  charts.getRange("A29").setValue("Month");
  charts.getRange("A30").setValue("Cumulative savings");
  charts.getRange("A31").setValue("Net capital cost (reference)");
  for (let i = 0; i < 12; i++) {
    const L = columnLetter(2 + i);
    charts.getRange(29, 2 + i).setValue(monthLabels[i]);
    charts.getRange(30, 2 + i).setFormula(`='Solar ROI'!${L}${cumRow}`);
    charts.getRange(31, 2 + i).setFormula(`='Solar ROI'!$N$${capRow}`);
  }
  charts.insertChart(charts.newChart().asLineChart()
    .addRange(charts.getRange("A29:M31"))
    .setOption("title", "Solar Payback").setOption("legend", { position: "bottom" })
    .setPosition(50, 1, 0, 0).build());

  charts.getRange("A45").setValue("Total Assets Trend").setFontWeight("bold");
  charts.getRange("A46").setValue("Month");
  charts.getRange("A47").setValue("Total assets");
  for (let i = 0; i < 12; i++) {
    const L = columnLetter(2 + i);
    charts.getRange(46, 2 + i).setValue(monthLabels[i]);
    charts.getRange(47, 2 + i).setFormula(`=Balances!${L}${assetRow}`);
  }
  charts.insertChart(charts.newChart().asLineChart()
    .addRange(charts.getRange("A46:M47"))
    .setOption("title", "Net Worth Trend").setOption("legend", { position: "bottom" })
    .setPosition(70, 1, 0, 0).build());

  SpreadsheetApp.getUi().alert("Charts generated! Scroll down on the Charts tab to see all four.");
}

/* ------------------------------ clears --------------------------------- */

function clearAllImports() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert("Clear All Imports", "This will wipe data in all " + IMPORT_TABS.length + " Imports tabs (rows 5+). Continue?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  IMPORT_TABS.map(t => t.name).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const last = sheet.getLastRow();
    if (last >= 5) { const range = sheet.getRange(5, 1, last - 4, sheet.getLastColumn()); range.clearContent(); range.setBackground(null); }
  });
  ui.alert("Imports cleared.");
}

function clearCCDetail() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert("Clear CC Detail", "This will erase all category numbers in the CC Detail tab. Continue?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("CC Detail").getRange(5, 2, CC_CATS.length, 12).clearContent();
}

/* ---------------------------- roll over -------------------------------- */

function rollOverYear() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert("Roll Over to Next Year",
    "BEFORE running this:\n  1. File → Make a copy. Rename original to 'FIS YYYY (archived)'.\n  2. Run this on the COPY.\n\nThis clears monthly data (Income Statement highlighted cells, CC Detail, Balances → carrying Dec into Jan) and resets the Solar log. Rules are kept. Continue?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const incSheet = ss.getSheetByName("Income Statement");
  const last = incSheet.getLastRow();
  const data = incSheet.getRange(1, 2, last, 12);
  const bgs = data.getBackgrounds();
  const vals = data.getValues();
  bgs.forEach((rowBgs, ri) => rowBgs.forEach((bg, ci) => { if (bg && bg.toUpperCase() === "#FFF2CC") vals[ri][ci] = ""; }));
  data.setValues(vals);

  ss.getSheetByName("CC Detail").getRange(5, 2, CC_CATS.length, 12).clearContent();

  const balSheet = ss.getSheetByName("Balances");
  const decRange = balSheet.getRange(5, 13, 12, 1);
  const decValues = decRange.getValues();
  balSheet.getRange(5, 2, 12, 12).clearContent();
  balSheet.getRange(5, 2, 12, 1).setValues(decValues);

  const solar = ss.getSheetByName("Solar ROI");
  if (solar) {
    const cumRow = findRowContains_(solar, "cumulative savings");
    const decCumulative = cumRow > 0 ? solar.getRange(cumRow, 13).getValue() : 0;
    solar.getRange("B9:M10").clearContent();
    solar.getRange("B5:M5").clearContent();
    solar.getRange("B6:M6").clearContent();
    solar.getRange("A8").setValue("Prior year cumulative savings: $" + Math.round(Number(decCumulative) || 0).toLocaleString())
      .setFontStyle("italic").setFontColor("#808080");
  }

  ui.alert("Year-end rollover complete. Update the title (Income Statement A1) to the new year and verify Jan starting balances.");
}
