/**
 * ============================================================================
 * FIRE Tracker — Shareable Template (Google Apps Script)
 * ============================================================================
 *
 * A personal-finance / FIRE (Financial Independence, Retire Early) workbook
 * generator. Contains NO personal data — running "Setup / Rebuild Template"
 * builds every tab, header, formula, and format from scratch.
 *
 * TABS CREATED
 *   • FIRE Dashboard        — net worth, savings rate, FIRE number, progress
 *   • Income Statement      — monthly income / expenses / net cash flow
 *   • CC Detail             — variable spending by category (auto-populated)
 *   • Amazon Reconciliation — split lumped Amazon charges by itemized orders
 *   • Asset ROI / Reinvest  — ROI & payback for any return-producing asset
 *   • Utility Recon         — manual utility bills vs. what left your account
 *   • Balances              — assets, liabilities, net worth by month
 *   • Imports               — paste raw bank/card transactions here
 *   • Rules                 — regex → category mapping for auto-categorizing
 *   • Charts                — auto-generated charts
 *
 * HOW A NEW USER GETS STARTED
 *   1. Create a blank Google Sheet.
 *   2. Extensions → Apps Script. Select all, delete, paste this whole file.
 *   3. Save (disk icon), then reload the Sheet.
 *   4. Menu "📊 FIRE Tracker" → "Setup / Rebuild Template".
 *   5. Paste transactions into the Imports tab, fill the Rules tab, then run
 *      "Categorize & Roll Up".
 *
 * IMPORTS FORMAT (Imports tab, data starts row 5)
 *   A = Date | B = Description | C = Amount (+ money in, − money out)
 *   D = Account (optional label) | E = Category (auto-filled — leave blank)
 * ============================================================================
 */

/* ----------------------------- CONFIG ----------------------------------- */

const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// Income lines (auto-fillable via Rules token  __INCOME__:<label> )
const INCOME_LINES = [
  "Salary / Wages — 1",
  "Salary / Wages — 2",
  "Interest / Dividends",
  "Other income",
];

// Fixed expense lines (auto-fillable via Rules token  __FIXED__:<label> )
const FIXED_LINES = [
  "Housing (rent / mortgage)",
  "Transportation / Auto",
  "Insurance",
  "Phone / Internet",
  "Debt payments",
  "Investments / Contributions",
];

// Utilities — ALWAYS manual entry. The importer never writes these rows.
const UTILITY_LINES = ["Electric", "Gas / Heating", "Water / Sewer / Trash"];

// Variable spending categories (auto-populated into CC Detail by the importer)
const VARIABLE_CATS = [
  "Groceries", "Dining out", "Fuel / Transport", "Subscriptions",
  "Kids / Activities", "Travel", "Personal Care", "Health / Pharmacy",
  "Home / Hardware", "Shopping (Amazon/misc)", "Entertainment",
  "Gifts / Donations", "Other / misc",
];

// One-time / large items — manual entry
const ONETIME_LINES = ["One-time / large item 1", "One-time / large item 2"];

const COLORS = {
  header:  "#1F4E78",
  section: "#D9E1F2",
  total:   "#EDEDED",
  input:   "#FFF9E6",
  good:    "#C6EFCE",
  goodTxt: "#006100",
  bad:     "#FFC7CE",
  badTxt:  "#9C0006",
};

/* ----------------------------- MENU ------------------------------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 FIRE Tracker")
    .addItem("Setup / Rebuild Template", "setupTemplate")
    .addSeparator()
    .addItem("Categorize & Roll Up", "categorizeAll")
    .addItem("Generate / Refresh Charts", "generateCharts")
    .addSeparator()
    .addItem("Clear Imports", "clearImports")
    .addItem("Roll Over to Next Year", "rollOverYear")
    .addToUi();
}

/* --------------------------- HELPERS ------------------------------------ */

function columnLetter(col) {
  let s = "";
  while (col > 0) { const m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = Math.floor((col - m - 1) / 26); }
  return s;
}

function sheetReset_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) { sh.clear(); sh.getCharts().forEach(c => sh.removeChart(c)); }
  else    { sh = ss.insertSheet(name); }
  return sh;
}

function findRow_(sheet, label) {
  const vals = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  const needle = label.trim().toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === needle) return i + 1;
  }
  return -1;
}

function setNamed_(ss, name, range) { ss.setNamedRange(name, range); }

function getMonthIndex_(dateValue) {
  if (dateValue instanceof Date) return dateValue.getMonth();
  const d = new Date(dateValue);
  return isNaN(d.getTime()) ? null : d.getMonth();
}

// Header row of month labels (+ optional trailing column title) at (row, startCol)
function writeMonthHeader_(sh, row, startCol, trailing) {
  const arr = MONTHS.slice();
  if (trailing) arr.push(trailing);
  const rng = sh.getRange(row, startCol, 1, arr.length);
  rng.setValues([arr]).setFontWeight("bold").setBackground(COLORS.header).setFontColor("#ffffff")
     .setHorizontalAlignment("center");
  return rng;
}

/* ============================ SETUP ===================================== */

function setupTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const year = new Date().getFullYear();

  // Clear any existing named ranges so a rebuild is clean
  ss.getNamedRanges().forEach(nr => nr.remove());

  const is  = buildIncomeStatement_(ss, year);   // returns key row numbers
  buildCCDetail_(ss);
  buildBalances_(ss, year);
  buildDashboard_(ss, year);
  buildAmazonRecon_(ss, year);
  buildAssetROI_(ss, year);
  buildUtilityRecon_(ss, year, is);
  buildImports_(ss);
  buildRules_(ss);
  buildChartsTab_(ss);
  buildReadme_(ss, year);

  // Order the tabs
  const order = ["README","FIRE Dashboard","Income Statement","CC Detail","Balances",
    "Amazon Reconciliation","Asset ROI / Reinvestment","Utility Recon",
    "Imports","Rules","Charts"];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  ss.setActiveSheet(ss.getSheetByName("FIRE Dashboard"));

  ui.alert("Template built.\n\nNext:\n1. Enter Assets/Liabilities on Balances.\n2. Paste transactions into Imports.\n3. Fill the Rules tab.\n4. Run 'Categorize & Roll Up', then 'Generate / Refresh Charts'.");
}

/* --------------------- INCOME STATEMENT --------------------------------- */

function buildIncomeStatement_(ss, year) {
  const sh = sheetReset_(ss, "Income Statement");
  sh.setColumnWidth(1, 250);
  for (let c = 2; c <= 14; c++) sh.setColumnWidth(c, 78);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);

  sh.getRange("A1").setValue("FIRE — Income Statement (" + year + ")")
    .setFontSize(14).setFontWeight("bold");

  const header = ["Category"].concat(MONTHS).concat(["YTD"]);
  sh.getRange(3, 1, 1, header.length).setValues([header])
    .setFontWeight("bold").setBackground(COLORS.header).setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sh.getRange("A3").setHorizontalAlignment("left");

  let r = 4;
  const section = (t) => { sh.getRange(r,1,1,14).setBackground(COLORS.section);
    sh.getRange(r,1).setValue(t).setFontWeight("bold"); r++; };
  const line = (label, opts) => {
    sh.getRange(r,1).setValue("    " + label);
    if (opts && opts.formulaByCol) {
      for (let c = 2; c <= 13; c++) sh.getRange(r,c).setFormula(opts.formulaByCol(c));
    } else if (opts && opts.input) {
      sh.getRange(r,2,1,12).setBackground(COLORS.input);
    }
    sh.getRange(r,14).setFormula("=SUM(B"+r+":M"+r+")");
    r++; return r - 1;
  };
  const totalRow = (label, first, last) => {
    sh.getRange(r,1).setValue(label).setFontWeight("bold");
    sh.getRange(r,1,1,14).setBackground(COLORS.total);
    for (let c = 2; c <= 14; c++) {
      const L = columnLetter(c);
      sh.getRange(r,c).setFormula("=SUM("+L+first+":"+L+last+")").setFontWeight("bold");
    }
    r++; return r - 1;
  };

  // INCOME
  section("INCOME");
  const incFirst = r;
  INCOME_LINES.forEach(l => line(l));
  const incLast = r - 1;
  const totIncome = totalRow("Total Income", incFirst, incLast);
  r++;

  // FIXED
  section("FIXED EXPENSES");
  const fxFirst = r;
  FIXED_LINES.forEach(l => line(l));
  const fxLast = r - 1;
  const totFixed = totalRow("Total Fixed Expenses", fxFirst, fxLast);
  r++;

  // UTILITIES (manual)
  section("UTILITIES (manual entry — not auto-imported)");
  const utFirst = r;
  UTILITY_LINES.forEach(l => line(l, { input: true }));
  const utLast = r - 1;
  const totUtil = totalRow("Total Utilities", utFirst, utLast);
  r++;

  // VARIABLE (formulas → CC Detail)
  section("VARIABLE SPENDING (auto-pulled from CC Detail)");
  const varFirst = r;
  VARIABLE_CATS.forEach((cat, i) => {
    const ccRow = 5 + i; // CC Detail cats start at row 5
    line(cat, { formulaByCol: (c) => "='CC Detail'!" + columnLetter(c) + ccRow });
  });
  const varLast = r - 1;
  const totVar = totalRow("Total Variable Spending", varFirst, varLast);
  r++;

  // ONE-TIME (manual)
  section("ONE-TIME / LARGE ITEMS");
  const otFirst = r;
  ONETIME_LINES.forEach(l => line(l, { input: true }));
  const otLast = r - 1;
  const totOne = totalRow("Total One-Time", otFirst, otLast);
  r++;

  // TOTAL EXPENSES = fixed + utilities + variable + one-time
  sh.getRange(r,1).setValue("TOTAL EXPENSES").setFontWeight("bold");
  sh.getRange(r,1,1,14).setBackground("#F4CCCC");
  for (let c = 2; c <= 14; c++) {
    const L = columnLetter(c);
    sh.getRange(r,c).setFormula("="+L+totFixed+"+"+L+totUtil+"+"+L+totVar+"+"+L+totOne).setFontWeight("bold");
  }
  const totExp = r; r++;

  // NET CASH FLOW = income − expenses
  sh.getRange(r,1).setValue("NET CASH FLOW (Income − Expenses)").setFontWeight("bold");
  sh.getRange(r,1,1,14).setBackground("#D9EAD3");
  for (let c = 2; c <= 14; c++) {
    const L = columnLetter(c);
    sh.getRange(r,c).setFormula("="+L+totIncome+"-"+L+totExp).setFontWeight("bold");
  }
  const netRow = r; r += 2;

  // Number formats
  sh.getRange(4, 2, netRow - 3, 13).setNumberFormat('$#,##0;[Red]($#,##0)');

  // Named ranges used by the Dashboard & Charts
  setNamed_(ss, "Total_Income_YTD",    sh.getRange(totIncome, 14));
  setNamed_(ss, "Total_Expenses_YTD",  sh.getRange(totExp,    14));
  setNamed_(ss, "Net_CashFlow_YTD",    sh.getRange(netRow,    14));
  setNamed_(ss, "Total_Expenses_Row",  sh.getRange(totExp, 2, 1, 12));
  setNamed_(ss, "Total_Income_Row",    sh.getRange(totIncome, 2, 1, 12));

  return { totIncome, totFixed, totUtil, totVar, totOne, totExp, netRow, utFirst, utLast };
}

/* --------------------------- CC DETAIL ---------------------------------- */

function buildCCDetail_(ss) {
  const sh = sheetReset_(ss, "CC Detail");
  sh.setColumnWidth(1, 200);
  sh.getRange("A1").setValue("Variable Spending Detail (auto-populated by Categorize & Roll Up)")
    .setFontSize(13).setFontWeight("bold");
  const header = ["Category"].concat(MONTHS).concat(["YTD"]);
  sh.getRange(4,1,1,header.length).setValues([header])
    .setFontWeight("bold").setBackground(COLORS.header).setFontColor("#ffffff");
  VARIABLE_CATS.forEach((cat, i) => {
    const row = 5 + i;
    sh.getRange(row,1).setValue(cat);
    sh.getRange(row,14).setFormula("=SUM(B"+row+":M"+row+")");
  });
  const last = 5 + VARIABLE_CATS.length - 1;
  sh.getRange(5,2,VARIABLE_CATS.length,13).setNumberFormat('$#,##0;[Red]($#,##0)');
  sh.setFrozenRows(4);
  return { first: 5, last: last };
}

/* --------------------------- BALANCES ----------------------------------- */

function buildBalances_(ss, year) {
  const sh = sheetReset_(ss, "Balances");
  sh.setColumnWidth(1, 240);
  sh.getRange("A1").setValue("Balances — Assets, Liabilities & Net Worth (" + year + ")")
    .setFontSize(13).setFontWeight("bold");
  writeMonthHeader_(sh, 3, 2, "").setValues([MONTHS]);
  sh.getRange(3,1).setValue("Item").setFontWeight("bold").setBackground(COLORS.header).setFontColor("#ffffff");

  let r = 4;
  const section = (t) => { sh.getRange(r,1,1,13).setBackground(COLORS.section);
    sh.getRange(r,1).setValue(t).setFontWeight("bold"); r++; };
  const line = (label) => { sh.getRange(r,1).setValue("    "+label);
    sh.getRange(r,2,1,12).setBackground(COLORS.input); r++; return r-1; };
  const total = (label, first, last) => {
    sh.getRange(r,1).setValue(label).setFontWeight("bold");
    sh.getRange(r,1,1,13).setBackground(COLORS.total);
    for (let c = 2; c <= 13; c++) { const L = columnLetter(c);
      sh.getRange(r,c).setFormula("=SUM("+L+first+":"+L+last+")").setFontWeight("bold"); }
    r++; return r-1;
  };

  section("ASSETS");
  const aFirst = r;
  ["Cash / checking / savings","Taxable investments","Retirement (401k / IRA)",
   "Real estate value","Other assets"].forEach(line);
  const aLast = r - 1;
  const totAssets = total("Total Assets", aFirst, aLast);
  r++;

  section("LIABILITIES");
  const lFirst = r;
  ["Mortgage balance","Auto / student loans","Credit card balances","Other debt"].forEach(line);
  const lLast = r - 1;
  const totLiab = total("Total Liabilities", lFirst, lLast);
  r++;

  sh.getRange(r,1).setValue("NET WORTH").setFontWeight("bold");
  sh.getRange(r,1,1,13).setBackground("#D9EAD3");
  for (let c = 2; c <= 13; c++) { const L = columnLetter(c);
    sh.getRange(r,c).setFormula("="+L+totAssets+"-"+L+totLiab).setFontWeight("bold"); }
  const nwRow = r;

  sh.getRange(4,2,nwRow-3,12).setNumberFormat('$#,##0;[Red]($#,##0)');

  // "Current" (latest non-empty month) helper cells in column O
  sh.getRange(totAssets,15).setFormula('=IFERROR(LOOKUP(2,1/(B'+totAssets+':M'+totAssets+'<>""),B'+totAssets+':M'+totAssets+'),0)');
  sh.getRange(nwRow,15).setFormula('=IFERROR(LOOKUP(2,1/(B'+nwRow+':M'+nwRow+'<>""),B'+nwRow+':M'+nwRow+'),0)');
  sh.getRange(totAssets,15).setNumberFormat('$#,##0');
  sh.getRange(nwRow,15).setNumberFormat('$#,##0');

  setNamed_(ss, "NetWorth_Row",     sh.getRange(nwRow, 2, 1, 12));
  setNamed_(ss, "NetWorth_Current", sh.getRange(nwRow, 15));
  setNamed_(ss, "Assets_Current",   sh.getRange(totAssets, 15));
  sh.setFrozenRows(3); sh.setFrozenColumns(1);
}

/* --------------------------- DASHBOARD ---------------------------------- */

function buildDashboard_(ss, year) {
  const sh = sheetReset_(ss, "FIRE Dashboard");
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 160);
  sh.getRange("A1").setValue("🔥 FIRE Dashboard (" + year + ")").setFontSize(16).setFontWeight("bold");
  sh.getRange("A2").setValue("All figures pull live from the Income Statement & Balances tabs.")
    .setFontStyle("italic").setFontColor("#666666");

  const kv = (row, label, formula, fmt, note) => {
    sh.getRange(row,1).setValue(label).setFontWeight("bold");
    const c = sh.getRange(row,2).setFormula(formula);
    if (fmt) c.setNumberFormat(fmt);
    if (note) sh.getRange(row,3).setValue(note).setFontStyle("italic").setFontColor("#888888");
  };

  // Assumptions (editable inputs)
  sh.getRange("A4").setValue("ASSUMPTIONS").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange("B4").setBackground(COLORS.section);
  sh.getRange("A5").setValue("Safe withdrawal rate (SWR)");
  sh.getRange("B5").setValue(0.04).setNumberFormat("0.0%").setBackground(COLORS.input);
  sh.getRange("A6").setValue("Current age");
  sh.getRange("B6").setValue(35).setBackground(COLORS.input);
  sh.getRange("A7").setValue("Target FIRE age");
  sh.getRange("B7").setValue(55).setBackground(COLORS.input);
  sh.getRange("A8").setValue("Expected real return (for Coast FIRE)");
  sh.getRange("B8").setValue(0.05).setNumberFormat("0.0%").setBackground(COLORS.input);

  // Spending & savings
  sh.getRange("A10").setValue("SPENDING & SAVINGS").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange("B10").setBackground(COLORS.section);
  kv(11, "Expenses YTD", "=Total_Expenses_YTD", "$#,##0");
  kv(12, "Avg monthly expenses", '=IFERROR(AVERAGEIF(Total_Expenses_Row,">0"),0)', "$#,##0", "months with data only");
  kv(13, "Estimated annual expenses", "=B12*12", "$#,##0", "avg monthly × 12");
  kv(14, "Savings rate (YTD)", "=IFERROR(Net_CashFlow_YTD/Total_Income_YTD,0)", "0.0%", "net cash flow ÷ income");

  // FIRE metrics
  sh.getRange("A16").setValue("FIRE PROGRESS").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange("B16").setBackground(COLORS.section);
  kv(17, "Current net worth", "=NetWorth_Current", "$#,##0");
  kv(18, "FIRE number", "=IFERROR(B13/B5,0)", "$#,##0", "annual spend ÷ SWR (the 25× rule at 4%)");
  kv(19, "Progress to FI", "=IFERROR(B17/B18,0)", "0.0%");
  kv(20, "Remaining to FI", "=MAX(B18-B17,0)", "$#,##0");
  kv(21, "Safe annual withdrawal now", "=B17*B5", "$#,##0", "what your current NW would throw off");
  kv(22, "Coast FIRE number (today)", "=IFERROR(B18/((1+B8)^(B7-B6)),0)", "$#,##0",
     "NW that, untouched, grows to your FIRE # by target age");
  kv(23, "Coast FIRE reached?", '=IF(B17>=B22,"YES ✔","not yet")', null);

  // Progress bar (visual)
  sh.getRange("A25").setValue("Progress to FI");
  sh.getRange("B25").setFormula('=SPARKLINE(MIN(B19,1),{"charttype","bar";"max",1;"color1","#2E7D32"})');

  sh.getRange("A27").setValue("Tip: keep Balances updated monthly — the dashboard reads the latest filled month.")
    .setFontStyle("italic").setFontColor("#888888");
  sh.setFrozenRows(2);
}

/* --------------------- AMAZON RECONCILIATION ---------------------------- */

function buildAmazonRecon_(ss, year) {
  const sh = sheetReset_(ss, "Amazon Reconciliation");
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 300);
  sh.setColumnWidth(4, 170); sh.setColumnWidth(5, 100);
  sh.setColumnWidth(7, 190); sh.setColumnWidth(8, 120);

  sh.getRange("A1").setValue("Amazon Reconciliation").setFontSize(14).setFontWeight("bold");
  sh.getRange("A2").setValue("Split lumped Amazon card charges into real categories by matching your itemized order history.")
    .setFontStyle("italic").setFontColor("#666666");

  // Match keyword input
  sh.getRange("A3").setValue("Match keyword in Imports:");
  sh.getRange("B3").setValue("AMAZON").setBackground(COLORS.input);
  sh.getRange("C3").setValue("(descriptions containing this are treated as Amazon; AMZN is also matched)")
     .setFontStyle("italic").setFontColor("#888888");

  // Section 1 — statement charges (auto from Imports)
  sh.getRange("A5").setValue("STATEMENT CHARGES (auto from Imports)").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange(5,2,1,12).setBackground(COLORS.section);
  writeMonthHeader_(sh, 6, 2, "");
  sh.getRange(6,2,1,12).setValues([MONTHS]);
  sh.getRange("A7").setValue("Amazon charged to cards");
  for (let c = 2; c <= 13; c++) {
    const L = columnLetter(c);
    const m = c - 1;
    const start = 'DATE(' + year + ',' + m + ',1)';
    const f =
      '=-SUMIFS(Imports!$C$5:$C,Imports!$B$5:$B,"*"&$B$3&"*",Imports!$A$5:$A,">="&' + start +
      ',Imports!$A$5:$A,"<"&EDATE(' + start + ',1))' +
      '-SUMIFS(Imports!$C$5:$C,Imports!$B$5:$B,"*AMZN*",Imports!$A$5:$A,">="&' + start +
      ',Imports!$A$5:$A,"<"&EDATE(' + start + ',1))';
    sh.getRange(7,c).setFormula(f);
  }
  sh.getRange(7,14).setFormula("=SUM(B7:M7)");
  sh.getRange(7,2,1,13).setNumberFormat('$#,##0.00');

  // Section 2 — itemized order entry
  sh.getRange("A9").setValue("ITEMIZED ORDERS (paste your Amazon order history below)").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange(9,2,1,4).setBackground(COLORS.section);
  const hdr = ["Order date","Order #","Item / description","Category","Amount"];
  sh.getRange(10,1,1,5).setValues([hdr]).setFontWeight("bold").setBackground(COLORS.header).setFontColor("#ffffff");
  // category dropdown validation for the order rows
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(VARIABLE_CATS, true).build();
  sh.getRange(11,4,500,1).setDataValidation(rule);
  sh.getRange(11,1,500,1).setNumberFormat('m/d/yyyy');
  sh.getRange(11,5,500,1).setNumberFormat('$#,##0.00');

  // Summary block (right side)
  sh.getRange("G5").setValue("RECONCILIATION").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange("H5").setBackground(COLORS.section);
  sh.getRange("G6").setValue("Itemized orders total");
  sh.getRange("H6").setFormula("=SUM(E11:E)").setNumberFormat('$#,##0.00');
  sh.getRange("G7").setValue("Statement charges total (YTD)");
  sh.getRange("H7").setFormula("=N7").setNumberFormat('$#,##0.00');
  sh.getRange("G8").setValue("Unreconciled difference");
  sh.getRange("H8").setFormula("=H7-H6").setNumberFormat('$#,##0.00');
  sh.getRange("G9").setValue("Status");
  sh.getRange("H9").setFormula('=IF(ABS(H8)<1,"Reconciled ✔","Unreconciled: "&TEXT(H8,"$#,##0.00"))');

  sh.getRange("G11").setValue("By category (from itemized orders):").setFontWeight("bold");
  VARIABLE_CATS.forEach((cat, i) => {
    const row = 12 + i;
    sh.getRange(row,7).setValue(cat);
    sh.getRange(row,8).setFormula('=SUMIF($D$11:$D,G'+row+',$E$11:$E)').setNumberFormat('$#,##0.00');
  });

  // Conditional format on status
  const rng = sh.getRange("H9");
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith("Unreconciled").setBackground(COLORS.bad).setFontColor(COLORS.badTxt).setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith("Reconciled").setBackground(COLORS.good).setFontColor(COLORS.goodTxt).setRanges([rng]).build());
  sh.setConditionalFormatRules(rules);
}

/* --------------------- ASSET ROI / REINVESTMENT ------------------------- */

function buildAssetROI_(ss, year) {
  const sh = sheetReset_(ss, "Asset ROI / Reinvestment");
  sh.setColumnWidth(1, 280);
  sh.getRange("A1").setValue("Asset ROI / Reinvestment").setFontSize(14).setFontWeight("bold");
  sh.getRange("A2").setValue("Works for any return-producing asset: solar, rental property, a business, dividend portfolio, etc.")
    .setFontStyle("italic").setFontColor("#666666");

  // Inputs
  sh.getRange("A4").setValue("ASSET INPUTS").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange("B4").setBackground(COLORS.section);
  const inp = (row, label, val, fmt) => {
    sh.getRange(row,1).setValue(label);
    const c = sh.getRange(row,2).setValue(val).setBackground(COLORS.input);
    if (fmt) c.setNumberFormat(fmt);
  };
  inp(5, "Asset name", "e.g., Rental @ 123 Main St");
  inp(6, "Asset type", "Rental property");
  inp(7, "Initial capital (net of incentives)", 0, "$#,##0");
  inp(8, "Expected annual return ($)", 0, "$#,##0");
  sh.getRange("A9").setValue("Expected annual return (%)");
  sh.getRange("B9").setFormula("=IFERROR(B8/B7,0)").setNumberFormat("0.0%");
  inp(10, "In-service / purchase date", new Date(year,0,1), "m/d/yyyy");

  // Monthly actuals
  sh.getRange("A12").setValue("MONTHLY RETURN / INCOME (enter actuals)").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange(12,2,1,13).setBackground(COLORS.section);
  writeMonthHeader_(sh, 13, 2, "YTD");
  sh.getRange("A14").setValue("Monthly return / income ($)");
  sh.getRange(14,2,1,12).setBackground(COLORS.input);
  sh.getRange(14,14).setFormula("=SUM(B14:M14)");
  sh.getRange("A15").setValue("Cumulative return ($)");
  sh.getRange(15,2).setFormula("=B14");
  for (let c = 3; c <= 13; c++) {
    const L = columnLetter(c), P = columnLetter(c-1);
    sh.getRange(15,c).setFormula("="+P+"15+"+L+"14");
  }
  sh.getRange(14,2,2,13).setNumberFormat('$#,##0');

  // Metrics
  sh.getRange("A17").setValue("ROI METRICS").setFontWeight("bold").setBackground(COLORS.section);
  sh.getRange("B17").setBackground(COLORS.section);
  const met = (row, label, formula, fmt, note) => {
    sh.getRange(row,1).setValue(label).setFontWeight("bold");
    sh.getRange(row,2).setFormula(formula).setNumberFormat(fmt || '$#,##0');
    if (note) sh.getRange(row,3).setValue(note).setFontStyle("italic").setFontColor("#888888");
  };
  met(18, "Cumulative return to date", '=IFERROR(LOOKUP(2,1/(B15:M15<>""),B15:M15),0)');
  met(19, "ROI to date (%)", "=IFERROR(B18/B7,0)", "0.0%");
  met(20, "Net position (cumulative − capital)", "=B18-B7");
  met(21, "Simple payback (years)", "=IFERROR(B7/B8,0)", "0.0", "capital ÷ expected annual return");
  met(22, "Est. break-even date", '=IFERROR(EDATE(B10,ROUNDUP(B7/(B8/12),0)),"—")', "m/d/yyyy");

  // Reference cell for the payback chart (net capital, negative)
  sh.getRange("O7").setFormula("=-B7").setNumberFormat('$#,##0');
  setNamed_(ss, "Asset_Capital_Ref", sh.getRange("O7"));
  setNamed_(ss, "Asset_Cumulative_Row", sh.getRange(15,2,1,12));
}

/* ------------------------- UTILITY RECON -------------------------------- */

function buildUtilityRecon_(ss, year, is) {
  const sh = sheetReset_(ss, "Utility Recon");
  sh.setColumnWidth(1, 230);
  sh.getRange("A1").setValue("Utility Reconciliation").setFontSize(14).setFontWeight("bold");
  sh.getRange("A2").setValue("Confirms the utility bills you enter manually match what actually left your account.")
    .setFontStyle("italic").setFontColor("#666666");
  sh.getRange("A3").setValue("Match keyword in Imports:");
  sh.getRange("B3").setValue("ELECTRIC").setBackground(COLORS.input);
  sh.getRange("C3").setValue('e.g., your utility name. Sums Imports outflows whose description contains this.')
    .setFontStyle("italic").setFontColor("#888888");

  writeMonthHeader_(sh, 5, 2, "YTD");
  sh.getRange(5,1).setValue("").setBackground(COLORS.header);

  // Manual (from Income Statement utility rows — Electric + Gas)
  const eRow = is.utFirst;       // Electric
  const gRow = is.utFirst + 1;   // Gas / Heating
  sh.getRange("A6").setValue("Manual (Electric + Gas)");
  for (let c = 2; c <= 13; c++) {
    const L = columnLetter(c);
    sh.getRange(6,c).setFormula("='Income Statement'!"+L+eRow+"+'Income Statement'!"+L+gRow);
  }
  sh.getRange(6,14).setFormula("=SUM(B6:M6)");

  // Bank (from Imports, matching keyword; outflows are negative → negate)
  sh.getRange("A7").setValue("From Imports (actual outflow)");
  for (let c = 2; c <= 13; c++) {
    const L = columnLetter(c), m = c - 1;
    const start = 'DATE(' + year + ',' + m + ',1)';
    const f = '=-SUMIFS(Imports!$C$5:$C,Imports!$B$5:$B,"*"&$B$3&"*",Imports!$A$5:$A,">="&' +
      start + ',Imports!$A$5:$A,"<"&EDATE(' + start + ',1))';
    sh.getRange(7,c).setFormula(f);
  }
  sh.getRange(7,14).setFormula("=SUM(B7:M7)");

  sh.getRange("A8").setValue("Difference (manual − actual)");
  for (let c = 2; c <= 14; c++) { const L = columnLetter(c);
    sh.getRange(8,c).setFormula("="+L+"6-"+L+"7"); }

  sh.getRange("A9").setValue("Flag");
  for (let c = 2; c <= 13; c++) { const L = columnLetter(c);
    sh.getRange(9,c).setFormula('=IF(AND('+L+'6=0,'+L+'7=0),"",IF(ABS('+L+'8)<1,"OK","CHECK "&TEXT('+L+'8,"$+#,##0;$-#,##0")))'); }
  sh.getRange(9,14).setFormula('=IF(ABS(N8)<1,"OK","CHECK — off by "&TEXT(N8,"$#,##0"))');

  sh.getRange(6,2,3,13).setNumberFormat('$#,##0');

  // Conditional format: CHECK cells red, OK green
  const flags = sh.getRange("B9:N9");
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH(B9&"","^CHECK")')
    .setBackground(COLORS.bad).setFontColor(COLORS.badTxt).setRanges([flags]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH(B9&"","^OK")')
    .setBackground(COLORS.good).setFontColor(COLORS.goodTxt).setRanges([flags]).build());
  sh.setConditionalFormatRules(rules);

  sh.getRange("A11").setValue("Note: the YTD flag (column N) is the one to trust — monthly gaps often just reflect billing dates that drift across month boundaries and wash out over the year.")
    .setFontStyle("italic").setFontColor("#888888");
}

/* ---------------------------- IMPORTS ----------------------------------- */

function buildImports_(ss) {
  const sh = sheetReset_(ss, "Imports");
  sh.setColumnWidth(1, 100); sh.setColumnWidth(2, 340); sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 160);
  sh.getRange("A1").setValue("Imports — paste raw transactions below (data starts row 5)").setFontSize(13).setFontWeight("bold");
  sh.getRange("A2").setValue("Sign convention: Amount is POSITIVE for money in, NEGATIVE for money out.")
    .setFontStyle("italic").setFontColor("#666666");
  const hdr = ["Date","Description","Amount","Account","Category (auto)"];
  sh.getRange(4,1,1,5).setValues([hdr]).setFontWeight("bold").setBackground(COLORS.header).setFontColor("#ffffff");
  sh.getRange(5,1,1000,1).setNumberFormat('m/d/yyyy');
  sh.getRange(5,3,1000,1).setNumberFormat('$#,##0.00;[Red]($#,##0.00)');
  sh.setFrozenRows(4);
}

/* ----------------------------- RULES ------------------------------------ */

function buildRules_(ss) {
  const sh = sheetReset_(ss, "Rules");
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 380);
  sh.getRange("A1").setValue("Categorization Rules").setFontSize(13).setFontWeight("bold");
  sh.getRange("A2").setValue("First matching rule wins (top to bottom). Pattern is a regex tested against the Description (case-insensitive).")
    .setFontStyle("italic").setFontColor("#666666");
  sh.getRange("A3").setValue("Category tokens:  __INCOME__:<line>   __FIXED__:<line>   __TRANSFER__ (ignore)   __MANUAL__ (ignore — you enter it)   or a variable category name.")
    .setFontStyle("italic").setFontColor("#888888");
  sh.getRange(4,1,1,3).setValues([["Pattern (regex)","Category token","Notes"]])
    .setFontWeight("bold").setBackground(COLORS.header).setFontColor("#ffffff");

  const examples = [
    ["PAYROLL|DIRECT DEP",         "__INCOME__:Salary / Wages — 1", "Map your paychecks to an income line"],
    ["INTEREST",                    "__INCOME__:Interest / Dividends", ""],
    ["MORTGAGE|RENT",               "__FIXED__:Housing (rent / mortgage)", ""],
    ["TOYOTA|HONDA|AUTO LOAN",      "__FIXED__:Transportation / Auto", ""],
    ["INSUR",                       "__FIXED__:Insurance", ""],
    ["VERIZON|T-MOBILE|COMCAST",    "__FIXED__:Phone / Internet", ""],
    ["ELECTRIC|GAS CO|WATER",       "__MANUAL__", "Utilities are entered manually — ignored on import"],
    ["TRANSFER|XFER|SAVINGS",       "__TRANSFER__", "Moves between your own accounts — ignored"],
    ["WHOLE FOODS|SAFEWAY|KROGER",  "Groceries", ""],
    ["AMAZON|AMZN",                 "Shopping (Amazon/misc)", "Reconcile in the Amazon tab"],
    ["STARBUCKS|RESTAURANT|DOORDASH","Dining out", ""],
    ["SHELL|EXXON|CHEVRON",         "Fuel / Transport", ""],
    ["NETFLIX|SPOTIFY|HULU",        "Subscriptions", ""],
  ];
  sh.getRange(5,1,examples.length,3).setValues(examples);
  sh.setFrozenRows(4);
}

/* ----------------------------- CHARTS ----------------------------------- */

function buildChartsTab_(ss) {
  const sh = sheetReset_(ss, "Charts");
  sh.getRange("A1").setValue("Charts — run '📊 FIRE Tracker → Generate / Refresh Charts'")
    .setFontSize(13).setFontWeight("bold");
}

/* ----------------------------- README ----------------------------------- */

function buildReadme_(ss, year) {
  const sh = sheetReset_(ss, "README");
  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 820);
  sh.setHiddenGridlines(true);

  // Rows of [style, text]. style: title|h2|body|bullet|mono|space
  const rows = [
    ["title", "🔥 FIRE Tracker — Read Me First"],
    ["body",  "A privacy-safe personal-finance / FIRE workbook. Everything below is built by the script — no personal data is included in the template."],
    ["space", ""],

    ["h2",    "Quick start"],
    ["bullet","1.  Menu \u201c\ud83d\udcca FIRE Tracker \u2192 Setup / Rebuild Template\u201d built every tab. (You already did this.)"],
    ["bullet","2.  Fill the Balances tab with your assets & liabilities for each month."],
    ["bullet","3.  Paste transactions into the Imports tab (format below)."],
    ["bullet","4.  Review/extend the Rules tab so transactions get categorized."],
    ["bullet","5.  Run \u201cCategorize & Roll Up\u201d, then \u201cGenerate / Refresh Charts\u201d."],
    ["bullet","6.  Watch the FIRE Dashboard update as you go."],
    ["space", ""],

    ["h2",    "The menu (\ud83d\udcca FIRE Tracker)"],
    ["bullet","Setup / Rebuild Template \u2014 (re)builds all tabs. Do NOT run once you have real data; it wipes everything."],
    ["bullet","Categorize & Roll Up \u2014 reads Imports, applies Rules, fills CC Detail + income/fixed rows."],
    ["bullet","Generate / Refresh Charts \u2014 rebuilds the four charts on the Charts tab."],
    ["bullet","Clear Imports \u2014 empties the Imports tab (row 5 down)."],
    ["bullet","Roll Over to Next Year \u2014 run on a COPY: clears monthly data, carries Dec balances \u2192 Jan."],
    ["space", ""],

    ["h2",    "Imports format (data starts on row 5)"],
    ["mono",  "A = Date   |   B = Description   |   C = Amount   |   D = Account (optional)   |   E = Category (auto)"],
    ["bullet","Amount sign: POSITIVE = money in, NEGATIVE = money out. Leave column E blank \u2014 the script fills it."],
    ["space", ""],

    ["h2",    "How categorization works (Rules tab)"],
    ["body",  "Each rule is a regex pattern matched against the Description (case-insensitive). First match wins, top to bottom. The Category token decides where the amount lands:"],
    ["mono",  "__INCOME__:<line>    \u2192 an income line on the Income Statement"],
    ["mono",  "__FIXED__:<line>     \u2192 a fixed-expense line on the Income Statement"],
    ["mono",  "<Variable category>  \u2192 a row on CC Detail (e.g., Groceries, Dining out)"],
    ["mono",  "__TRANSFER__         \u2192 ignored (moves between your own accounts)"],
    ["mono",  "__MANUAL__           \u2192 ignored (you type it yourself \u2014 e.g., utilities)"],
    ["body",  "Anything that matches no rule is counted as \u201cOther / misc\u201d variable spending."],
    ["space", ""],

    ["h2",    "Why utilities are entered by hand"],
    ["body",  "Gas & electric often bill together (one payee) and can\u2019t be split reliably from a bank feed, so the Utility rows on the Income Statement are MANUAL and the importer never overwrites them. The Utility Recon tab compares what you typed against what actually left your account and flags gaps (trust the YTD flag \u2014 monthly gaps are usually just billing-date drift)."],
    ["space", ""],

    ["h2",    "Tab guide"],
    ["bullet","FIRE Dashboard \u2014 net worth, savings rate, FIRE number, progress, Coast FIRE. Edit the yellow assumption cells (SWR, ages, return)."],
    ["bullet","Income Statement \u2014 monthly income, expenses, net cash flow. Yellow cells are manual; totals are formulas."],
    ["bullet","CC Detail \u2014 variable spending by category (auto-filled)."],
    ["bullet","Balances \u2014 assets, liabilities, net worth by month (fill the yellow cells)."],
    ["bullet","Amazon Reconciliation \u2014 split lumped Amazon charges by pasting itemized orders."],
    ["bullet","Asset ROI / Reinvestment \u2014 ROI & payback for any return-producing asset (rental, solar, business...)."],
    ["bullet","Utility Recon \u2014 manual utility bills vs. actual outflow, with red/green flags."],
    ["bullet","Imports \u2014 paste raw transactions here."],
    ["bullet","Rules \u2014 the regex \u2192 category map."],
    ["bullet","Charts \u2014 auto-generated visuals."],
    ["space", ""],

    ["h2",    "Colors & conventions"],
    ["bullet","Yellow cells = you type here.   Grey/colored rows = formulas, leave them alone."],
    ["bullet","Green \u201cOK\u201d / red \u201cCHECK\u201d flags appear on the recon tabs."],
    ["space", ""],

    ["h2",    "Privacy & safety"],
    ["bullet","This script only edits THIS spreadsheet. No external accounts, API keys, or OAuth apps are needed."],
    ["bullet","First run asks you to authorize it to manage this sheet \u2014 that\u2019s normal for bound Apps Scripts."],
    ["bullet","The template ships with zero personal data, so it\u2019s safe to share the blank version with others."],
    ["space", ""],

    ["body",  "Template generated for " + year + ". Rebuild anytime from the menu."],
  ];

  let r = 1;
  rows.forEach(([style, text]) => {
    const cell = sh.getRange(r, 2).setValue(text).setVerticalAlignment("middle");
    switch (style) {
      case "title": cell.setFontSize(18).setFontWeight("bold").setFontColor(COLORS.header); sh.setRowHeight(r, 34); break;
      case "h2":    cell.setFontSize(13).setFontWeight("bold").setFontColor("#ffffff");
                    sh.getRange(r,1,1,2).setBackground(COLORS.header); sh.setRowHeight(r, 26); break;
      case "mono":  cell.setFontFamily("Courier New").setFontSize(10).setBackground("#F3F3F3"); break;
      case "bullet":cell.setFontSize(11); break;
      case "space": sh.setRowHeight(r, 8); break;
      default:      cell.setFontSize(11).setWrap(true); break;
    }
    r++;
  });

  sh.setFrozenRows(1);
}

/* ======================= OPERATIONAL FUNCTIONS ========================== */

function loadRules_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Rules");
  if (!sheet) throw new Error("Rules tab not found — run Setup first.");
  const last = sheet.getLastRow();
  if (last < 5) return [];
  return sheet.getRange(5,1,last-4,2).getValues()
    .filter(r => r[0] && r[1])
    .map(r => ({ regex: new RegExp(String(r[0]), "i"), category: String(r[1]).trim() }));
}

function categorize_(description, rules) {
  if (!description) return { category: "Other / misc", matched: false };
  for (const rule of rules) {
    if (rule.regex.test(description)) return { category: rule.category, matched: true };
  }
  return { category: "Other / misc", matched: false };
}

function categorizeAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let rules;
  try { rules = loadRules_(); } catch (e) { ui.alert(e.message); return; }

  const imp = ss.getSheetByName("Imports");
  if (!imp) { ui.alert("Imports tab not found — run Setup first."); return; }
  const last = imp.getLastRow();
  if (last < 5) { ui.alert("No transactions found in Imports."); return; }

  const n = last - 4;
  const data = imp.getRange(5,1,n,4).getValues(); // A:date B:desc C:amount D:account
  const catCol = [];

  const monthlyVar = {}, monthlyInc = {}, monthlyFix = {};
  const bucket = (obj, m, key, val) => { (obj[m] = obj[m] || {})[key] = (obj[m][key] || 0) + val; };

  data.forEach(row => {
    const [date, desc, amount] = row;
    if (!date || desc === "" || amount === "") { catCol.push([""]); return; }
    const m = getMonthIndex_(date);
    if (m === null) { catCol.push([""]); return; }

    const res = categorize_(desc, rules);
    const cat = res.category;
    const amt = Number(amount);

    if (cat === "__TRANSFER__" || cat === "__MANUAL__") {
      // ignored — transfers, and manually-entered categories (utilities)
    } else if (cat.indexOf("__INCOME__:") === 0) {
      bucket(monthlyInc, m, cat.replace("__INCOME__:", ""), amt);            // inflow positive
    } else if (cat.indexOf("__FIXED__:") === 0) {
      bucket(monthlyFix, m, cat.replace("__FIXED__:", ""), Math.abs(amt));   // expense magnitude
    } else {
      bucket(monthlyVar, m, cat, -amt);                                      // outflow negative → positive spend
    }
    catCol.push([cat]);
  });

  // Write category column back to Imports (col E)
  imp.getRange(5,5,n,1).setValues(catCol);

  // Write CC Detail grid
  const cc = ss.getSheetByName("CC Detail");
  const ccGrid = VARIABLE_CATS.map(cat => {
    const r = new Array(12).fill("");
    for (let m = 0; m < 12; m++) { const v = (monthlyVar[m] && monthlyVar[m][cat]) || 0; if (v) r[m] = v; }
    return r;
  });
  cc.getRange(5,2,VARIABLE_CATS.length,12).setValues(ccGrid);

  // Write Income & Fixed rows on Income Statement (never touch totals or utilities)
  const is = ss.getSheetByName("Income Statement");
  const writeRows = (bucketObj, labels) => {
    labels.forEach(label => {
      const row = findRow_(is, "    " + label);
      if (row < 0) return;
      const existing = is.getRange(row,2,1,12).getValues()[0];
      for (let m = 0; m < 12; m++) {
        if (bucketObj[m] && bucketObj[m][label] !== undefined) existing[m] = bucketObj[m][label];
      }
      is.getRange(row,2,1,12).setValues([existing]);
    });
  };
  writeRows(monthlyInc, INCOME_LINES);
  writeRows(monthlyFix, FIXED_LINES);

  ui.alert("Categorize complete — " + n + " rows processed.\n\nVariable spending → CC Detail.\nIncome & fixed → Income Statement.\nUtilities, transfers & manual items were left untouched.");
}

function generateCharts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const charts = ss.getSheetByName("Charts");
  const is = ss.getSheetByName("Income Statement");
  const bal = ss.getSheetByName("Balances");
  const cc = ss.getSheetByName("CC Detail");
  const asset = ss.getSheetByName("Asset ROI / Reinvestment");
  if (!charts) { ui.alert("Charts tab not found — run Setup first."); return; }
  charts.getCharts().forEach(c => charts.removeChart(c));
  charts.getRange("A4:N60").clearContent();

  const incRow = findRow_(is, "Total Income");
  const expRow = findRow_(is, "TOTAL EXPENSES");
  const netRow = findRow_(is, "NET CASH FLOW (Income − Expenses)");
  const nwRow  = findRow_(bal, "NET WORTH");

  // Cash-flow source block
  charts.getRange("A4").setValue("Month");
  charts.getRange("A5").setValue("Total Income");
  charts.getRange("A6").setValue("Total Expenses");
  charts.getRange("A7").setValue("Net Cash Flow");
  charts.getRange("A8").setValue("Net Worth");
  for (let i = 0; i < 12; i++) {
    const L = columnLetter(2 + i);
    charts.getRange(4, 2+i).setValue(MONTHS[i]);
    charts.getRange(5, 2+i).setFormula("='Income Statement'!"+L+incRow);
    charts.getRange(6, 2+i).setFormula("='Income Statement'!"+L+expRow);
    charts.getRange(7, 2+i).setFormula("='Income Statement'!"+L+netRow);
    charts.getRange(8, 2+i).setFormula("=Balances!"+L+nwRow);
  }

  charts.insertChart(charts.newChart().asComboChart()
    .addRange(charts.getRange("A4:M7"))
    .setOption("title","Monthly Cash Flow")
    .setOption("series",{0:{type:"bars",color:"#2E7D32"},1:{type:"bars",color:"#C00000"},2:{type:"line",color:"#1F4E78",lineWidth:3}})
    .setOption("legend",{position:"bottom"})
    .setPosition(11,1,0,0).build());

  charts.insertChart(charts.newChart().asLineChart()
    .addRange(charts.getRange("A4:M4")).addRange(charts.getRange("A8:M8"))
    .setOption("title","Net Worth Trend").setOption("legend",{position:"bottom"})
    .setPosition(30,1,0,0).build());

  charts.insertChart(charts.newChart().asColumnChart()
    .addRange(cc.getRange(4,1,1,13)).addRange(cc.getRange(5,1,VARIABLE_CATS.length,13))
    .setOption("title","Variable Spending by Category").setOption("isStacked",true)
    .setOption("legend",{position:"right"}).setPosition(49,1,0,0).build());

  // Asset cumulative return vs capital
  charts.getRange("A46").setValue("Asset cumulative");
  charts.getRange("A47").setValue("Capital (ref)");
  for (let i = 0; i < 12; i++) {
    const L = columnLetter(2 + i);
    charts.getRange(46,2+i).setFormula("='Asset ROI / Reinvestment'!"+L+"15");
    charts.getRange(47,2+i).setFormula("=Asset_Capital_Ref");
  }
  charts.insertChart(charts.newChart().asLineChart()
    .addRange(charts.getRange("A4:M4")).addRange(charts.getRange("A46:M47"))
    .setOption("title","Asset Payback").setOption("legend",{position:"bottom"})
    .setPosition(68,1,0,0).build());

  ui.alert("Charts generated. Scroll down the Charts tab to see all four.");
}

function clearImports() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert("Clear Imports", "Wipe all rows in the Imports tab (row 5+)?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Imports");
  const last = sh.getLastRow();
  if (last >= 5) sh.getRange(5,1,last-4,sh.getLastColumn()).clearContent();
  ui.alert("Imports cleared.");
}

function rollOverYear() {
  const ui = SpreadsheetApp.getUi();
  const msg = "BEFORE running: File → Make a copy, and run this on the COPY (next year's sheet).\n\n" +
    "This clears monthly data in Income Statement, CC Detail, Balances (carrying Dec → Jan), " +
    "and Asset actuals. Rules are kept. Continue?";
  if (ui.alert("Roll Over to Next Year", msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Income Statement: clear input (non-formula) numeric cells in B:M
  const is = ss.getSheetByName("Income Statement");
  const isRange = is.getRange(4,2,is.getLastRow()-3,12);
  const formulas = isRange.getFormulas();
  const vals = isRange.getValues();
  for (let i = 0; i < vals.length; i++)
    for (let j = 0; j < vals[i].length; j++)
      if (!formulas[i][j]) vals[i][j] = "";
  isRange.setValues(vals);

  // CC Detail
  ss.getSheetByName("CC Detail").getRange(5,2,VARIABLE_CATS.length,12).clearContent();

  // Balances: carry Dec (col M / 13) → Jan (col B / 2) for input rows, clear rest
  const bal = ss.getSheetByName("Balances");
  const bRange = bal.getRange(4,2,bal.getLastRow()-3,12);
  const bForm = bRange.getFormulas();
  const dec = bal.getRange(4,13,bal.getLastRow()-3,1).getValues();
  const bVals = bRange.getValues();
  for (let i = 0; i < bVals.length; i++) {
    for (let j = 0; j < bVals[i].length; j++) if (!bForm[i][j]) bVals[i][j] = "";
    if (!bForm[i][0] && dec[i][0] !== "") bVals[i][0] = dec[i][0]; // Jan = prior Dec
  }
  bRange.setValues(bVals);

  // Asset actuals
  ss.getSheetByName("Asset ROI / Reinvestment").getRange(14,2,1,12).clearContent();

  ui.alert("Rollover complete. Update the year in the tab titles and verify Jan starting balances.");
}
