/**
 * MonthlyEntry.gs — guided monthly data entry for the Family Income Statement.
 *
 * WHAT IT DOES
 *   1) "Post Monthly Entry" reads the 'Monthly Entry' tab, and for the month you
 *      picked in B4, writes each filled-in value to the correct cell across
 *      Balances / Solar Reinvestment / Solar ROI / Amazon Reconciliation / Comp & Benefits.
 *      Blank staging cells are skipped, so you only overwrite what you actually entered.
 *   2) "Create Monthly Form" (optional) builds a Google Form with the same questions.
 *      Form submissions flow into the 'Monthly Entry' tab; then Post Monthly Entry fans them out.
 *
 * SETUP
 *   - Extensions → Apps Script → add this file (alongside your existing script).
 *   - If you already have an onOpen(), merge the menu items below into it.
 *   - Reload the sheet; use the "📊 Family Finances" menu.
 *
 * SAFETY
 *   - Review the MAP below once — it is the single source of truth for where each
 *     field lands. Nothing is written unless the staging cell has a value.
 */

var ENTRY_SHEET = 'Monthly Entry';
var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/**
 * MAP entries describe where each staging cell goes.
 *   stage : cell on 'Monthly Entry' holding the value you typed
 *   sheet : destination tab
 *   kind  : 'monthCol'  -> destination row is fixed, column = the chosen month (B..M)
 *           'monthRow'  -> destination column is fixed, row = 14 + monthIndex (Solar log)
 *           'fixed'     -> a single fixed destination cell (YTD values)
 *   row/col: destination coordinates as needed by 'kind' (1-based)
 */
var MAP = [
  // (1) Balances (month-end) -- months are columns B..M
  {stage:'B7',  sheet:'Balances', kind:'monthCol', row:5},   // Checking
  {stage:'B8',  sheet:'Balances', kind:'monthCol', row:6},   // Savings
  {stage:'B9',  sheet:'Balances', kind:'monthCol', row:7},   // Venmo float
  {stage:'B10', sheet:'Balances', kind:'monthCol', row:8},   // Fidelity MMkt
  {stage:'B11', sheet:'Balances', kind:'monthCol', row:9},   // Schwab
  {stage:'B12', sheet:'Balances', kind:'monthCol', row:10},  // 529
  {stage:'B13', sheet:'Balances', kind:'monthCol', row:11},  // FSA Dep Care bal
  {stage:'B14', sheet:'Balances', kind:'monthCol', row:12},  // FSA Health Care bal
  {stage:'B15', sheet:'Balances', kind:'monthCol', row:13},  // Matt 401k
  {stage:'B16', sheet:'Balances', kind:'monthCol', row:14},  // Sue 403b
  {stage:'B17', sheet:'Balances', kind:'monthCol', row:15},  // Sue pension est
  {stage:'B18', sheet:'Balances', kind:'monthCol', row:16},  // Other

  // (2) Solar energy log -- months are ROWS (14..25), metrics are columns
  {stage:'B21', sheet:'Solar Reinvestment', kind:'monthRow', col:2}, // Produced
  {stage:'B22', sheet:'Solar Reinvestment', kind:'monthRow', col:3}, // Consumed
  {stage:'B23', sheet:'Solar Reinvestment', kind:'monthRow', col:4}, // Imported
  {stage:'B24', sheet:'Solar Reinvestment', kind:'monthRow', col:5}, // Exported

  // (3) Solar bill -- months are columns
  {stage:'B26', sheet:'Solar ROI', kind:'monthCol', row:9},  // baseline
  {stage:'B27', sheet:'Solar ROI', kind:'monthCol', row:10}, // actual bill

  // (4) Amazon split -- months are columns
  {stage:'B31', sheet:'Amazon Reconciliation', kind:'monthCol', row:10}, // Groceries
  {stage:'B32', sheet:'Amazon Reconciliation', kind:'monthCol', row:11}, // Kids
  {stage:'B33', sheet:'Amazon Reconciliation', kind:'monthCol', row:12}, // Personal Care
  {stage:'B34', sheet:'Amazon Reconciliation', kind:'monthCol', row:13}, // Pharmacy/Health
  {stage:'B35', sheet:'Amazon Reconciliation', kind:'monthCol', row:14}, // Home/Hardware
  {stage:'B36', sheet:'Amazon Reconciliation', kind:'monthCol', row:15}, // Subscriptions
  {stage:'B37', sheet:'Amazon Reconciliation', kind:'monthCol', row:16}, // Other/gifts

  // (5) Comp & Benefits -- YTD single cells (column C)
  {stage:'B41', sheet:'Comp & Benefits', kind:'fixed', row:6,  col:3},  // Matt gross YTD
  {stage:'B42', sheet:'Comp & Benefits', kind:'fixed', row:8,  col:3},  // 401k YTD
  {stage:'B43', sheet:'Comp & Benefits', kind:'fixed', row:9,  col:3},  // 401k match YTD
  {stage:'B44', sheet:'Comp & Benefits', kind:'fixed', row:11, col:3},  // FSA Dep Care YTD
  {stage:'B45', sheet:'Comp & Benefits', kind:'fixed', row:12, col:3},  // FSA Health YTD
  {stage:'B46', sheet:'Comp & Benefits', kind:'fixed', row:13, col:3},  // Tuition YTD
  {stage:'B47', sheet:'Comp & Benefits', kind:'fixed', row:14, col:3},  // Stakeholder YTD
  {stage:'B48', sheet:'Comp & Benefits', kind:'fixed', row:20, col:3},  // Sue 403b YTD
  {stage:'B49', sheet:'Comp & Benefits', kind:'fixed', row:26, col:3}   // 529 YTD
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Family Finances')
    .addItem('Post Monthly Entry', 'postMonthlyEntry')
    .addSeparator()
    .addItem('Create Monthly Form', 'createMonthlyForm')
    .addItem('Clear Monthly Entry inputs', 'clearMonthlyEntry')
    .addToMenu();
}

function postMonthlyEntry() {
  var ss = SpreadsheetApp.getActive();
  var entry = ss.getSheetByName(ENTRY_SHEET);
  var ui = SpreadsheetApp.getUi();
  if (!entry) { ui.alert('No "' + ENTRY_SHEET + '" tab found.'); return; }

  var month = String(entry.getRange('B4').getValue()).trim();
  var mi = MONTHS.indexOf(month);
  if (mi < 0) { ui.alert('Pick a valid month (Jan–Dec) in B4.'); return; }
  var monthCol = 2 + mi;        // Jan -> col B (2)
  var monthRow = 14 + mi;       // Solar log row for the month

  var writes = [];
  for (var i = 0; i < MAP.length; i++) {
    var m = MAP[i];
    var v = entry.getRange(m.stage).getValue();
    if (v === '' || v === null) continue;   // skip blanks
    var r, c;
    if (m.kind === 'monthCol')      { r = m.row;    c = monthCol; }
    else if (m.kind === 'monthRow') { r = monthRow; c = m.col; }
    else                            { r = m.row;    c = m.col; }   // fixed
    writes.push({sheet: m.sheet, row: r, col: c, value: v});
  }

  if (!writes.length) { ui.alert('Nothing to post — all input cells are blank.'); return; }

  var resp = ui.alert('Post ' + writes.length + ' value(s) for ' + month + '?',
                      'They will overwrite the matching cells on the destination tabs.',
                      ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) return;

  writes.forEach(function(w) {
    var sh = ss.getSheetByName(w.sheet);
    if (sh) sh.getRange(w.row, w.col).setValue(w.value);
  });

  entry.getRange('B5').setNote('Last posted: ' + month + ' on ' + new Date());
  ui.alert('Posted ' + writes.length + ' value(s) for ' + month + '.');
}

function clearMonthlyEntry() {
  var ss = SpreadsheetApp.getActive();
  var entry = ss.getSheetByName(ENTRY_SHEET);
  if (!entry) return;
  MAP.forEach(function(m) { entry.getRange(m.stage).clearContent(); });
  SpreadsheetApp.getUi().alert('Cleared the Monthly Entry input cells (kept the month selector).');
}

/* ------------------------------------------------------------------ *
 *  OPTIONAL: Google Form that feeds the Monthly Entry tab
 * ------------------------------------------------------------------ */

// Maps a form question title -> the staging cell it should fill.
var FORM_MAP = [
  ['Month (Jan–Dec)', 'B4'],
  ['Checking balance', 'B7'], ['Savings balance', 'B8'], ['Venmo float', 'B9'],
  ['Fidelity Money Market', 'B10'], ['Schwab investment', 'B11'], ['529 balance', 'B12'],
  ['FSA Dep Care balance', 'B13'], ['FSA Health Care balance', 'B14'],
  ['Matt 401k balance', 'B15'], ['Sue 403b balance', 'B16'], ['Sue pension (est.)', 'B17'],
  ['Other assets', 'B18'],
  ['Solar produced (kWh)', 'B21'], ['Solar consumed (kWh)', 'B22'],
  ['Grid imported (kWh)', 'B23'], ['Grid exported (kWh)', 'B24'],
  ['Baseline electric (would-have-paid $)', 'B26'], ['Actual electric bill ($)', 'B27'],
  ['Amazon → Groceries', 'B31'], ['Amazon → Kids/Activities', 'B32'],
  ['Amazon → Personal Care', 'B33'], ['Amazon → Pharmacy/Health', 'B34'],
  ['Amazon → Home/Hardware', 'B35'], ['Amazon → Subscriptions', 'B36'],
  ['Amazon → Other/gifts', 'B37'],
  ['Matt gross earnings YTD', 'B41'], ['Matt 401k YTD', 'B42'], ['Matt 401k match YTD', 'B43'],
  ['FSA Dep Care YTD', 'B44'], ['FSA Health Care YTD', 'B45'], ['Tuition reimbursement YTD', 'B46'],
  ['Stakeholder earnings YTD', 'B47'], ['Sue 403b YTD', 'B48'], ['529 contributions YTD', 'B49']
];

function createMonthlyForm() {
  var ss = SpreadsheetApp.getActive();
  var form = FormApp.create('Family Finances — Monthly Entry');
  form.setDescription('Fill in what you have for the month; blanks are fine. ' +
                      'After submitting, open the sheet and run 📊 Family Finances → Post Monthly Entry.');

  var monthItem = form.addListItem();
  monthItem.setTitle('Month (Jan–Dec)').setRequired(true);
  monthItem.setChoiceValues(MONTHS);

  form.addSectionHeaderItem().setTitle('Account balances (month-end)');
  ['Checking balance','Savings balance','Venmo float','Fidelity Money Market','Schwab investment',
   '529 balance','FSA Dep Care balance','FSA Health Care balance','Matt 401k balance',
   'Sue 403b balance','Sue pension (est.)','Other assets'].forEach(addNum_(form));

  form.addSectionHeaderItem().setTitle('Solar — energy (Enphase, kWh)');
  ['Solar produced (kWh)','Solar consumed (kWh)','Grid imported (kWh)','Grid exported (kWh)'].forEach(addNum_(form));

  form.addSectionHeaderItem().setTitle('Solar — electric bill');
  ['Baseline electric (would-have-paid $)','Actual electric bill ($)'].forEach(addNum_(form));

  form.addSectionHeaderItem().setTitle("Amazon — split this month's total");
  ['Amazon → Groceries','Amazon → Kids/Activities','Amazon → Personal Care','Amazon → Pharmacy/Health',
   'Amazon → Home/Hardware','Amazon → Subscriptions','Amazon → Other/gifts'].forEach(addNum_(form));

  form.addSectionHeaderItem().setTitle('Comp & Benefits (YTD from paystub)');
  ['Matt gross earnings YTD','Matt 401k YTD','Matt 401k match YTD','FSA Dep Care YTD',
   'FSA Health Care YTD','Tuition reimbursement YTD','Stakeholder earnings YTD',
   'Sue 403b YTD','529 contributions YTD'].forEach(addNum_(form));

  // Install a submit trigger that funnels answers into the Monthly Entry tab.
  ScriptApp.newTrigger('onMonthlyFormSubmit').forForm(form).onFormSubmit().create();

  ss.getSheetByName(ENTRY_SHEET).getRange('C5')
    .setNote('Form URL: ' + form.getPublishedUrl());
  SpreadsheetApp.getUi().alert('Form created. Editable/live URLs are in your Google Drive; the live URL ' +
                               'is also saved as a note on Monthly Entry!C5.');
}

// numeric-ish text question helper (Forms number validation is limited; text keeps it simple)
function addNum_(form) {
  return function(title) {
    form.addTextItem().setTitle(title);
  };
}

function onMonthlyFormSubmit(e) {
  var ss = SpreadsheetApp.getActive();
  var entry = ss.getSheetByName(ENTRY_SHEET);
  var named = e && e.namedValues ? e.namedValues : {};
  FORM_MAP.forEach(function(pair) {
    var title = pair[0], cell = pair[1];
    if (named[title] && named[title][0] !== '') {
      var val = named[title][0];
      var num = Number(val);
      entry.getRange(cell).setValue(isNaN(num) ? val : num);
    }
  });
  // Auto-distribute after a form submit:
  postMonthlyEntry();
}
