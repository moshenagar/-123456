/**
 * ============================================================================
 *  CASH FLOW & PROFITABILITY MANAGEMENT SYSTEM
 *  Real-time cash flow, tax-liability isolation & break-even engine for SMBs
 * ============================================================================
 *  HOW TO USE
 *  1. Open (or create) a Google Sheet -> Extensions -> Apps Script.
 *  2. Paste this entire file as Code.gs and Save.
 *  3. Reload the Sheet. A menu "מערכת תזרים מזומנים" will appear.
 *  4. Run "🔧 התקן / אפס את המערכת" (or run setupCashFlowSystem from the
 *     script editor). Authorize the script when prompted.
 *  5. Fill in the yellow input cells on tabs 1-4 and 6. Everything else
 *     (blue/gray cells) is calculated automatically.
 *  6. The script is fully idempotent — re-running setupCashFlowSystem()
 *     rebuilds every tab from scratch (structure + formulas + formatting)
 *     without touching data you already typed into INPUT columns, except
 *     when a full reset is intentionally requested by re-running it.
 *
 *  KEY ASSUMPTIONS (documented so you can tune them to your business):
 *   - "5_Tax_Engine" computes VAT / tax-advance / social-security liability
 *     based on activity dated in the CURRENT calendar month. It is meant to
 *     be reviewed/re-run around the beginning of each month.
 *   - The 90-day dashboard applies that same monthly liability on every
 *     15th that falls inside the 90-day window (a practical approximation;
 *     for multi-month precision, duplicate the Tax Engine per month).
 *   - All monetary inputs on tabs 2 & 3 are "Including VAT" (gross); the
 *     system nets out VAT internally wherever a "net" figure is required.
 * ============================================================================
 */

// ============================================================================
// CONFIGURATION
// ============================================================================
const CONFIG = {
  sheets: {
    banks: '1_Banks_Credit',
    receivables: '2_Receivables',
    payables: '3_Payables',
    fixed: '4_Fixed_Expenses',
    tax: '5_Tax_Engine',
    pricing: '6_Pricing_BreakEven',
    dashboard: '7_Master_Forecast_Dashboard'
  },
  rows: {
    banksData: 12,          // data rows 2..13
    receivablesData: 150,   // data rows 2..151
    payablesData: 150,      // data rows 2..151
    fixedData: 25,          // data rows 2..26
    pricingData: 25,        // data rows 5..29
    forecastStart: 19,      // first day of the 90-day table (after the monthly outlook block)
    forecastDays: 91        // today (+0) .. today+90  => "90+ days ahead"
  },
  terms: ['מיידי', 'שוטף+30', 'שוטף+60', 'שוטף+90'],
  stressOptions: ['תרחיש רגיל', 'תרחיש קיצון (Stress Test)'],
  expenseTypes: ['קבועה', 'משתנה'],
  // Reference bank of common Israeli SMB fixed/operating expense categories with a
  // best-guess default VAT applicability. These are general defaults only - actual
  // VAT treatment can depend on the supplier's status (עוסק מורשה/פטור) and should
  // be confirmed with the business's accountant.
  expenseVatBank: [
    { name: 'שכר עבודה', vat: 'לא', note: 'משכורות אינן עסקה החייבת במע"מ (יחסי עובד-מעסיק)' },
    { name: 'ביטוח לאומי מעסיק', vat: 'לא', note: 'תשלום לביטוח לאומי אינו חייב במע"מ' },
    { name: 'ארנונה עסקית', vat: 'לא', note: 'ארנונה לרשות מקומית פטורה ממע"מ' },
    { name: 'מים וביוב (רשות מקומית)', vat: 'לא', note: 'פטור ממע"מ, כמו ארנונה' },
    { name: 'דלק לרכב (לא מגולם בתלוש)', vat: 'כן', note: 'רכישת דלק חייבת במע"מ מלא' },
    { name: 'טיפולים ותיקוני רכב / מוסך', vat: 'כן', note: 'שירות החייב במע"מ' },
    { name: 'ביטוח רכב', vat: 'לא', note: 'שירותי ביטוח פטורים ממע"מ' },
    { name: 'ביטוח עסק / אחריות מקצועית', vat: 'לא', note: 'שירותי ביטוח פטורים ממע"מ' },
    { name: 'שכירות נכס עסקי', vat: 'כן', note: 'בד"כ חייב אם המשכיר עוסק מורשה - לוודא' },
    { name: 'חשמל', vat: 'כן', note: 'חייב במע"מ מלא' },
    { name: 'טלפון ואינטרנט', vat: 'כן', note: 'חייב במע"מ מלא' },
    { name: 'הנהלת חשבונות / רואה חשבון', vat: 'כן', note: 'שירות מקצועי החייב במע"מ' },
    { name: 'עורך דין', vat: 'כן', note: 'שירות מקצועי החייב במע"מ' },
    { name: 'דמי ניהול / ועד בית / קניון', vat: 'כן', note: 'חייב במע"מ' },
    { name: 'מנוי תוכנה / SaaS', vat: 'כן', note: 'חייב במע"מ אם מספק ישראלי' },
    { name: 'אחסון ושרתים (Hosting)', vat: 'כן', note: 'חייב במע"מ אם מספק ישראלי' },
    { name: 'שיווק ופרסום', vat: 'כן', note: 'חייב במע"מ' },
    { name: 'דמי חבר בלשכה / ארגון מקצועי', vat: 'לא', note: 'תלוי בסטטוס הארגון - לרוב פטור, יש לוודא' },
    { name: 'עמלות בנק', vat: 'לא', note: 'שירותים פיננסיים פטורים ממע"מ' },
    { name: 'ריבית הלוואות', vat: 'לא', note: 'שירותים פיננסיים פטורים ממע"מ' },
    { name: 'ליסינג תפעולי (רכב/ציוד)', vat: 'כן', note: 'חייב במע"מ' },
    { name: 'אחר (הקלדה חופשית)', vat: '', note: 'קטגוריה מותאמת אישית - יש לקבוע ידנית האם כוללת מע"מ' }
  ],
  colors: {
    headerBg: '#1F3864',
    headerFont: '#FFFFFF',
    sectionBg: '#2E5395',
    sectionFont: '#FFFFFF',
    inputBg: '#FFF9DB',
    formulaBg: '#F3F3F3',
    totalBg: '#D9E1F2',
    goodBg: '#D9EAD3',
    goodFont: '#274E13',
    warnBg: '#FFF2CC',
    warnFont: '#7F6000',
    badBg: '#F4CCCC',
    badFont: '#7A1E1E'
  },
  fmt: {
    currency: '#,##0 ₪;[red]-#,##0 ₪',
    percent: '0.0%',
    date: 'dd/mm/yyyy'
  }
};

// ============================================================================
// MENU / ENTRY POINTS
// ============================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 מערכת תזרים מזומנים')
    .addItem('🔧 התקן / אפס את המערכת', 'setupCashFlowSystem')
    .addItem('🔄 רענן חישובים', 'refreshCalculations')
    .addToUi();
}

function refreshCalculations() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast('החישובים רועננו בהצלחה', 'מערכת תזרים מזומנים', 4);
}

/**
 * Main entry point. Builds (or rebuilds) all 7 tabs end-to-end.
 */
function setupCashFlowSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('בונה את הטאבים...', 'מערכת תזרים מזומנים', 5);

    buildBanksCreditTab(ss);
    buildReceivablesTab(ss);
    buildPayablesTab(ss);
    buildFixedExpensesTab(ss);
    buildTaxEngineTab(ss);
    buildPricingTab(ss);
    buildDashboardTab(ss);

    reorderSheets(ss);
    removeDefaultBlankSheet(ss);

    SpreadsheetApp.flush();
    ss.setActiveSheet(ss.getSheetByName(CONFIG.sheets.dashboard));

    ui.alert(
      'המערכת הותקנה בהצלחה ✅',
      'כל 7 הטאבים נבנו ועודכנו.\n\n' +
        '1. מלאו את התאים הצהובים (קלט) בטאבים 1-4 ו-6.\n' +
        '2. טאב 5 (מנוע מיסים) וטאב 7 (לוח בקרה) יתעדכנו אוטומטית.\n' +
        '3. ניתן להריץ את התקנה מחדש בכל עת - היא תבנה הכל מחדש בבטחה.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('שגיאה בהתקנת המערכת', String(err), ui.ButtonSet.OK);
    throw err;
  }
}

function reorderSheets(ss) {
  const order = [
    CONFIG.sheets.banks,
    CONFIG.sheets.receivables,
    CONFIG.sheets.payables,
    CONFIG.sheets.fixed,
    CONFIG.sheets.tax,
    CONFIG.sheets.pricing,
    CONFIG.sheets.dashboard
  ];
  order.forEach((name, idx) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) ss.setActiveSheet(sheet) && ss.moveActiveSheet(idx + 1);
  });
}

function removeDefaultBlankSheet(ss) {
  ['Sheet1', 'גיליון1'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow === 0 && lastCol === 0) ss.deleteSheet(sheet);
    }
  });
}

// ============================================================================
// SHARED HELPERS
// ============================================================================
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  resetSheet(sheet);
  return sheet;
}

function resetSheet(sheet) {
  sheet.clear();
  sheet.clearConditionalFormatRules();
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxRows, maxCols).clearDataValidations().clearNote();
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
    if (p.canEdit()) p.remove();
  });
  const sheetProtection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  sheetProtection.forEach(p => {
    if (p.canEdit()) p.remove();
  });
  sheet.setRightToLeft(true);
  sheet.setHiddenGridlines(false);
}

function styleHeaderRow(sheet, row, numCols) {
  const range = sheet.getRange(row, 1, 1, numCols);
  range
    .setBackground(CONFIG.colors.headerBg)
    .setFontColor(CONFIG.colors.headerFont)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(row, 34);
}

function styleSectionRow(sheet, a1Range, text) {
  const range = sheet.getRange(a1Range);
  range.merge();
  range
    .setValue(text)
    .setBackground(CONFIG.colors.sectionBg)
    .setFontColor(CONFIG.colors.sectionFont)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function styleTitleRow(sheet, a1Range, text) {
  const range = sheet.getRange(a1Range);
  range.merge();
  range
    .setValue(text)
    .setBackground('#0B1F3A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(range.getRow(), 36);
}

function setCurrency(range) {
  range.setNumberFormat(CONFIG.fmt.currency);
}
function setPercent(range) {
  range.setNumberFormat(CONFIG.fmt.percent);
}
function setDateFmt(range) {
  range.setNumberFormat(CONFIG.fmt.date);
}

function markInput(range) {
  range.setBackground(CONFIG.colors.inputBg);
}
function markFormula(range) {
  range.setBackground(CONFIG.colors.formulaBg).setFontColor('#434343');
}

function protectFormula(range, description) {
  const protection = range.protect().setDescription(description);
  protection.setWarningOnly(true);
}

function listValidation(values) {
  return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
}

function numberRangeValidation(min, max) {
  return SpreadsheetApp.newDataValidation().requireNumberBetween(min, max).setAllowInvalid(true).build();
}

function dateValidation() {
  return SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('יש להזין תאריך תקין (למשל 31/08/2026), לא טקסט חופשי.')
    .build();
}

/** Builds a column of row-specific formulas as a 2D array ready for setFormulas(). */
function colFormulas(startRow, endRow, template) {
  const out = [];
  for (let r = startRow; r <= endRow; r++) out.push([template(r)]);
  return out;
}

function fillDefault(sheet, a1Range, value) {
  sheet.getRange(a1Range).setValue(value);
}

function fillDefaultColumn(sheet, startRow, endRow, col, value) {
  const numRows = endRow - startRow + 1;
  const values = Array.from({ length: numRows }, () => [value]);
  sheet.getRange(startRow, col, numRows, 1).setValues(values);
}

// ============================================================================
// TAB 1: 1_Banks_Credit
// ============================================================================
function buildBanksCreditTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.banks);
  sheet.setTabColor('#1F3864');

  const headers = [
    'שם חשבון / כרטיס אשראי',
    'יתרה נוכחית (₪)',
    'מסגרת אשראי מאושרת (₪)',
    'יום חיוב חודשי (לכרטיסים)',
    'אחוז כרית ביטחון לחירום',
    'נזילות זמינה מותאמת סיכון (₪)'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, 1, headers.length);
  sheet.setFrozenRows(1);

  const first = 2;
  const last = 1 + CONFIG.rows.banksData;

  markInput(sheet.getRange(first, 1, CONFIG.rows.banksData, 5));
  fillDefaultColumn(sheet, first, last, 5, 0.05);

  const liquidityFormulas = colFormulas(first, last, r => `=IF($B${r}="","",$B${r}+$C${r}*(1-$E${r}))`);
  sheet.getRange(first, 6, liquidityFormulas.length, 1).setFormulas(liquidityFormulas);
  markFormula(sheet.getRange(first, 6, CONFIG.rows.banksData, 1));

  setCurrency(sheet.getRange(first, 2, CONFIG.rows.banksData, 2));
  setCurrency(sheet.getRange(first, 6, CONFIG.rows.banksData, 1));
  setPercent(sheet.getRange(first, 5, CONFIG.rows.banksData, 1));
  sheet.getRange(first, 4, CONFIG.rows.banksData, 1).setDataValidation(numberRangeValidation(1, 31));

  const totalRow = last + 2;
  sheet.getRange(totalRow, 1).setValue('סה"כ').setFontWeight('bold');
  sheet.getRange(totalRow, 2).setFormula(`=SUM(B${first}:B${last})`);
  sheet.getRange(totalRow, 3).setFormula(`=SUM(C${first}:C${last})`);
  sheet.getRange(totalRow, 6).setFormula(`=SUM(F${first}:F${last})`);
  const totalRange = sheet.getRange(totalRow, 1, 1, 6);
  totalRange.setBackground(CONFIG.colors.totalBg).setFontWeight('bold');
  setCurrency(sheet.getRange(totalRow, 2, 1, 1));
  setCurrency(sheet.getRange(totalRow, 3, 1, 1));
  setCurrency(sheet.getRange(totalRow, 6, 1, 1));
  protectFormula(sheet.getRange(totalRow, 2, 1, 3), 'סה"כ - שדה מחושב');
  protectFormula(sheet.getRange(totalRow, 6, 1, 1), 'סה"כ - שדה מחושב');
  protectFormula(sheet.getRange(first, 6, CONFIG.rows.banksData, 1), 'נזילות מותאמת - שדה מחושב');

  sheet.setColumnWidths(1, 1, 220);
  sheet.setColumnWidths(2, 3, 150);
  sheet.setColumnWidths(4, 1, 130);
  sheet.setColumnWidths(5, 2, 160);

  sheet.getRange(first, 1, CONFIG.rows.banksData, headers.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Named "total balance" reference used across the whole system:
  // '1_Banks_Credit'!B{totalRow}
  sheet.getRange('A' + totalRow).setNote('שורת סיכום - יתרת הבנק הכוללת ("' + `B${totalRow}` + '") משמשת כבסיס לתחזית בטאב 7 ולטאב 5.');
}

// ============================================================================
// TAB 2: 2_Receivables
// ============================================================================
function buildReceivablesTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.receivables);
  sheet.setTabColor('#38761D');

  const headers = [
    'שם לקוח',
    'מספר חשבונית',
    'סכום חשבונית כולל מע"מ (₪)',
    'תאריך חשבונית',
    'תנאי אשראי',
    'תאריך פירעון תיאורטי',
    'חיץ פיגור (ימים נוספים)',
    'אחוז הסתברות גבייה',
    'תאריך תשלום צפוי מתואם',
    'מזומן צפוי נטו ממע"מ (₪)',
    'מזומן צפוי לתזרים כולל מע"מ (₪)',
    'סטטוס גבייה'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, 1, headers.length);
  sheet.setFrozenRows(1);

  const first = 2;
  const last = 1 + CONFIG.rows.receivablesData;
  const n = CONFIG.rows.receivablesData;

  markInput(sheet.getRange(first, 1, n, 5));
  markInput(sheet.getRange(first, 7, n, 2));
  markInput(sheet.getRange(first, 12, n, 1));

  sheet.getRange(first, 4, n, 1).setDataValidation(dateValidation());
  sheet.getRange(first, 5, n, 1).setDataValidation(listValidation(CONFIG.terms));
  sheet.getRange(first, 12, n, 1).setDataValidation(listValidation(['ממתין', 'התקבל', 'נסגר']));
  fillDefaultColumn(sheet, first, last, 7, 15);
  fillDefaultColumn(sheet, first, last, 8, 0.8);
  fillDefaultColumn(sheet, first, last, 12, 'ממתין');

  const dueDateFormulas = colFormulas(
    first,
    last,
    r =>
      `=IF($D${r}="","",IF($E${r}="מיידי",$D${r},IF($E${r}="שוטף+30",$D${r}+30,IF($E${r}="שוטף+60",$D${r}+60,IF($E${r}="שוטף+90",$D${r}+90,$D${r})))))`
  );
  sheet.getRange(first, 6, n, 1).setFormulas(dueDateFormulas);

  const adjDateFormulas = colFormulas(first, last, r => `=IF($F${r}="","",$F${r}+$G${r})`);
  sheet.getRange(first, 9, n, 1).setFormulas(adjDateFormulas);

  const netCashFormulas = colFormulas(
    first,
    last,
    r => `=IF($C${r}="","",($C${r}/(1+'${CONFIG.sheets.tax}'!$B$4))*$H${r})`
  );
  sheet.getRange(first, 10, n, 1).setFormulas(netCashFormulas);

  const grossCashFormulas = colFormulas(first, last, r => `=IF($C${r}="","",$C${r}*$H${r})`);
  sheet.getRange(first, 11, n, 1).setFormulas(grossCashFormulas);

  markFormula(sheet.getRange(first, 6, n, 1));
  markFormula(sheet.getRange(first, 9, n, 3));
  protectFormula(sheet.getRange(first, 6, n, 1), 'תאריך פירעון תיאורטי - שדה מחושב');
  protectFormula(sheet.getRange(first, 9, n, 3), 'שדות מחושבים');

  setCurrency(sheet.getRange(first, 3, n, 1));
  setDateFmt(sheet.getRange(first, 4, n, 1));
  setDateFmt(sheet.getRange(first, 6, n, 1));
  setDateFmt(sheet.getRange(first, 9, n, 1));
  setPercent(sheet.getRange(first, 8, n, 1));
  setCurrency(sheet.getRange(first, 10, n, 2));

  sheet.setColumnWidths(1, 1, 170);
  sheet.setColumnWidths(2, 1, 110);
  sheet.setColumnWidths(3, 1, 170);
  sheet.setColumnWidths(4, 1, 120);
  sheet.setColumnWidths(5, 1, 110);
  sheet.setColumnWidths(6, 1, 140);
  sheet.setColumnWidths(7, 1, 130);
  sheet.setColumnWidths(8, 1, 130);
  sheet.setColumnWidths(9, 1, 150);
  sheet.setColumnWidths(10, 2, 190);
  sheet.setColumnWidths(12, 1, 120);

  applyStatusConditionalFormatting(sheet, `L${first}:L${last}`, {
    'התקבל': CONFIG.colors.goodBg,
    'נסגר': CONFIG.colors.goodBg,
    'ממתין': CONFIG.colors.warnBg
  });

  sheet.getRange(first, 1, n, headers.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

// ============================================================================
// TAB 3: 3_Payables
// ============================================================================
function buildPayablesTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.payables);
  sheet.setTabColor('#B45309');

  const headers = [
    'שם ספק',
    'תיאור רכישה',
    'סכום כולל מע"מ (₪)',
    'תאריך חשבונית/הזמנה',
    'תנאי תשלום',
    'תאריך פירעון צפוי',
    'אמצעי תשלום',
    'אחוז חיץ לשינוי מחיר בלתי צפוי',
    'סכום פירעון צפוי מותאם (₪)',
    'סטטוס'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, 1, headers.length);
  sheet.setFrozenRows(1);

  const first = 2;
  const last = 1 + CONFIG.rows.payablesData;
  const n = CONFIG.rows.payablesData;

  markInput(sheet.getRange(first, 1, n, 4));
  markInput(sheet.getRange(first, 5, n, 1));
  markInput(sheet.getRange(first, 7, n, 2));
  markInput(sheet.getRange(first, 10, n, 1));

  sheet.getRange(first, 4, n, 1).setDataValidation(dateValidation());
  sheet.getRange(first, 5, n, 1).setDataValidation(listValidation(CONFIG.terms));
  sheet.getRange(first, 7, n, 1).setDataValidation(listValidation(['צ׳ק', 'העברה בנקאית', 'הוראת קבע', 'אשראי', 'מזומן']));
  sheet.getRange(first, 10, n, 1).setDataValidation(listValidation(['ממתין', 'נפרע']));
  fillDefaultColumn(sheet, first, last, 8, 0.03);
  fillDefaultColumn(sheet, first, last, 10, 'ממתין');

  const dueDateFormulas = colFormulas(
    first,
    last,
    r =>
      `=IF($D${r}="","",IF($E${r}="מיידי",$D${r},IF($E${r}="שוטף+30",$D${r}+30,IF($E${r}="שוטף+60",$D${r}+60,IF($E${r}="שוטף+90",$D${r}+90,$D${r})))))`
  );
  sheet.getRange(first, 6, n, 1).setFormulas(dueDateFormulas);

  const adjAmountFormulas = colFormulas(first, last, r => `=IF($C${r}="","",$C${r}*(1+$H${r}))`);
  sheet.getRange(first, 9, n, 1).setFormulas(adjAmountFormulas);

  markFormula(sheet.getRange(first, 6, n, 1));
  markFormula(sheet.getRange(first, 9, n, 1));
  protectFormula(sheet.getRange(first, 6, n, 1), 'תאריך פירעון צפוי - שדה מחושב');
  protectFormula(sheet.getRange(first, 9, n, 1), 'סכום מותאם - שדה מחושב');

  setCurrency(sheet.getRange(first, 3, n, 1));
  setDateFmt(sheet.getRange(first, 4, n, 1));
  setDateFmt(sheet.getRange(first, 6, n, 1));
  setPercent(sheet.getRange(first, 8, n, 1));
  setCurrency(sheet.getRange(first, 9, n, 1));

  sheet.setColumnWidths(1, 1, 170);
  sheet.setColumnWidths(2, 1, 200);
  sheet.setColumnWidths(3, 1, 150);
  sheet.setColumnWidths(4, 1, 140);
  sheet.setColumnWidths(5, 1, 110);
  sheet.setColumnWidths(6, 1, 140);
  sheet.setColumnWidths(7, 1, 150);
  sheet.setColumnWidths(8, 1, 160);
  sheet.setColumnWidths(9, 1, 180);
  sheet.setColumnWidths(10, 1, 110);

  applyStatusConditionalFormatting(sheet, `J${first}:J${last}`, {
    'נפרע': CONFIG.colors.goodBg,
    'ממתין': CONFIG.colors.warnBg
  });

  sheet.getRange(first, 1, n, headers.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

// ============================================================================
// TAB 4: 4_Fixed_Expenses
// ============================================================================
function buildFixedExpensesTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.fixed);
  sheet.setTabColor('#741B47');

  const headers = [
    'קטגוריית הוצאה',
    'סוג (קבועה / משתנה)',
    'יום חיוב בחודש (לקבועות - חוזר כל חודש)',
    'תאריך מדויק (למשתנות / תשלום חד-פעמי)',
    'סכום בסיס (₪)',
    'מכפיל דינמי / כרית הגנה (למשל 1.25 לעונתיות)',
    'סה"כ תזרים צפוי (₪)',
    'כולל מע"מ?'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, 1, headers.length);
  sheet.setFrozenRows(1);

  const first = 2;
  const last = 1 + CONFIG.rows.fixedData;
  const n = CONFIG.rows.fixedData;

  const categoryNames = CONFIG.expenseVatBank.map(item => item.name);
  markInput(sheet.getRange(first, 1, n, 6));
  markInput(sheet.getRange(first, 8, n, 1));

  sheet.getRange(first, 1, n, 1).setDataValidation(listValidation(categoryNames));
  sheet.getRange(first, 2, n, 1).setDataValidation(listValidation(CONFIG.expenseTypes));
  sheet.getRange(first, 3, n, 1).setDataValidation(numberRangeValidation(1, 31));
  sheet.getRange(first, 4, n, 1).setDataValidation(dateValidation());
  sheet.getRange(first, 8, n, 1).setDataValidation(listValidation(['כן', 'לא']));
  fillDefaultColumn(sheet, first, last, 2, 'קבועה');
  fillDefaultColumn(sheet, first, last, 6, 1);
  fillDefaultColumn(sheet, first, last, 8, 'כן');

  const outflowFormulas = colFormulas(first, last, r => `=IF($E${r}="","",$E${r}*$F${r})`);
  sheet.getRange(first, 7, n, 1).setFormulas(outflowFormulas);
  markFormula(sheet.getRange(first, 7, n, 1));
  protectFormula(sheet.getRange(first, 7, n, 1), 'תזרים - שדה מחושב');

  setCurrency(sheet.getRange(first, 5, n, 1));
  setDateFmt(sheet.getRange(first, 4, n, 1));

  const totalRow = last + 2;
  sheet.getRange(totalRow, 1).setValue('סה"כ (קבועות + משתנות בטבלה)').setFontWeight('bold');
  sheet.getRange(totalRow, 7).setFormula(`=SUM(G${first}:G${last})`);
  sheet.getRange(totalRow, 1, 1, 8).setBackground(CONFIG.colors.totalBg).setFontWeight('bold');
  setCurrency(sheet.getRange(totalRow, 7, 1, 1));
  protectFormula(sheet.getRange(totalRow, 7, 1, 1), 'סה"כ - שדה מחושב');
  sheet
    .getRange('A' + totalRow)
    .setNote(
      'תקורה חודשית קבועה בלבד (ללא משתנות) מוזנת אוטומטית לטאב 6 מתוך שורות "קבועה" בלבד. הוצאות "משתנות" נספרות רק בחודש התאריך המדויק שלהן בתחזית טאב 7 - לתשלומים מפוצלים, הוסיפו שורה נפרדת לכל תשלום עם תאריך משלו.'
    );

  sheet.setColumnWidths(1, 1, 210);
  sheet.setColumnWidths(2, 1, 140);
  sheet.setColumnWidths(3, 1, 190);
  sheet.setColumnWidths(4, 1, 190);
  sheet.setColumnWidths(5, 1, 140);
  sheet.setColumnWidths(6, 1, 220);
  sheet.setColumnWidths(7, 1, 170);
  sheet.setColumnWidths(8, 1, 110);

  sheet.getRange(first, 1, n, headers.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  buildExpenseVatBankReference(sheet);
}

/**
 * Writes the visible, editable reference table of common expense categories with
 * their default VAT applicability, used both as the source list for the category
 * dropdown (column A) and, via onEdit(), to auto-fill the "כולל מע"מ?" column.
 */
function buildExpenseVatBankReference(sheet) {
  const startCol = 10; // column J, leaving column I as a spacer
  const bank = CONFIG.expenseVatBank;

  const titleRange = sheet.getRange(1, startCol, 1, 3);
  titleRange.merge();
  titleRange
    .setValue('בנק קטגוריות הוצאה - ברירת מחדל למע"מ (לבחירה מהירה בעמודה "קטגוריית הוצאה")')
    .setBackground(CONFIG.colors.sectionBg)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);

  const bankHeaders = ['קטגוריה', 'מע"מ כברירת מחדל', 'הערה'];
  sheet.getRange(2, startCol, 1, 3).setValues([bankHeaders]);
  sheet
    .getRange(2, startCol, 1, 3)
    .setBackground(CONFIG.colors.headerBg)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const rows = bank.map(item => [item.name, item.vat, item.note]);
  sheet.getRange(3, startCol, rows.length, 3).setValues(rows);
  sheet.getRange(3, startCol, rows.length, 3).setBackground('#F8F8F8').setWrap(true);
  sheet.getRange(3, startCol + 1, rows.length, 1).setHorizontalAlignment('center');

  const noteRow = 3 + rows.length + 1;
  const noteRange = sheet.getRange(noteRow, startCol, 1, 3);
  noteRange.merge();
  noteRange
    .setValue('הערה: ברירות המחדל הן הנחיה כללית בלבד - יש לוודא מול רואה החשבון בהתאם לסטטוס הספק (עוסק מורשה/פטור).')
    .setFontStyle('italic')
    .setFontSize(9)
    .setWrap(true);

  sheet.setColumnWidths(startCol, 1, 220);
  sheet.setColumnWidths(startCol + 1, 1, 140);
  sheet.setColumnWidths(startCol + 2, 1, 340);
}

/**
 * Simple onEdit trigger: when the expense category (column A) on 4_Fixed_Expenses
 * is set to a value matching the VAT bank, auto-fill the "כולל מע"מ?" column (H)
 * with the bank's default. The user can still manually overwrite H afterward -
 * this only re-fires when column A itself is edited again.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.sheets.fixed) return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    const first = 2;
    const last = 1 + CONFIG.rows.fixedData;
    if (row < first || row > last || col !== 1) return;

    const category = String(e.range.getValue()).trim();
    if (!category) return;

    const match = CONFIG.expenseVatBank.find(item => item.name === category);
    if (match && match.vat) {
      sheet.getRange(row, 8).setValue(match.vat);
    }
  } catch (err) {
    // Never block manual editing due to an auto-fill failure.
  }
}

// ============================================================================
// TAB 5: 5_Tax_Engine
// ============================================================================
function buildTaxEngineTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.tax);
  sheet.setTabColor('#990000');

  styleTitleRow(sheet, 'A1:C1', 'מנוע מיסים - מע"מ, מקדמות מס הכנסה וביטוח לאומי');
  sheet.setFrozenRows(1);

  styleSectionRow(sheet, 'A3:C3', 'פרמטרים (קלט)');
  const params = [
    ['שיעור מע"מ', 0.17],
    ['אחוז מקדמת מס הכנסה מהמחזור', 0.07],
    ['ביטוח לאומי - סכום חודשי קבוע (₪)', 0],
    ['אחוז כרית ביטחון למיסים (Safety Buffer)', 0.025]
  ];
  sheet.getRange(4, 1, params.length, 2).setValues(params);
  markInput(sheet.getRange(4, 2, params.length, 1));
  setPercent(sheet.getRange(4, 2, 1, 1));
  setPercent(sheet.getRange(5, 2, 1, 1));
  setCurrency(sheet.getRange(6, 2, 1, 1));
  setPercent(sheet.getRange(7, 2, 1, 1));

  const receivablesLast = 1 + CONFIG.rows.receivablesData;
  const payablesLast = 1 + CONFIG.rows.payablesData;
  const fixedLast = 1 + CONFIG.rows.fixedData;

  styleSectionRow(sheet, 'A9:C9', 'חישובי מע"מ ומיסים (חודש נוכחי)');
  const calcLabels = [
    ['מע"מ עסקאות (פלט) - מכירות החודש', ''],
    ['מע"מ תשומות (קלט) - ספקים החודש', ''],
    ['מע"מ תשומות (קלט) - הוצאות קבועות', ''],
    ['מע"מ נטו לתשלום ב-15 לחודש', ''],
    ['מחזור עסקאות ללא מע"מ - החודש', ''],
    ['מקדמת מס הכנסה', ''],
    ['ביטוח לאומי', ''],
    ['סה"כ התחייבות מיסים לפני כרית ביטחון', ''],
    ['כרית ביטחון למיסים', ''],
    ['סה"כ להפרשה ב-15 לחודש (כולל כרית)', '']
  ];
  sheet.getRange(10, 1, calcLabels.length, 2).setValues(calcLabels);

  sheet
    .getRange('B10')
    .setFormula(
      `=SUMPRODUCT((YEAR('${CONFIG.sheets.receivables}'!$D$2:$D$${receivablesLast})=YEAR(TODAY()))*(MONTH('${CONFIG.sheets.receivables}'!$D$2:$D$${receivablesLast})=MONTH(TODAY()))*'${CONFIG.sheets.receivables}'!$C$2:$C$${receivablesLast})/(1+$B$4)*$B$4`
    );
  sheet
    .getRange('B11')
    .setFormula(
      `=SUMPRODUCT((YEAR('${CONFIG.sheets.payables}'!$D$2:$D$${payablesLast})=YEAR(TODAY()))*(MONTH('${CONFIG.sheets.payables}'!$D$2:$D$${payablesLast})=MONTH(TODAY()))*'${CONFIG.sheets.payables}'!$C$2:$C$${payablesLast})/(1+$B$4)*$B$4`
    );
  sheet
    .getRange('B12')
    .setFormula(`=SUMIF('${CONFIG.sheets.fixed}'!$H$2:$H$${fixedLast},"כן",'${CONFIG.sheets.fixed}'!$G$2:$G$${fixedLast})/(1+$B$4)*$B$4`);
  sheet.getRange('B13').setFormula('=MAX(0,B10-B11-B12)');
  sheet
    .getRange('B14')
    .setFormula(
      `=SUMPRODUCT((YEAR('${CONFIG.sheets.receivables}'!$D$2:$D$${receivablesLast})=YEAR(TODAY()))*(MONTH('${CONFIG.sheets.receivables}'!$D$2:$D$${receivablesLast})=MONTH(TODAY()))*'${CONFIG.sheets.receivables}'!$C$2:$C$${receivablesLast})/(1+$B$4)`
    );
  sheet.getRange('B15').setFormula('=B14*B5');
  sheet.getRange('B16').setFormula('=B6');
  sheet.getRange('B17').setFormula('=B13+B15+B16');
  sheet.getRange('B18').setFormula('=B17*B7');
  sheet.getRange('B19').setFormula('=B17+B18');

  setCurrency(sheet.getRange('B10:B19'));
  markFormula(sheet.getRange('B10:B19'));
  protectFormula(sheet.getRange('B10:B19'), 'חישובי מס - שדות מחושבים');
  sheet.getRange('A19').setNote('סכום זה ("B19") מנוכה אוטומטית בטאב 7 בכל 15 לחודש שנופל בטווח 90 הימים.');

  styleSectionRow(sheet, 'A21:C21', 'סיכום - יתרה חופשית ממיסים');
  sheet.getRange('A22').setValue('סה"כ יתרת בנק נוכחית');
  sheet.getRange('B22').setFormula(`='${CONFIG.sheets.banks}'!B${1 + CONFIG.rows.banksData + 2}`);
  sheet.getRange('A23').setValue('יתרה חופשית ממיסים (True Free Cash Balance)');
  sheet.getRange('B23').setFormula('=B22-B19');
  setCurrency(sheet.getRange('B22:B23'));
  sheet.getRange('B22:B23').setFontWeight('bold').setFontSize(12);
  markFormula(sheet.getRange('B22:B23'));
  protectFormula(sheet.getRange('B22:B23'), 'סיכום - שדות מחושבים');

  applyNumericConditionalFormatting(sheet, 'B23:B23', 0);

  sheet.setColumnWidths(1, 1, 380);
  sheet.setColumnWidths(2, 1, 170);
  sheet.getRange('A4:B23').setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

// ============================================================================
// TAB 6: 6_Pricing_BreakEven
// ============================================================================
function buildPricingTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.pricing);
  sheet.setTabColor('#0B5394');

  styleTitleRow(sheet, 'A1:R1', 'תמחור ונקודת איזון - כלכלת יחידה (Unit Economics)');
  sheet.getRange('A2').setValue('סה"כ תקורה חודשית קבועה (מטאב 4, שורות "קבועה" בלבד)');
  const fixedLastRow = 1 + CONFIG.rows.fixedData;
  sheet
    .getRange('B2')
    .setFormula(`=SUMIF('${CONFIG.sheets.fixed}'!$B$2:$B$${fixedLastRow},"קבועה",'${CONFIG.sheets.fixed}'!$G$2:$G$${fixedLastRow})`);
  setCurrency(sheet.getRange('B2'));
  sheet.getRange('B2').setFontWeight('bold').setFontSize(12);
  markFormula(sheet.getRange('A2:B2'));
  protectFormula(sheet.getRange('B2'), 'תקורה חודשית - שדה מחושב מטאב 4');

  const headers = [
    'שם מוצר / שירות',
    'עלות ספק - COGS (₪, לפני מע"מ)',
    'מחיר מכירה מתוכנן (₪, לפני מע"מ)',
    'תנאי תשלום לספק',
    'תנאי תשלום מלקוח',
    'עמלת סליקת אשראי %',
    'אחוז חיץ כשל/גריעה/משלוח',
    'רווח גולמי תיאורטי (₪)',
    'רווח גולמי תיאורטי (%)',
    'רווח גולמי נטו מזומן (₪)',
    'רווח גולמי נטו מזומן (%)',
    'כמות מכירות חודשית צפויה (יח\')',
    'הקצאת תקורה ליחידה (₪)',
    'רווח תפעולי נטו ליחידה (₪)',
    'נקודת איזון - כמות חודשית (יח\')',
    'נקודת איזון - מחזור (₪)',
    'אחוז רווח נטו ליחידה',
    'סטטוס'
  ];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, 4, headers.length);
  sheet.setFrozenRows(4);

  const first = 5;
  const last = 4 + CONFIG.rows.pricingData;
  const n = CONFIG.rows.pricingData;

  markInput(sheet.getRange(first, 1, n, 3));
  markInput(sheet.getRange(first, 4, n, 4));
  markInput(sheet.getRange(first, 12, n, 1));

  sheet.getRange(first, 4, n, 1).setDataValidation(listValidation(CONFIG.terms));
  sheet.getRange(first, 5, n, 1).setDataValidation(listValidation(CONFIG.terms));
  fillDefaultColumn(sheet, first, last, 6, 0.018);
  fillDefaultColumn(sheet, first, last, 7, 0.02);

  const f = {
    H: r => `=IF($C${r}="","",$C${r}-$B${r})`,
    I: r => `=IF($C${r}="","",$H${r}/$C${r})`,
    J: r => `=IF($C${r}="","",$C${r}*(1-$F${r}-$G${r})-$B${r})`,
    K: r => `=IF($C${r}="","",$J${r}/$C${r})`,
    M: r => `=IF(OR($L${r}="",$L${r}=0),0,$B$2/$L${r})`,
    N: r => `=IF($C${r}="","",$J${r}-$M${r})`,
    O: r => `=IF($J${r}<=0,"אין איזון (הפסד)",$B$2/$J${r})`,
    P: r => `=IF(ISNUMBER($O${r}),$O${r}*$C${r},"—")`,
    Q: r => `=IF($C${r}="","",$N${r}/$C${r})`,
    R: r =>
      `=IF($C${r}="","",IF($Q${r}>0.2,"🟢 רווח בריא",IF($Q${r}>=0,"🟡 גבול איזון",IF($H${r}>0,"🔴 הפסד מוסווה","🔴 הפסד"))))`
  };
  const colIndex = { H: 8, I: 9, J: 10, K: 11, M: 13, N: 14, O: 15, P: 16, Q: 17, R: 18 };
  Object.keys(f).forEach(key => {
    const formulas = colFormulas(first, last, f[key]);
    sheet.getRange(first, colIndex[key], n, 1).setFormulas(formulas);
  });

  const formulaCols = [8, 9, 10, 11, 13, 14, 15, 16, 17, 18];
  formulaCols.forEach(c => markFormula(sheet.getRange(first, c, n, 1)));
  protectFormula(sheet.getRange(first, 8, n, 4), 'שדות מחושבים - תמחור ונקודת איזון');
  protectFormula(sheet.getRange(first, 13, n, 6), 'שדות מחושבים - תמחור ונקודת איזון');

  setCurrency(sheet.getRange(first, 2, n, 1));
  setCurrency(sheet.getRange(first, 3, n, 1));
  setCurrency(sheet.getRange(first, 8, n, 1));
  setPercent(sheet.getRange(first, 6, n, 2));
  setPercent(sheet.getRange(first, 9, n, 1));
  setCurrency(sheet.getRange(first, 10, n, 1));
  setPercent(sheet.getRange(first, 11, n, 1));
  setCurrency(sheet.getRange(first, 13, n, 2));
  setCurrency(sheet.getRange(first, 16, n, 1));
  setPercent(sheet.getRange(first, 17, n, 1));

  applyNumericConditionalFormattingBand(sheet, `Q${first}:Q${last}`, 0, 0.2);

  sheet.setColumnWidths(1, 1, 170);
  sheet.setColumnWidths(2, 2, 160);
  sheet.setColumnWidths(4, 2, 130);
  sheet.setColumnWidths(6, 2, 150);
  sheet.setColumnWidths(8, 10, 150);
  sheet.setColumnWidths(18, 1, 150);

  sheet.getRange(first, 1, n, headers.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

// ============================================================================
// TAB 7: 7_Master_Forecast_Dashboard
// ============================================================================
function buildDashboardTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.dashboard);
  sheet.setTabColor('#38761D');

  styleTitleRow(sheet, 'A1:I1', 'לוח בקרה ראשי - תזרים מזומנים ורווחיות (90 יום קדימה)');

  sheet.getRange('A3').setValue('מצב תרחיש:').setFontWeight('bold');
  sheet.getRange('B3').setValue(CONFIG.stressOptions[0]);
  sheet.getRange('B3').setDataValidation(listValidation(CONFIG.stressOptions));
  markInput(sheet.getRange('B3'));
  sheet.getRange('B3').setFontWeight('bold');

  sheet.getRange('D3').setValue('מכפיל הוצאות (Stress)');
  sheet.getRange('E3').setFormula(`=IF($B$3="${CONFIG.stressOptions[1]}",1.15,1)`);
  sheet.getRange('F3').setValue('מקדם גבייה (Stress)');
  sheet.getRange('G3').setFormula(`=IF($B$3="${CONFIG.stressOptions[1]}",0.85,1)`);
  setPercent(sheet.getRange('E3'));
  setPercent(sheet.getRange('G3'));
  markFormula(sheet.getRange('E3'));
  markFormula(sheet.getRange('G3'));
  protectFormula(sheet.getRange('E3:G3'), 'מכפילי תרחיש קיצון - שדות מחושבים');

  const banksTotalRow = 1 + CONFIG.rows.banksData + 2;
  const forecastStart = CONFIG.rows.forecastStart;
  const forecastEnd = forecastStart + CONFIG.rows.forecastDays - 1;
  const row15 = forecastStart + 15;
  const row30 = forecastStart + 30;
  const row60 = forecastStart + 60;
  const row90 = forecastStart + 90;

  buildKpiCard(sheet, 'A5:B5', 'A6:B6', 'יתרת בנק נוכחית', `='${CONFIG.sheets.banks}'!B${banksTotalRow}`, 'currency');
  buildKpiCard(sheet, 'C5:D5', 'C6:D6', 'יתרה חופשית ממיסים', `='${CONFIG.sheets.tax}'!B23`, 'currency');
  buildKpiCard(sheet, 'E5:F5', 'E6:F6', 'תחזית נזילות +15 יום', `=$H$${row15}`, 'currency');
  buildKpiCard(sheet, 'G5:H5', 'G6:H6', 'תחזית נזילות +30 יום', `=$H$${row30}`, 'currency');

  buildKpiCard(sheet, 'A8:B8', 'A9:B9', 'תחזית נזילות +60 יום', `=$H$${row60}`, 'currency');
  buildKpiCard(sheet, 'C8:D8', 'C9:D9', 'תחזית נזילות +90 יום', `=$H$${row90}`, 'currency');
  buildKpiCard(
    sheet,
    'E8:F8',
    'E9:F9',
    'נקודת שפל צפויה (תאריך | סכום)',
    `=TEXT(INDEX($A$${forecastStart}:$A$${forecastEnd},MATCH(MIN($H$${forecastStart}:$H$${forecastEnd}),$H$${forecastStart}:$H$${forecastEnd},0)),"dd/mm/yyyy")&" | "&TEXT(MIN($H$${forecastStart}:$H$${forecastEnd}),"#,##0")&" ₪"`,
    'text'
  );
  buildKpiCard(
    sheet,
    'G8:H8',
    'G9:H9',
    'סטטוס נקודת איזון - החודש',
    `=IF(SUMPRODUCT('${CONFIG.sheets.pricing}'!$L$5:$L$${4 + CONFIG.rows.pricingData},IFERROR('${CONFIG.sheets.pricing}'!$N$5:$N$${4 + CONFIG.rows.pricingData},0))>=0,"✅ מעל נקודת האיזון","⚠️ מתחת לנקודת האיזון")`,
    'text'
  );

  protectFormula(sheet.getRange('A6:H6'), 'כרטיסי מדדים - שדות מחושבים');
  protectFormula(sheet.getRange('A9:H9'), 'כרטיסי מדדים - שדות מחושבים');

  const fixedLast = 1 + CONFIG.rows.fixedData;

  buildMonthlyOutlook(sheet, forecastStart, forecastEnd, fixedLast);

  // ---- Daily rolling 90-day forecast table ----
  const headers = [
    'תאריך',
    'יום',
    'יתרת פתיחה (₪)',
    'תקבולים צפויים (₪)',
    'תשלומים לספקים (₪)',
    'הוצאות קבועות ומשתנות (₪)',
    'תשלום מס - 15 לחודש (₪)',
    'יתרת סגירה (₪)',
    'מצב'
  ];
  const dailyHeaderRow = forecastStart - 1;
  sheet.getRange(dailyHeaderRow, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, dailyHeaderRow, headers.length);
  sheet.setFrozenRows(dailyHeaderRow);

  const receivablesLast = 1 + CONFIG.rows.receivablesData;
  const payablesLast = 1 + CONFIG.rows.payablesData;

  const aFormulas = [];
  const bFormulas = [];
  const cFormulas = [];
  const dFormulas = [];
  const eFormulas = [];
  const fFormulas = [];
  const gFormulas = [];
  const hFormulas = [];
  const iFormulas = [];

  for (let r = forecastStart; r <= forecastEnd; r++) {
    aFormulas.push([r === forecastStart ? '=TODAY()' : `=$A${r - 1}+1`]);
    bFormulas.push([`=TEXT($A${r},"ddd")`]);
    cFormulas.push([r === forecastStart ? `='${CONFIG.sheets.banks}'!B${banksTotalRow}` : `=$H${r - 1}`]);
    dFormulas.push([
      `=SUMIFS('${CONFIG.sheets.receivables}'!$K$2:$K$${receivablesLast},'${CONFIG.sheets.receivables}'!$I$2:$I$${receivablesLast},$A${r},'${CONFIG.sheets.receivables}'!$L$2:$L$${receivablesLast},"ממתין")*$G$3`
    ]);
    eFormulas.push([
      `=SUMIFS('${CONFIG.sheets.payables}'!$I$2:$I$${payablesLast},'${CONFIG.sheets.payables}'!$F$2:$F$${payablesLast},$A${r},'${CONFIG.sheets.payables}'!$J$2:$J$${payablesLast},"ממתין")*$E$3`
    ]);
    fFormulas.push([
      `=(SUMIFS('${CONFIG.sheets.fixed}'!$G$2:$G$${fixedLast},'${CONFIG.sheets.fixed}'!$B$2:$B$${fixedLast},"קבועה",'${CONFIG.sheets.fixed}'!$C$2:$C$${fixedLast},DAY($A${r}))+SUMIFS('${CONFIG.sheets.fixed}'!$G$2:$G$${fixedLast},'${CONFIG.sheets.fixed}'!$B$2:$B$${fixedLast},"משתנה",'${CONFIG.sheets.fixed}'!$D$2:$D$${fixedLast},$A${r}))*$E$3`
    ]);
    gFormulas.push([`=IF(DAY($A${r})=15,'${CONFIG.sheets.tax}'!$B$19,0)`]);
    hFormulas.push([`=$C${r}+$D${r}-$E${r}-$F${r}-$G${r}`]);
    iFormulas.push([
      `=IF($H${r}<0,"🔴 קריטי",IF($H${r}<'${CONFIG.sheets.banks}'!$B$${banksTotalRow}*0.15,"🟡 אזהרה","🟢 תקין"))`
    ]);
  }

  const numRows = forecastEnd - forecastStart + 1;
  sheet.getRange(forecastStart, 1, numRows, 1).setFormulas(aFormulas);
  sheet.getRange(forecastStart, 2, numRows, 1).setFormulas(bFormulas);
  sheet.getRange(forecastStart, 3, numRows, 1).setFormulas(cFormulas);
  sheet.getRange(forecastStart, 4, numRows, 1).setFormulas(dFormulas);
  sheet.getRange(forecastStart, 5, numRows, 1).setFormulas(eFormulas);
  sheet.getRange(forecastStart, 6, numRows, 1).setFormulas(fFormulas);
  sheet.getRange(forecastStart, 7, numRows, 1).setFormulas(gFormulas);
  sheet.getRange(forecastStart, 8, numRows, 1).setFormulas(hFormulas);
  sheet.getRange(forecastStart, 9, numRows, 1).setFormulas(iFormulas);

  setDateFmt(sheet.getRange(forecastStart, 1, numRows, 1));
  setCurrency(sheet.getRange(forecastStart, 3, numRows, 6));
  sheet.getRange(forecastStart, 8, numRows, 1).setFontWeight('bold');

  markFormula(sheet.getRange(forecastStart, 1, numRows, 9));
  protectFormula(sheet.getRange(forecastStart, 1, numRows, 9), 'תחזית תזרים יומית - שדות מחושבים');

  applyEndingBalanceConditionalFormatting(sheet, `H${forecastStart}:H${forecastEnd}`, '$A$6');

  sheet.setColumnWidths(1, 1, 110);
  sheet.setColumnWidths(2, 1, 60);
  sheet.setColumnWidths(3, 1, 150);
  sheet.setColumnWidths(4, 1, 150);
  sheet.setColumnWidths(5, 1, 150);
  sheet.setColumnWidths(6, 1, 150);
  sheet.setColumnWidths(7, 1, 170);
  sheet.setColumnWidths(8, 1, 160);
  sheet.setColumnWidths(9, 1, 110);

  sheet.getRange(forecastStart, 1, numRows, headers.length).setBorder(true, true, true, true, true, true, '#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Monthly outlook block: this month + 3 months ahead. Fixed-expense totals are
 * constant across months (that's what "fixed" means) so the seasonal multiplier
 * column on 4_Fixed_Expenses is the lever to tune them; variable-expense totals
 * are pulled per calendar month from their exact dated rows on that tab, so
 * installment plans naturally land in the month you dated each installment for.
 */
function buildMonthlyOutlook(sheet, forecastStart, forecastEnd, fixedLast) {
  const titleRow = 11;
  const headerRow = 12;
  const dataFirstRow = 13;

  styleSectionRow(sheet, `A${titleRow}:G${titleRow}`, 'תחזית חודשית מסכמת - החודש ועד 3 חודשים קדימה');

  const monthlyHeaders = [
    'חודש',
    'תאריך יעד (סוף חודש)',
    'הוצאות קבועות צפויות (₪)',
    'הוצאות משתנות צפויות (₪)',
    'יתרת סגירה חזויה (₪)',
    'שינוי מיתרה נוכחית (₪)',
    'מצב'
  ];
  sheet.getRange(headerRow, 1, 1, monthlyHeaders.length).setValues([monthlyHeaders]);
  styleHeaderRow(sheet, headerRow, monthlyHeaders.length);

  const monthLabels = ['סוף החודש הנוכחי', 'בעוד חודש', 'בעוד חודשיים', 'בעוד 3 חודשים'];
  monthLabels.forEach((label, i) => {
    const row = dataFirstRow + i;
    sheet.getRange(row, 1).setValue(label);
    sheet.getRange(row, 2).setFormula(`=MIN(EOMONTH(TODAY(),${i}),$A$${forecastEnd})`);
    sheet
      .getRange(row, 3)
      .setFormula(
        `=SUMIF('${CONFIG.sheets.fixed}'!$B$2:$B$${fixedLast},"קבועה",'${CONFIG.sheets.fixed}'!$G$2:$G$${fixedLast})`
      );
    sheet
      .getRange(row, 4)
      .setFormula(
        `=SUMPRODUCT(('${CONFIG.sheets.fixed}'!$B$2:$B$${fixedLast}="משתנה")*(YEAR('${CONFIG.sheets.fixed}'!$D$2:$D$${fixedLast})=YEAR(EOMONTH(TODAY(),${i})))*(MONTH('${CONFIG.sheets.fixed}'!$D$2:$D$${fixedLast})=MONTH(EOMONTH(TODAY(),${i})))*'${CONFIG.sheets.fixed}'!$G$2:$G$${fixedLast})`
      );
    sheet
      .getRange(row, 5)
      .setFormula(
        `=INDEX($H$${forecastStart}:$H$${forecastEnd},MATCH(B${row},$A$${forecastStart}:$A$${forecastEnd},0))`
      );
    sheet.getRange(row, 6).setFormula(`=E${row}-$A$6`);
    sheet
      .getRange(row, 7)
      .setFormula(`=IF(E${row}<0,"🔴 קריטי",IF(E${row}<$A$6*0.15,"🟡 אזהרה","🟢 תקין"))`);
  });

  setDateFmt(sheet.getRange(dataFirstRow, 2, monthLabels.length, 1));
  setCurrency(sheet.getRange(dataFirstRow, 3, monthLabels.length, 4));
  markFormula(sheet.getRange(dataFirstRow, 1, monthLabels.length, 7));
  protectFormula(sheet.getRange(dataFirstRow, 2, monthLabels.length, 6), 'תחזית חודשית - שדות מחושבים');

  sheet
    .getRange(dataFirstRow, 1, monthLabels.length, monthlyHeaders.length)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

function buildKpiCard(sheet, labelA1, valueA1, labelText, valueFormula, type) {
  const labelRange = sheet.getRange(labelA1);
  labelRange.merge();
  labelRange
    .setValue(labelText)
    .setBackground(CONFIG.colors.sectionBg)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setWrap(true);

  const valueRange = sheet.getRange(valueA1);
  valueRange.merge();
  valueRange.setFormula(valueFormula);
  valueRange
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#F3F3F3');
  if (type === 'currency') setCurrency(valueRange);
  sheet.setRowHeight(labelRange.getRow(), 30);
  sheet.setRowHeight(valueRange.getRow(), 34);
}

// ============================================================================
// CONDITIONAL FORMATTING HELPERS
// ============================================================================
function applyStatusConditionalFormatting(sheet, a1Range, statusColorMap) {
  const rules = sheet.getConditionalFormatRules();
  const range = sheet.getRange(a1Range);
  Object.keys(statusColorMap).forEach(status => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(statusColorMap[status])
        .setRanges([range])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);
}

function applyNumericConditionalFormatting(sheet, a1Range, threshold) {
  const rules = sheet.getConditionalFormatRules();
  const range = sheet.getRange(a1Range);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(threshold)
      .setBackground(CONFIG.colors.badBg)
      .setFontColor(CONFIG.colors.badFont)
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(threshold)
      .setBackground(CONFIG.colors.goodBg)
      .setFontColor(CONFIG.colors.goodFont)
      .setRanges([range])
      .build()
  );
  sheet.setConditionalFormatRules(rules);
}

function applyNumericConditionalFormattingBand(sheet, a1Range, lowThreshold, highThreshold) {
  const rules = sheet.getConditionalFormatRules();
  const range = sheet.getRange(a1Range);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(lowThreshold)
      .setBackground(CONFIG.colors.badBg)
      .setFontColor(CONFIG.colors.badFont)
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(lowThreshold, highThreshold)
      .setBackground(CONFIG.colors.warnBg)
      .setFontColor(CONFIG.colors.warnFont)
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(highThreshold)
      .setBackground(CONFIG.colors.goodBg)
      .setFontColor(CONFIG.colors.goodFont)
      .setRanges([range])
      .build()
  );
  sheet.setConditionalFormatRules(rules);
}

function applyEndingBalanceConditionalFormatting(sheet, a1Range, bankBalanceRef) {
  const rules = sheet.getConditionalFormatRules();
  const range = sheet.getRange(a1Range);
  const firstCell = range.getCell(1, 1).getA1Notation();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=${firstCell}<0`)
      .setBackground(CONFIG.colors.badBg)
      .setFontColor(CONFIG.colors.badFont)
      .setBold(true)
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=${firstCell}<${bankBalanceRef}*0.15`)
      .setBackground(CONFIG.colors.warnBg)
      .setFontColor(CONFIG.colors.warnFont)
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=${firstCell}>=${bankBalanceRef}*0.15`)
      .setBackground(CONFIG.colors.goodBg)
      .setFontColor(CONFIG.colors.goodFont)
      .setRanges([range])
      .build()
  );
  sheet.setConditionalFormatRules(rules);
}
