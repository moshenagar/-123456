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
    instructions: '0_הוראות_שימוש',
    banks: '1_Banks_Credit',
    receivables: '2_Receivables',
    payables: '3_Payables',
    fixed: '4_Fixed_Expenses',
    employees: 'עלות_עובדים_ותמחור',
    tax: '5_Tax_Engine',
    pricing: '6_Pricing_BreakEven',
    dashboard: '7_Master_Forecast_Dashboard',
    visualDashboard: '8_Dashboard_Visual'
  },
  rows: {
    banksData: 8,            // bank accounts - real liquid cash
    creditCardsData: 10,     // credit cards - liabilities, need frequent updating
    loansData: 8,            // loans - liabilities
    receivablesData: 150,   // data rows 2..151
    payablesData: 150,      // data rows 2..151
    fixedData: 25,          // data rows 2..26
    employeesData: 15,      // employee cost/pricing rows
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

    buildInstructionsTab(ss);
    buildBanksCreditTab(ss);
    buildReceivablesTab(ss);
    buildPayablesTab(ss);
    buildFixedExpensesTab(ss);
    buildEmployeesTab(ss);
    buildTaxEngineTab(ss);
    buildPricingTab(ss);
    buildDashboardTab(ss);
    buildVisualDashboardTab(ss);

    reorderSheets(ss);
    removeDefaultBlankSheet(ss);

    SpreadsheetApp.flush();
    ss.setActiveSheet(ss.getSheetByName(CONFIG.sheets.dashboard));

    ui.alert(
      'המערכת הותקנה בהצלחה ✅',
      'כל הטאבים נבנו ועודכנו.\n\n' +
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
    CONFIG.sheets.instructions,
    CONFIG.sheets.banks,
    CONFIG.sheets.receivables,
    CONFIG.sheets.payables,
    CONFIG.sheets.fixed,
    CONFIG.sheets.employees,
    CONFIG.sheets.tax,
    CONFIG.sheets.pricing,
    CONFIG.sheets.dashboard,
    CONFIG.sheets.visualDashboard
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

/**
 * Single source of truth for 1_Banks_Credit's row layout: three stacked
 * sections (Banks / Credit Cards / Loans), each with a title, header, data
 * block and total row. Every other tab that needs to reference a total on
 * this sheet (e.g. "the real cash balance") should call this instead of
 * recomputing row offsets inline, to avoid drift bugs across tabs.
 */
function getBanksLayout() {
  const banksTitleRow = 1;
  const banksHeaderRow = 2;
  const banksFirst = 3;
  const banksLast = banksFirst + CONFIG.rows.banksData - 1;
  const banksTotalRow = banksLast + 1;

  const cardsTitleRow = banksTotalRow + 2;
  const cardsHeaderRow = cardsTitleRow + 1;
  const cardsFirst = cardsHeaderRow + 1;
  const cardsLast = cardsFirst + CONFIG.rows.creditCardsData - 1;
  const cardsTotalRow = cardsLast + 1;

  const loansTitleRow = cardsTotalRow + 2;
  const loansHeaderRow = loansTitleRow + 1;
  const loansFirst = loansHeaderRow + 1;
  const loansLast = loansFirst + CONFIG.rows.loansData - 1;
  const loansTotalRow = loansLast + 1;

  return {
    banksTitleRow,
    banksHeaderRow,
    banksFirst,
    banksLast,
    banksTotalRow,
    cardsTitleRow,
    cardsHeaderRow,
    cardsFirst,
    cardsLast,
    cardsTotalRow,
    loansTitleRow,
    loansHeaderRow,
    loansFirst,
    loansLast,
    loansTotalRow
  };
}

/**
 * Single source of truth for 6_Pricing_BreakEven's row layout: title, the
 * monthly overhead pull, the company-wide break-even summary block, then the
 * per-product table. Other tabs (7_Master_Forecast_Dashboard) that need this
 * sheet's totals should call this instead of hardcoding row numbers.
 */
function getPricingLayout() {
  const titleRow = 1;
  const overheadRow = 2;
  const summaryTitleRow = 4;
  const revenueRow = 5;
  const marginRow = 6;
  const safetyRow = 7;
  const statusRow = 8;
  const hourlyTitleRow = 10;
  const hourlyInputRow = 11;
  const hourlyCalcRow = 12;
  const hourlyHoursRow = 13;
  const hourlyStatusRow = 14;
  const headerRow = 16;
  const first = headerRow + 1;
  const last = first + CONFIG.rows.pricingData - 1;
  return {
    titleRow,
    overheadRow,
    summaryTitleRow,
    revenueRow,
    marginRow,
    safetyRow,
    statusRow,
    hourlyTitleRow,
    hourlyInputRow,
    hourlyCalcRow,
    hourlyHoursRow,
    hourlyStatusRow,
    headerRow,
    first,
    last
  };
}

/**
 * Single source of truth for עלות_עובדים_ותמחור's row layout: a statutory-
 * rate parameters block (National Insurance employer brackets, pension,
 * severance - shared by all employee rows since these are the same in law
 * for everyone), then the per-employee cost/pricing table.
 */
function getEmployeesLayout() {
  const titleRow = 1;
  const paramsTitleRow = 3;
  const paramsFirst = 4;
  const paramsLast = 8;
  const headerRow = 10;
  const first = headerRow + 1;
  const last = first + CONFIG.rows.employeesData - 1;
  const totalRow = last + 2;
  return { titleRow, paramsTitleRow, paramsFirst, paramsLast, headerRow, first, last, totalRow };
}

/**
 * Builds a SUMPRODUCT formula fragment (no leading "=") that totals
 * 4_Fixed_Expenses rows occurring on/in a given target date or month.
 * Each row is defined by: C=first payment date, D=frequency in months,
 * E=number of payments (blank/0 = unlimited), H=final outflow amount.
 * A row "occurs" at the target when the target is on/after the first date,
 * the number of months since the first date is an exact multiple of the
 * frequency, and (for a finite payment count) that occurrence index hasn't
 * been exhausted yet. This single model covers monthly/bi-monthly/quarterly
 * recurring costs and multi-installment one-off expenses from one row each.
 *
 * dayLevel=true also requires the day-of-month to match (used for the daily
 * 90-day forecast); dayLevel=false matches at month granularity only (used
 * for the monthly outlook, where any occurrence within the month counts).
 */
function buildExpenseOccurrenceFormula(fixedLast, targetExpr, dayLevel, typeFilter, vatFilter) {
  const sheetName = CONFIG.sheets.fixed;
  const C = `'${sheetName}'!$C$2:$C$${fixedLast}`;
  const D = `'${sheetName}'!$D$2:$D$${fixedLast}`;
  const E = `'${sheetName}'!$E$2:$E$${fixedLast}`;
  const H = `'${sheetName}'!$H$2:$H$${fixedLast}`;
  const B = `'${sheetName}'!$B$2:$B$${fixedLast}`;
  const I = `'${sheetName}'!$I$2:$I$${fixedLast}`;

  const monthsSince = `((YEAR(${targetExpr})-YEAR(${C}))*12+(MONTH(${targetExpr})-MONTH(${C})))`;
  const freqSafe = `MAX(${D},1)`;
  const countSafe = `IF(${E}=0,999999,${E})`;
  const typeTerm = typeFilter ? `(${B}="${typeFilter}")*` : '';
  const vatTerm = vatFilter ? `(${I}="${vatFilter}")*` : '';
  const onOrAfter = dayLevel ? `(${targetExpr}>=${C})*` : `(${monthsSince}>=0)*`;
  const dayMatch = dayLevel ? `(DAY(${targetExpr})=DAY(${C}))*` : '';

  return (
    `SUMPRODUCT(${typeTerm}${vatTerm}(${C}<>"")*${onOrAfter}${dayMatch}` +
    `(MOD(${monthsSince},${freqSafe})=0)*` +
    `(${monthsSince}/${freqSafe}<${countSafe})*${H})`
  );
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
// TAB 0: 0_הוראות_שימוש (Instructions)
// ============================================================================
function buildInstructionsTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.instructions);
  sheet.setTabColor('#434343');
  sheet.setRightToLeft(true);

  styleTitleRow(sheet, 'A1:B1', 'מדריך שימוש מהיר במערכת');

  sheet.getRange('A3').setValue('כלל אצבע: תאים צהובים = אתם ממלאים. תאים אפורים = המערכת מחשבת לבד, אל תיגעו.');
  sheet.getRange('A3').setFontWeight('bold').setFontSize(11).setWrap(true);
  sheet.getRange('A3:B3').merge();

  sheet
    .getRange('A4')
    .setValue(
      '💡 מאיפה מתחילים בפועל: אין צורך להזין אף חשבונית כדי לדעת כמה כסף העסק צריך להכניס בחודש. ' +
        'מלאו קודם רק את טאב 4 (הוצאות קבועות) וטאב 6 (תמחור ונקודת איזון) - זה כבר נותן תשובה: כמה מחזור/כמה שעות/כמה יחידות נדרשות לאיזון. ' +
        'רק אחר כך, כשרוצים לעקוב יום-יום אחרי תזרים אמיתי, מוסיפים בהדרגה חשבוניות בטאבים 2-3.'
    );
  sheet.getRange('A4:B4').merge();
  sheet.getRange('A4').setWrap(true).setFontStyle('italic').setFontSize(10).setBackground('#FFF2CC');
  sheet.setRowHeight(4, 55);

  const sections = [
    {
      title: '1️⃣ בנקים, כרטיסי אשראי והלוואות',
      color: '#1F3864',
      fill: 'שמות החשבונות/כרטיסים/הלוואות שלכם והיתרות שלהם. בבנקים - כמה כסף באמת יש לכם בעו"ש עכשיו (זה הבסיס לכל התחזית). בכרטיסי אשראי - כמה חייבים על כל כרטיס ולאיזה בנק זה מתחייב (עדכנו שוטף). בהלוואות - מה נשאר לשלם, כמה לחודש, וכמה תשלומים נותרו.',
      get: 'תמונת מצב אמיתית של הנכסים וההתחייבויות שלכם.'
    },
    {
      title: '2️⃣ לקוחות (הכנסות שאמורות להיכנס)',
      color: '#38761D',
      fill: 'לכל חשבונית - למי, כמה, מתי הוצאה, תנאי התשלום, וכמה אתם בטוחים שתקבלו (אחוז).',
      get: 'תאריך משוער שבו הכסף באמת ייכנס לבנק, בהתחשב באיחורים אפשריים.'
    },
    {
      title: '3️⃣ ספקים (הוצאות שאתם צריכים לשלם)',
      color: '#B45309',
      fill: 'לכל חשבונית מספק - למי, כמה, מתי, ותנאי התשלום.',
      get: 'תאריך משוער שבו הכסף באמת ייצא מהבנק.'
    },
    {
      title: '4️⃣ הוצאות קבועות ומשתנות',
      color: '#741B47',
      fill: 'כל הוצאה (שכר, שכירות, ארנונה, חשמל...) - בחרו קטגוריה, מתי משלמים בפעם הראשונה, כל כמה חודשים זה חוזר, וכמה. מע"מ מתמלא אוטומטית לפי הקטגוריה שבחרתם.',
      get: 'כל ההוצאות שלכם מסודרות במקום אחד, שמוזנות אוטומטית לתחזית ולחישוב המע"מ.'
    },
    {
      title: '👥 עלות עובדים ותמחור שעתי',
      color: '#674EA7',
      fill: 'לכל עובד - שכר ברוטו, שעות עבודה חייבות ללקוח בחודש, ואחוז הרווח הרצוי. עלויות המעסיק על פי חוק (ביטוח לאומי, פנסיה, פיצויים) מחושבות אוטומטית מפרמטרים משותפים בראש הטאב.',
      get: 'כמה עולה לכם באמת כל עובד לחודש, וכמה חייבים לגבות מלקוח על כל שעת עבודה שלו כדי להרוויח. חשוב: הוסיפו את העלות הכוללת שמתקבלת כאן כשורה בטאב 4, כדי שהיא תיכנס לתחזית ולנקודת האיזון הכוללת.'
    },
    {
      title: '5️⃣ מנוע מיסים',
      color: '#990000',
      fill: 'רק פעם אחת בהתחלה - אחוז המע"מ, אחוז מקדמת מס הכנסה, וסכום ביטוח לאומי.',
      get: 'כמה כסף אתם חייבים לשים בצד למיסים (מע"מ + מקדמות + ביטוח לאומי), וכמה כסף באמת שלכם אחרי זה ("יתרה חופשית ממיסים").'
    },
    {
      title: '6️⃣ תמחור ונקודת איזון',
      color: '#0B5394',
      fill: 'לכל מוצר/שירות - כמה עולה לכם, כמה מוכרים אותו, וכמה יחידות אתם מצפים למכור בחודש. ולעסקי שירות שמוכרים שעות - כמה שעות זמינות לכם בחודש (משרה מלאה/חלקית) ומה התעריף השעתי.',
      get: 'האם כל מוצר רווחי באמת (אחרי עמלות סליקה וכו\'), וכמה כסף/שעות/יחידות חייב להיכנס בחודש כדי שכדאי יהיה להחזיק את העסק פתוח (נקודת האיזון הכוללת - לפי מחזור או לפי שעות).'
    },
    {
      title: '7️⃣ לוח בקרה (הלשונית לפני האחרונה - כאן מסתכלים)',
      color: '#38761D',
      fill: 'כלום, חוץ מ"מצב תרחיש" למעלה (רגיל / קיצון) אם רוצים לבדוק תרחיש גרוע.',
      get: 'כל מה שחשוב במבט אחד - כמה כסף יש עכשיו, כמה יהיה בעוד 15/30/60/90 יום, מתי צפויה הנקודה הכי נמוכה, האם העסק מעל נקודת האיזון, ותחזית יום-אחר-יום ל-90 יום קדימה.'
    },
    {
      title: '8️⃣ דאשבורד ויזואלי למנהלים',
      color: '#B45309',
      fill: 'כלום - כל התרשימים כאן מבוססים אוטומטית על שאר הטאבים.',
      get: 'תמונה חזותית להצגה למנהלים/שותפים: גרף קו של יתרת הבנק ל-90 יום, עמודות של התחזית החודשית, השוואת הכנסות מול תקורה (נקודת האיזון), ופיצול ההוצאות הקבועות/משתנות לפי קטגוריה - הכל בעמוד אחד, נוח להדפסה או שיתוף מסך.'
    }
  ];

  let row = 5;
  sections.forEach(section => {
    sheet.getRange(row, 1, 1, 2).merge();
    sheet
      .getRange(row, 1)
      .setValue(section.title)
      .setBackground(section.color)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(row, 30);
    row++;

    sheet.getRange(row, 1).setValue('מה למלא').setFontWeight('bold').setVerticalAlignment('top');
    sheet.getRange(row, 2).setValue(section.fill).setWrap(true).setVerticalAlignment('top');
    sheet.setRowHeight(row, 60);
    row++;

    sheet.getRange(row, 1).setValue('מה מקבלים').setFontWeight('bold').setVerticalAlignment('top');
    sheet.getRange(row, 2).setValue(section.get).setWrap(true).setVerticalAlignment('top');
    sheet.setRowHeight(row, 50);
    row++;

    row++; // spacer
  });

  sheet.getRange(3, 1, row - 3, 2).setBorder(true, true, true, true, true, true, '#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidths(1, 1, 150);
  sheet.setColumnWidths(2, 1, 550);
  sheet.setFrozenRows(3);
}

// ============================================================================
// TAB 1: 1_Banks_Credit
// ============================================================================
function buildBanksCreditTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.banks);
  sheet.setTabColor('#1F3864');
  const L = getBanksLayout();

  // ---- Section 1: Banks (real liquid cash - the "exit point" of every asset) ----
  styleTitleRow(sheet, `A${L.banksTitleRow}:E${L.banksTitleRow}`, 'בנקים - הנכסים הנזילים שלי');
  const bankHeaders = [
    'שם חשבון בנק',
    'יתרת עו"ש נוכחית (₪)',
    'מסגרת משיכת יתר בעו"ש (₪)',
    'אחוז כרית ביטחון לחירום',
    'נזילות זמינה מותאמת סיכון (₪)'
  ];
  sheet.getRange(L.banksHeaderRow, 1, 1, bankHeaders.length).setValues([bankHeaders]);
  styleHeaderRow(sheet, L.banksHeaderRow, bankHeaders.length);
  sheet.setFrozenRows(L.banksHeaderRow);

  markInput(sheet.getRange(L.banksFirst, 1, CONFIG.rows.banksData, 4));
  fillDefaultColumn(sheet, L.banksFirst, L.banksLast, 4, 0.05);

  const liquidityFormulas = colFormulas(L.banksFirst, L.banksLast, r => `=IF($B${r}="","",$B${r}+$C${r}*(1-$D${r}))`);
  sheet.getRange(L.banksFirst, 5, liquidityFormulas.length, 1).setFormulas(liquidityFormulas);
  markFormula(sheet.getRange(L.banksFirst, 5, CONFIG.rows.banksData, 1));
  protectFormula(sheet.getRange(L.banksFirst, 5, CONFIG.rows.banksData, 1), 'נזילות מותאמת - שדה מחושב');

  setCurrency(sheet.getRange(L.banksFirst, 2, CONFIG.rows.banksData, 2));
  setCurrency(sheet.getRange(L.banksFirst, 5, CONFIG.rows.banksData, 1));
  setPercent(sheet.getRange(L.banksFirst, 4, CONFIG.rows.banksData, 1));

  sheet.getRange(L.banksTotalRow, 1).setValue('סה"כ בנקים').setFontWeight('bold');
  sheet.getRange(L.banksTotalRow, 2).setFormula(`=SUM(B${L.banksFirst}:B${L.banksLast})`);
  sheet.getRange(L.banksTotalRow, 3).setFormula(`=SUM(C${L.banksFirst}:C${L.banksLast})`);
  sheet.getRange(L.banksTotalRow, 5).setFormula(`=SUM(E${L.banksFirst}:E${L.banksLast})`);
  sheet.getRange(L.banksTotalRow, 1, 1, 5).setBackground(CONFIG.colors.totalBg).setFontWeight('bold');
  setCurrency(sheet.getRange(L.banksTotalRow, 2, 1, 1));
  setCurrency(sheet.getRange(L.banksTotalRow, 3, 1, 1));
  setCurrency(sheet.getRange(L.banksTotalRow, 5, 1, 1));
  protectFormula(sheet.getRange(L.banksTotalRow, 2, 1, 2), 'סה"כ - שדה מחושב');
  protectFormula(sheet.getRange(L.banksTotalRow, 5, 1, 1), 'סה"כ - שדה מחושב');
  sheet
    .getRange('A' + L.banksTotalRow)
    .setNote(
      'יתרת הבנק הכוללת ("B' + L.banksTotalRow + '") היא הבסיס היחיד ליתרת הבנק בכל המערכת (טאב 5 וטאב 7) - כרטיסי אשראי והלוואות אינם נכללים כאן, הם ניהול נפרד למטה.'
    );

  sheet.getRange(L.banksFirst, 1, CONFIG.rows.banksData, bankHeaders.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // ---- Section 2: Credit Cards (liabilities - need frequent manual updating) ----
  styleTitleRow(sheet, `A${L.cardsTitleRow}:E${L.cardsTitleRow}`, 'כרטיסי אשראי - יש לעדכן באופן שוטף');
  const cardHeaders = [
    'שם כרטיס אשראי',
    'יתרת חיוב נוכחית לחיוב הבא (₪)',
    'מסגרת אשראי מאושרת (₪)',
    'יום חיוב חודשי',
    'חשבון בנק מקושר לחיוב'
  ];
  sheet.getRange(L.cardsHeaderRow, 1, 1, cardHeaders.length).setValues([cardHeaders]);
  styleHeaderRow(sheet, L.cardsHeaderRow, cardHeaders.length);

  markInput(sheet.getRange(L.cardsFirst, 1, CONFIG.rows.creditCardsData, 4));
  markInput(sheet.getRange(L.cardsFirst, 5, CONFIG.rows.creditCardsData, 1));
  sheet.getRange(L.cardsFirst, 4, CONFIG.rows.creditCardsData, 1).setDataValidation(numberRangeValidation(1, 31));
  sheet
    .getRange(L.cardsFirst, 5, CONFIG.rows.creditCardsData, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(sheet.getRange(L.banksFirst, 1, CONFIG.rows.banksData, 1), true)
        .setAllowInvalid(true)
        .build()
    );

  setCurrency(sheet.getRange(L.cardsFirst, 2, CONFIG.rows.creditCardsData, 2));

  sheet.getRange(L.cardsTotalRow, 1).setValue('סה"כ כרטיסי אשראי').setFontWeight('bold');
  sheet.getRange(L.cardsTotalRow, 2).setFormula(`=SUM(B${L.cardsFirst}:B${L.cardsLast})`);
  sheet.getRange(L.cardsTotalRow, 3).setFormula(`=SUM(C${L.cardsFirst}:C${L.cardsLast})`);
  sheet.getRange(L.cardsTotalRow, 1, 1, 5).setBackground(CONFIG.colors.totalBg).setFontWeight('bold');
  setCurrency(sheet.getRange(L.cardsTotalRow, 2, 1, 2));
  protectFormula(sheet.getRange(L.cardsTotalRow, 2, 1, 2), 'סה"כ - שדה מחושב');
  sheet
    .getRange('A' + L.cardsTotalRow)
    .setNote(
      'כדי שהחיוב החודשי ייכנס אוטומטית לתחזית התזרים בטאב 7, הוסיפו שורה מתאימה בטאב 4 (הוצאות קבועות) עם יום/תאריך החיוב, תדירות=1 והסכום.'
    );

  sheet.getRange(L.cardsFirst, 1, CONFIG.rows.creditCardsData, cardHeaders.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // ---- Section 3: Loans (liabilities) ----
  styleTitleRow(sheet, `A${L.loansTitleRow}:G${L.loansTitleRow}`, 'הלוואות');
  const loanHeaders = [
    'שם הלוואה / גורם מלווה',
    'סכום מקורי (₪)',
    'יתרה לתשלום (₪)',
    'תשלום חודשי (₪)',
    'יום חיוב חודשי',
    'מס\' תשלומים שנותרו',
    'ריבית שנתית %'
  ];
  sheet.getRange(L.loansHeaderRow, 1, 1, loanHeaders.length).setValues([loanHeaders]);
  styleHeaderRow(sheet, L.loansHeaderRow, loanHeaders.length);

  markInput(sheet.getRange(L.loansFirst, 1, CONFIG.rows.loansData, 7));
  sheet.getRange(L.loansFirst, 5, CONFIG.rows.loansData, 1).setDataValidation(numberRangeValidation(1, 31));

  setCurrency(sheet.getRange(L.loansFirst, 2, CONFIG.rows.loansData, 3));
  setPercent(sheet.getRange(L.loansFirst, 7, CONFIG.rows.loansData, 1));

  sheet.getRange(L.loansTotalRow, 1).setValue('סה"כ הלוואות').setFontWeight('bold');
  sheet.getRange(L.loansTotalRow, 3).setFormula(`=SUM(C${L.loansFirst}:C${L.loansLast})`);
  sheet.getRange(L.loansTotalRow, 4).setFormula(`=SUM(D${L.loansFirst}:D${L.loansLast})`);
  sheet.getRange(L.loansTotalRow, 1, 1, 7).setBackground(CONFIG.colors.totalBg).setFontWeight('bold');
  setCurrency(sheet.getRange(L.loansTotalRow, 3, 1, 2));
  protectFormula(sheet.getRange(L.loansTotalRow, 3, 1, 2), 'סה"כ - שדה מחושב');
  sheet
    .getRange('A' + L.loansTotalRow)
    .setNote(
      'כדי שההחזר החודשי ייכנס אוטומטית לתחזית התזרים בטאב 7, הוסיפו שורה מתאימה בטאב 4 (הוצאות קבועות) עם תדירות=1 ומס\' תשלומים = "מס\' תשלומים שנותרו" כאן.'
    );

  sheet.getRange(L.loansFirst, 1, CONFIG.rows.loansData, loanHeaders.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidths(1, 1, 220);
  sheet.setColumnWidths(2, 3, 170);
  sheet.setColumnWidths(5, 1, 150);
  sheet.setColumnWidths(6, 1, 170);
  sheet.setColumnWidths(7, 1, 130);
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
    'תאריך תשלום ראשון',
    'תדירות (כל כמה חודשים)',
    'מס\' תשלומים (ריק = ללא הגבלה, חוזר תמיד)',
    'סכום בסיס לתשלום (₪)',
    'מכפיל דינמי / כרית הגנה (למשל 1.25 לעונתיות)',
    'סה"כ תזרים צפוי לתשלום (₪)',
    'כולל מע"מ?'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, 1, headers.length);
  sheet.setFrozenRows(1);

  const first = 2;
  const last = 1 + CONFIG.rows.fixedData;
  const n = CONFIG.rows.fixedData;

  const categoryNames = CONFIG.expenseVatBank.map(item => item.name);
  markInput(sheet.getRange(first, 1, n, 7));
  markInput(sheet.getRange(first, 9, n, 1));

  sheet.getRange(first, 1, n, 1).setDataValidation(listValidation(categoryNames));
  sheet.getRange(first, 2, n, 1).setDataValidation(listValidation(CONFIG.expenseTypes));
  sheet.getRange(first, 3, n, 1).setDataValidation(dateValidation());
  sheet.getRange(first, 4, n, 1).setDataValidation(numberRangeValidation(1, 36));
  sheet.getRange(first, 5, n, 1).setDataValidation(numberRangeValidation(0, 360));
  sheet.getRange(first, 9, n, 1).setDataValidation(listValidation(['כן', 'לא']));
  fillDefaultColumn(sheet, first, last, 2, 'קבועה');
  fillDefaultColumn(sheet, first, last, 4, 1);
  fillDefaultColumn(sheet, first, last, 7, 1);
  fillDefaultColumn(sheet, first, last, 9, 'כן');

  const outflowFormulas = colFormulas(first, last, r => `=IF($F${r}="","",$F${r}*$G${r})`);
  sheet.getRange(first, 8, n, 1).setFormulas(outflowFormulas);
  markFormula(sheet.getRange(first, 8, n, 1));
  protectFormula(sheet.getRange(first, 8, n, 1), 'תזרים - שדה מחושב');

  setCurrency(sheet.getRange(first, 6, n, 1));
  setDateFmt(sheet.getRange(first, 3, n, 1));

  const totalRow = last + 2;
  sheet.getRange(totalRow, 1).setValue('סה"כ (כל תשלום בטבלה, ללא התחשבות בתדירות)').setFontWeight('bold');
  sheet.getRange(totalRow, 8).setFormula(`=SUM(H${first}:H${last})`);
  sheet.getRange(totalRow, 1, 1, 9).setBackground(CONFIG.colors.totalBg).setFontWeight('bold');
  setCurrency(sheet.getRange(totalRow, 8, 1, 1));
  protectFormula(sheet.getRange(totalRow, 8, 1, 1), 'סה"כ - שדה מחושב');
  sheet
    .getRange('A' + totalRow)
    .setNote(
      'שורה אחת מכסה גם הוצאות דו-חודשיות/רבעוניות (למשל ארנונה/חשמל) וגם תשלומים מפוצלים: קבעו "תאריך תשלום ראשון", "תדירות" (1=כל חודש, 2=דו-חודשי וכו\') ו"מס\' תשלומים" (השאירו ריק בקבועות שחוזרות תמיד; מלאו מספר סופי במשתנות/תשלומים). כל תשלום ייכנס אוטומטית לחודש הנכון בטאב 7 ובתחזית החודשית - אין צורך בשורה נפרדת לכל תשלום. תקורת "קבועה" בטאב 6 מחושבת כשווה-ערך חודשי (סכום חלקי תדירות).'
    );

  sheet.setColumnWidths(1, 1, 210);
  sheet.setColumnWidths(2, 1, 140);
  sheet.setColumnWidths(3, 1, 150);
  sheet.setColumnWidths(4, 1, 160);
  sheet.setColumnWidths(5, 1, 220);
  sheet.setColumnWidths(6, 1, 160);
  sheet.setColumnWidths(7, 1, 220);
  sheet.setColumnWidths(8, 1, 190);
  sheet.setColumnWidths(9, 1, 110);

  sheet.getRange(first, 1, n, headers.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  buildExpenseVatBankReference(sheet);
}

/**
 * Writes the visible, editable reference table of common expense categories with
 * their default VAT applicability, used both as the source list for the category
 * dropdown (column A) and, via onEdit(), to auto-fill the "כולל מע"מ?" column.
 */
function buildExpenseVatBankReference(sheet) {
  const startCol = 11; // column K, leaving column J as a spacer
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
 * Simple onEdit trigger on 4_Fixed_Expenses:
 *  - Column A (category) edited -> auto-fill "כולל מע"מ?" (column I) from the VAT bank.
 *  - Column B (type) edited -> auto-fill sensible defaults for "תדירות" (D) and
 *    "מס' תשלומים" (E) ONLY when those cells are still blank, so it never clobbers
 *    a value the user already set. קבועה defaults to monthly/unlimited; משתנה
 *    defaults to a single one-time payment.
 * All auto-fills stay plain values the user can freely overwrite afterward.
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
    if (row < first || row > last) return;

    if (col === 1) {
      const category = String(e.range.getValue()).trim();
      if (!category) return;
      const match = CONFIG.expenseVatBank.find(item => item.name === category);
      if (match && match.vat) {
        sheet.getRange(row, 9).setValue(match.vat);
      }
      return;
    }

    if (col === 2) {
      const type = String(e.range.getValue()).trim();
      const freqCell = sheet.getRange(row, 4);
      const countCell = sheet.getRange(row, 5);
      if (type === 'קבועה') {
        if (freqCell.getValue() === '') freqCell.setValue(1);
        // Count stays blank = unlimited/ongoing.
      } else if (type === 'משתנה') {
        if (freqCell.getValue() === '') freqCell.setValue(1);
        if (countCell.getValue() === '') countCell.setValue(1);
      }
      return;
    }
  } catch (err) {
    // Never block manual editing due to an auto-fill failure.
  }
}

// ============================================================================
// TAB: עלות_עובדים_ותמחור (Employee Cost & Hourly Pricing)
// ============================================================================
function buildEmployeesTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.employees);
  sheet.setTabColor('#674EA7');
  const E = getEmployeesLayout();

  styleTitleRow(sheet, `A${E.titleRow}:N${E.titleRow}`, 'עלות עובדים ותמחור שעתי');

  styleSectionRow(
    sheet,
    `A${E.paramsTitleRow}:D${E.paramsTitleRow}`,
    'פרמטרים לפי חוק (אותו דבר לכל העובדים - יש לוודא מול רואה חשבון/חברת שכר, השיעורים מתעדכנים מעת לעת)'
  );
  const params = [
    ['שכר ממוצע במשק (₪, לצורך מדרגת ביטוח לאומי)', 10551],
    ['אחוז ביטוח לאומי מעסיק - עד 60% מהשכר הממוצע', 0.0355],
    ['אחוז ביטוח לאומי מעסיק - מעל 60% מהשכר הממוצע', 0.076],
    ['אחוז הפרשת מעסיק לפנסיה (תגמולים)', 0.065],
    ['אחוז הפרשת מעסיק לפיצויים', 0.0833]
  ];
  sheet.getRange(E.paramsFirst, 1, params.length, 2).setValues(params);
  markInput(sheet.getRange(E.paramsFirst, 2, params.length, 1));
  setCurrency(sheet.getRange(E.paramsFirst, 2, 1, 1));
  setPercent(sheet.getRange(E.paramsFirst + 1, 2, 4, 1));
  sheet
    .getRange(`A${E.paramsFirst}`)
    .setNote(
      'ברירות המחדל משקפות את הצו וההנחיות הנפוצות (פנסיית חובה: 6.5% תגמולי מעסיק + 8.33% פיצויים; ביטוח לאומי מעסיק במדרגות לפי אחוז מהשכר הממוצע במשק) - אך הסכומים והשיעורים מתעדכנים מעת לעת. עדכנו כאן לפי הנתונים העדכניים לפני שמסתמכים על התוצאה.'
    );

  const headers = [
    'שם עובד',
    'שכר ברוטו חודשי (₪)',
    'ביטוח לאומי מעסיק (₪)',
    'הפרשת מעסיק לפנסיה (₪)',
    'הפרשת מעסיק לפיצויים (₪)',
    'הבראה שנתית (₪)',
    'הבראה - שווה ערך חודשי (₪)',
    'הוצאות נוספות קבועות לעובד (₪)',
    'עלות מעסיק חודשית כוללת (₪)',
    'שעות עבודה חייבות ללקוח בחודש',
    'עלות לשעת עבודה (₪)',
    'אחוז רווח רצוי מעל העלות',
    'תעריף מומלץ ללקוח לשעה (₪)',
    'רווח בפועל לשעה (₪)'
  ];
  sheet.getRange(E.headerRow, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, E.headerRow, headers.length);
  sheet.setFrozenRows(E.headerRow);

  const first = E.first;
  const last = E.last;
  const n = CONFIG.rows.employeesData;

  markInput(sheet.getRange(first, 1, n, 2));
  markInput(sheet.getRange(first, 6, n, 1));
  markInput(sheet.getRange(first, 8, n, 1));
  markInput(sheet.getRange(first, 10, n, 1));
  markInput(sheet.getRange(first, 12, n, 1));

  sheet.getRange(first, 10, n, 1).setDataValidation(numberRangeValidation(1, 744));
  sheet.getRange(first, 12, n, 1).setDataValidation(numberRangeValidation(0, 5));
  fillDefaultColumn(sheet, first, last, 10, 160);
  fillDefaultColumn(sheet, first, last, 12, 0.3);

  const empFormulas = {
    C: r => `=IF($B${r}="","",IF($B${r}<=$B$4*0.6,$B${r}*$B$5,$B$4*0.6*$B$5+($B${r}-$B$4*0.6)*$B$6))`,
    D: r => `=IF($B${r}="","",$B${r}*$B$7)`,
    E: r => `=IF($B${r}="","",$B${r}*$B$8)`,
    G: r => `=IF($F${r}="","",$F${r}/12)`,
    I: r => `=IF($B${r}="","",$B${r}+$C${r}+$D${r}+$E${r}+$G${r}+$H${r})`,
    K: r => `=IF(OR($J${r}="",$J${r}=0),"",$I${r}/$J${r})`,
    M: r => `=IF($K${r}="","",$K${r}*(1+$L${r}))`,
    N: r => `=IF($K${r}="","",$M${r}-$K${r})`
  };
  const empColIndex = { C: 3, D: 4, E: 5, G: 7, I: 9, K: 11, M: 13, N: 14 };
  Object.keys(empFormulas).forEach(key => {
    const formulas = colFormulas(first, last, empFormulas[key]);
    sheet.getRange(first, empColIndex[key], n, 1).setFormulas(formulas);
  });

  const empFormulaCols = [3, 4, 5, 7, 9, 11, 13, 14];
  empFormulaCols.forEach(c => markFormula(sheet.getRange(first, c, n, 1)));
  protectFormula(sheet.getRange(first, 3, n, 3), 'שדות מחושבים - עלות מעסיק');
  protectFormula(sheet.getRange(first, 7, n, 1), 'שדה מחושב - הבראה חודשית');
  protectFormula(sheet.getRange(first, 9, n, 1), 'שדה מחושב - עלות כוללת');
  protectFormula(sheet.getRange(first, 11, n, 1), 'שדה מחושב - עלות לשעה');
  protectFormula(sheet.getRange(first, 13, n, 2), 'שדות מחושבים - תעריף ורווח');

  setCurrency(sheet.getRange(first, 2, n, 8));
  setCurrency(sheet.getRange(first, 11, n, 1));
  setPercent(sheet.getRange(first, 12, n, 1));
  setCurrency(sheet.getRange(first, 13, n, 2));

  const totalRow = E.totalRow;
  sheet.getRange(totalRow, 1).setValue('סה"כ עלות מעסיק חודשית - כל העובדים').setFontWeight('bold');
  sheet.getRange(totalRow, 9).setFormula(`=SUM(I${first}:I${last})`);
  sheet.getRange(totalRow, 1, 1, 14).setBackground(CONFIG.colors.totalBg).setFontWeight('bold');
  setCurrency(sheet.getRange(totalRow, 9, 1, 1));
  protectFormula(sheet.getRange(totalRow, 9, 1, 1), 'סה"כ - שדה מחושב');
  sheet
    .getRange('A' + totalRow)
    .setNote(
      'הסכום הזה נכנס אוטומטית לחישוב התקורה הכוללת בטאב 6 (נקודת איזון) - אין צורך להעתיק אותו לשום מקום כדי לדעת כמה כסף העסק צריך להכניס בחודש. ' +
        'אבל כדי שתשלומי השכר יופיעו בתאריך הנכון בתחזית התזרים היומית (טאב 7), עדיין צריך להוסיף לכל עובד שורה בטאב 4 עם יום/תאריך התשלום בפועל, תדירות=1, סוג=קבועה, כולל מע"מ=לא - זה משפיע רק על התזמון היומי, לא על נקודת האיזון.'
    );

  sheet.getRange(first, 1, n, headers.length).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidths(1, 1, 170);
  sheet.setColumnWidths(2, 1, 150);
  sheet.setColumnWidths(3, 3, 170);
  sheet.setColumnWidths(6, 1, 140);
  sheet.setColumnWidths(7, 1, 180);
  sheet.setColumnWidths(8, 1, 190);
  sheet.setColumnWidths(9, 1, 190);
  sheet.setColumnWidths(10, 1, 190);
  sheet.setColumnWidths(11, 1, 160);
  sheet.setColumnWidths(12, 1, 190);
  sheet.setColumnWidths(13, 1, 190);
  sheet.setColumnWidths(14, 1, 160);
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
    .setFormula(`=${buildExpenseOccurrenceFormula(fixedLast, 'TODAY()', false, null, 'כן')}/(1+$B$4)*$B$4`);
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
  sheet.getRange('B22').setFormula(`='${CONFIG.sheets.banks}'!B${getBanksLayout().banksTotalRow}`);
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
  const P = getPricingLayout();

  styleTitleRow(sheet, `A${P.titleRow}:R${P.titleRow}`, 'תמחור ונקודת איזון - כלכלת יחידה (Unit Economics)');
  sheet
    .getRange(`A${P.overheadRow}`)
    .setValue('סה"כ תקורה חודשית קבועה (שורות "קבועה" מטאב 4 חוץ מ"שכר עבודה", שווה-ערך חודשי, + עלות כל העובדים מטאב העובדים)');
  const fixedLastRow = 1 + CONFIG.rows.fixedData;
  const employeesTotalRow = getEmployeesLayout().totalRow;
  sheet
    .getRange(`B${P.overheadRow}`)
    .setFormula(
      `=SUMPRODUCT(('${CONFIG.sheets.fixed}'!$B$2:$B$${fixedLastRow}="קבועה")*('${CONFIG.sheets.fixed}'!$A$2:$A$${fixedLastRow}<>"שכר עבודה")*'${CONFIG.sheets.fixed}'!$H$2:$H$${fixedLastRow}/MAX('${CONFIG.sheets.fixed}'!$D$2:$D$${fixedLastRow},1))+'${CONFIG.sheets.employees}'!I${employeesTotalRow}`
    );
  setCurrency(sheet.getRange(`B${P.overheadRow}`));
  sheet.getRange(`B${P.overheadRow}`).setFontWeight('bold').setFontSize(12);
  markFormula(sheet.getRange(`A${P.overheadRow}:B${P.overheadRow}`));
  protectFormula(sheet.getRange(`B${P.overheadRow}`), 'תקורה חודשית - שדה מחושב מטאב 4 + טאב עובדים');
  sheet
    .getRange(`A${P.overheadRow}`)
    .setNote(
      'הסכום הזה כולל אוטומטית גם את "סה"כ עלות מעסיק חודשית" מטאב העובדים - אין צורך להעתיק את זה ידנית לטאב 4. ' +
        'כדי שתשלומי השכר יופיעו בתאריך הנכון בתחזית התזרים היומית (טאב 7), עדיין צריך להוסיף בטאב 4 שורה לכל עובד עם יום/תאריך התשלום בפועל - אבל חשוב לבחור בקטגוריה "שכר עבודה" בדיוק, כדי שהמערכת תדע לא לספור את הסכום פעמיים בחישוב נקודת האיזון כאן (השורה עדיין תיכנס נכון לתזרים היומי בטאב 7, רק לא תיכנס בכפילות לתקורה כאן).'
    );

  // ---- Company-wide break-even: how much revenue justifies keeping the business open ----
  styleSectionRow(sheet, `A${P.summaryTitleRow}:D${P.summaryTitleRow}`, 'נקודת איזון כוללת לעסק (כל המוצרים יחד)');

  sheet.getRange(`A${P.revenueRow}`).setValue('סה"כ הכנסות חודשיות צפויות (₪)');
  sheet.getRange(`B${P.revenueRow}`).setFormula(`=SUMPRODUCT($C$${P.first}:$C$${P.last},$L$${P.first}:$L$${P.last})`);
  sheet.getRange(`C${P.revenueRow}`).setValue('סה"כ תרומה חודשית צפויה (₪)');
  sheet.getRange(`D${P.revenueRow}`).setFormula(`=SUMPRODUCT($J$${P.first}:$J$${P.last},$L$${P.first}:$L$${P.last})`);

  sheet.getRange(`A${P.marginRow}`).setValue('אחוז תרומה משוקלל (%)');
  sheet.getRange(`B${P.marginRow}`).setFormula(`=IFERROR(D${P.revenueRow}/B${P.revenueRow},0)`);
  sheet.getRange(`C${P.marginRow}`).setValue('נקודת איזון - מחזור חודשי לעסק (₪)');
  sheet
    .getRange(`D${P.marginRow}`)
    .setFormula(`=IF(B${P.marginRow}>0,$B$${P.overheadRow}/B${P.marginRow},"אין נקודת איזון (תרומה שלילית/אפס)")`);

  sheet.getRange(`A${P.safetyRow}`).setValue('מרווח ביטחון (₪)');
  sheet
    .getRange(`B${P.safetyRow}`)
    .setFormula(`=IF(ISNUMBER(D${P.marginRow}),B${P.revenueRow}-D${P.marginRow},"—")`);
  sheet.getRange(`C${P.safetyRow}`).setValue('מרווח ביטחון (%)');
  sheet
    .getRange(`D${P.safetyRow}`)
    .setFormula(`=IF(AND(ISNUMBER(B${P.safetyRow}),B${P.revenueRow}>0),B${P.safetyRow}/B${P.revenueRow},"—")`);

  sheet.getRange(`A${P.statusRow}`).setValue('סטטוס העסק');
  sheet
    .getRange(`B${P.statusRow}:D${P.statusRow}`)
    .merge()
    .setFormula(
      `=IF(NOT(ISNUMBER(D${P.marginRow})),"🔴 אין תרומה חיובית - בדקו מחירי מכירה מול עלויות",` +
        `IF(D${P.safetyRow}>0.2,"🟢 מעל נקודת האיזון בבטחה (מרווח ביטחון "&TEXT(D${P.safetyRow},"0%")&")",` +
        `IF(D${P.safetyRow}>=0,"🟡 קרוב לנקודת האיזון (מרווח ביטחון "&TEXT(D${P.safetyRow},"0%")&")",` +
        `"🔴 מתחת לנקודת האיזון (חסר "&TEXT(-B${P.safetyRow},"#,##0")&" ₪ בחודש)")))`
    )
    .setHorizontalAlignment('center');

  setCurrency(sheet.getRange(`B${P.revenueRow}`));
  setCurrency(sheet.getRange(`D${P.revenueRow}`));
  setPercent(sheet.getRange(`B${P.marginRow}`));
  setCurrency(sheet.getRange(`D${P.marginRow}`));
  setCurrency(sheet.getRange(`B${P.safetyRow}`));
  setPercent(sheet.getRange(`D${P.safetyRow}`));

  sheet.getRange(`A${P.revenueRow}:D${P.safetyRow}`).setFontWeight('bold');
  sheet.getRange(`B${P.statusRow}:D${P.statusRow}`).setFontWeight('bold').setFontSize(12);
  markFormula(sheet.getRange(`B${P.revenueRow}`));
  markFormula(sheet.getRange(`D${P.revenueRow}`));
  markFormula(sheet.getRange(`B${P.marginRow}`));
  markFormula(sheet.getRange(`D${P.marginRow}`));
  markFormula(sheet.getRange(`B${P.safetyRow}`));
  markFormula(sheet.getRange(`D${P.safetyRow}`));
  markFormula(sheet.getRange(`B${P.statusRow}:D${P.statusRow}`));
  protectFormula(sheet.getRange(`B${P.revenueRow}`), 'נקודת איזון כוללת - שדה מחושב');
  protectFormula(sheet.getRange(`D${P.revenueRow}`), 'נקודת איזון כוללת - שדה מחושב');
  protectFormula(sheet.getRange(`B${P.marginRow}`), 'נקודת איזון כוללת - שדה מחושב');
  protectFormula(sheet.getRange(`D${P.marginRow}`), 'נקודת איזון כוללת - שדה מחושב');
  protectFormula(sheet.getRange(`B${P.safetyRow}`), 'נקודת איזון כוללת - שדה מחושב');
  protectFormula(sheet.getRange(`D${P.safetyRow}`), 'נקודת איזון כוללת - שדה מחושב');
  protectFormula(sheet.getRange(`B${P.statusRow}:D${P.statusRow}`), 'נקודת איזון כוללת - שדה מחושב');

  sheet
    .getRange(`A${P.revenueRow}:D${P.statusRow}`)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  sheet
    .getRange(`A${P.statusRow}`)
    .setNote(
      'נקודת האיזון הכוללת מחושבת מהתקורה החודשית (B' +
        P.overheadRow +
        ') חלקי אחוז התרומה המשוקלל של כל המוצרים ביחד, לפי הכמויות הצפויות שהוזנו בטבלה למטה. שנו כמויות/מחירים בטבלה כדי לראות איך זה משפיע.'
    );

  // ---- Hourly/time-based break-even: for service businesses that sell hours, not units ----
  styleSectionRow(sheet, `A${P.hourlyTitleRow}:D${P.hourlyTitleRow}`, 'תמחור לפי שעת עבודה (לעסקי שירות)');

  sheet.getRange(`A${P.hourlyInputRow}`).setValue('שעות עבודה זמינות לחודש (משרה מלאה = 160)');
  sheet.getRange(`B${P.hourlyInputRow}`).setValue(160);
  sheet.getRange(`C${P.hourlyInputRow}`).setValue('תעריף שעתי נוכחי/מתוכנן (₪)');
  sheet.getRange(`D${P.hourlyInputRow}`).setValue(0);
  markInput(sheet.getRange(`B${P.hourlyInputRow}`));
  markInput(sheet.getRange(`D${P.hourlyInputRow}`));
  sheet.getRange(`B${P.hourlyInputRow}`).setDataValidation(numberRangeValidation(1, 744));
  sheet.getRange(`D${P.hourlyInputRow}`).setDataValidation(numberRangeValidation(0, 100000));

  sheet.getRange(`A${P.hourlyCalcRow}`).setValue('הכנסה חודשית צפויה בתעריף הנוכחי (₪)');
  sheet.getRange(`B${P.hourlyCalcRow}`).setFormula(`=B${P.hourlyInputRow}*D${P.hourlyInputRow}`);
  sheet.getRange(`C${P.hourlyCalcRow}`).setValue('תעריף שעתי נדרש לאיזון (₪/שעה)');
  sheet
    .getRange(`D${P.hourlyCalcRow}`)
    .setFormula(`=IF(B${P.hourlyInputRow}>0,$B$${P.overheadRow}/B${P.hourlyInputRow},"—")`);

  sheet.getRange(`A${P.hourlyHoursRow}`).setValue('מס\' שעות נדרש בתעריף הנוכחי לכיסוי ההוצאות');
  sheet
    .getRange(`B${P.hourlyHoursRow}`)
    .setFormula(`=IF(D${P.hourlyInputRow}>0,$B$${P.overheadRow}/D${P.hourlyInputRow},"—")`);

  sheet.getRange(`A${P.hourlyStatusRow}`).setValue('סטטוס (שעתי)');
  sheet
    .getRange(`B${P.hourlyStatusRow}:D${P.hourlyStatusRow}`)
    .merge()
    .setFormula(
      `=IF(OR(B${P.hourlyInputRow}=0,D${P.hourlyInputRow}=0),"⚪ הזינו שעות זמינות ותעריף שעתי כדי לראות סטטוס",` +
        `IF((B${P.hourlyCalcRow}-$B$${P.overheadRow})/B${P.hourlyCalcRow}>0.2,"🟢 מעל נקודת האיזון בבטחה",` +
        `IF(B${P.hourlyCalcRow}>=$B$${P.overheadRow},"🟡 קרוב לנקודת האיזון",` +
        `"🔴 מתחת לנקודת האיזון (חסר "&TEXT($B$${P.overheadRow}-B${P.hourlyCalcRow},"#,##0")&" ₪ בחודש)")))`
    )
    .setHorizontalAlignment('center');

  setCurrency(sheet.getRange(`D${P.hourlyInputRow}`));
  setCurrency(sheet.getRange(`B${P.hourlyCalcRow}`));
  setCurrency(sheet.getRange(`D${P.hourlyCalcRow}`));

  sheet.getRange(`A${P.hourlyInputRow}:D${P.hourlyHoursRow}`).setFontWeight('bold');
  sheet.getRange(`B${P.hourlyStatusRow}:D${P.hourlyStatusRow}`).setFontWeight('bold').setFontSize(12);
  markFormula(sheet.getRange(`B${P.hourlyCalcRow}`));
  markFormula(sheet.getRange(`D${P.hourlyCalcRow}`));
  markFormula(sheet.getRange(`B${P.hourlyHoursRow}`));
  markFormula(sheet.getRange(`B${P.hourlyStatusRow}:D${P.hourlyStatusRow}`));
  protectFormula(sheet.getRange(`B${P.hourlyCalcRow}`), 'תמחור שעתי - שדה מחושב');
  protectFormula(sheet.getRange(`D${P.hourlyCalcRow}`), 'תמחור שעתי - שדה מחושב');
  protectFormula(sheet.getRange(`B${P.hourlyHoursRow}`), 'תמחור שעתי - שדה מחושב');
  protectFormula(sheet.getRange(`B${P.hourlyStatusRow}:D${P.hourlyStatusRow}`), 'תמחור שעתי - שדה מחושב');

  sheet
    .getRange(`A${P.hourlyInputRow}:D${P.hourlyStatusRow}`)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  sheet
    .getRange(`A${P.hourlyStatusRow}`)
    .setNote(
      'לעסק שמוכר שעות עבודה (ולא מוצרים) - הזינו כמה שעות זמינות לכם בחודש (תלוי במשרה מלאה/חלקית) ואת התעריף השעתי. תראו כאן גם מה התעריף המינימלי לאיזון, וגם כמה שעות חייבים לעבוד בתעריף הנוכחי.'
    );

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
  sheet.getRange(P.headerRow, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, P.headerRow, headers.length);
  sheet.setFrozenRows(P.headerRow);

  const first = P.first;
  const last = P.last;
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

  const banksTotalRow = getBanksLayout().banksTotalRow;
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
    'סטטוס נקודת איזון - העסק כולו',
    `='${CONFIG.sheets.pricing}'!B${getPricingLayout().statusRow}`,
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
    fFormulas.push([`=${buildExpenseOccurrenceFormula(fixedLast, `$A${r}`, true, null)}*$E$3`]);
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
 * Monthly outlook block: this month + 3 months ahead. Both the fixed and
 * variable expense totals are recomputed per month via buildExpenseOccurrenceFormula,
 * so a bi-monthly/quarterly recurring cost (e.g. ארנונה every 2 months) only shows
 * up in the months it actually falls due, and multi-installment plans land exactly
 * in the months their payments were scheduled for - no duplicate rows needed. The
 * seasonal multiplier column on 4_Fixed_Expenses is still the lever to tune amounts.
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
    const monthTarget = `EOMONTH(TODAY(),${i})`;
    sheet.getRange(row, 3).setFormula(`=${buildExpenseOccurrenceFormula(fixedLast, monthTarget, false, 'קבועה')}`);
    sheet.getRange(row, 4).setFormula(`=${buildExpenseOccurrenceFormula(fixedLast, monthTarget, false, 'משתנה')}`);
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
// TAB 8: 8_Dashboard_Visual - chart-based view for managers/partners
// ============================================================================
function buildVisualDashboardTab(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.sheets.visualDashboard);
  sheet.setTabColor('#B45309');
  sheet.setRightToLeft(true);

  // Idempotent rebuild: drop any charts from a previous run before re-adding them.
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));

  styleTitleRow(sheet, 'A1:J1', '📊 דאשבורד ויזואלי למנהלים');
  sheet.getRange('A3:J3').merge();
  sheet
    .getRange('A3')
    .setValue('כל התרשימים כאן מתעדכנים אוטומטית משאר הטאבים - אין צורך למלא או לערוך כאן דבר. נוח להציג/להדפיס בפגישות.')
    .setFontStyle('italic')
    .setFontSize(10)
    .setBackground('#FFF2CC')
    .setWrap(true);
  sheet.setRowHeight(3, 30);

  const dashboardSheet = ss.getSheetByName(CONFIG.sheets.dashboard);
  const P = getPricingLayout();

  const forecastStart = CONFIG.rows.forecastStart;
  const forecastEnd = forecastStart + CONFIG.rows.forecastDays - 1;
  const monthlyHeaderRow = 12; // header row of the monthly-outlook table on Tab 7
  const monthlyLastRow = 16; // 4 data rows (this month .. +3 months)

  // ---- Helper (backing) data for charts that need local aggregation ----
  const helperRow = 40;
  sheet
    .getRange(helperRow, 1)
    .setValue('🔧 נתוני עזר לתרשימים - נוצר אוטומטית, אין לערוך')
    .setFontWeight('bold')
    .setFontColor('#999999');

  const revHeaderRow = helperRow + 1;
  const revLabelRow = helperRow + 2;
  const overheadLabelRow = helperRow + 3;
  sheet.getRange(revHeaderRow, 1).setValue('מדד');
  sheet.getRange(revHeaderRow, 2).setValue('סכום (₪)');
  sheet.getRange(revLabelRow, 1).setValue('הכנסות חודשיות צפויות');
  sheet.getRange(revLabelRow, 2).setFormula(`='${CONFIG.sheets.pricing}'!B${P.revenueRow}`);
  sheet.getRange(overheadLabelRow, 1).setValue('תקורה חודשית (נקודת איזון)');
  sheet.getRange(overheadLabelRow, 2).setFormula(`='${CONFIG.sheets.pricing}'!B${P.overheadRow}`);
  setCurrency(sheet.getRange(revLabelRow, 2, 2, 1));
  markFormula(sheet.getRange(revLabelRow, 2, 2, 1));
  protectFormula(sheet.getRange(revLabelRow, 2, 2, 1), 'נתוני עזר לתרשים - שדה מחושב');

  const fixedFirst = 2;
  const fixedLast = 1 + CONFIG.rows.fixedData;
  const categories = CONFIG.expenseVatBank.map(item => item.name).filter(name => name !== 'אחר (הקלדה חופשית)');
  const categoryHeaderRow = helperRow + 5;
  const categoryFirst = categoryHeaderRow + 1;
  const categoryLast = categoryFirst + categories.length - 1;
  sheet.getRange(categoryHeaderRow, 1).setValue('קטגוריה');
  sheet.getRange(categoryHeaderRow, 2).setValue('סה"כ חודשי (₪)');
  categories.forEach((name, i) => {
    const row = categoryFirst + i;
    sheet.getRange(row, 1).setValue(name);
    sheet
      .getRange(row, 2)
      .setFormula(
        `=SUMIFS('${CONFIG.sheets.fixed}'!$H$${fixedFirst}:$H$${fixedLast},'${CONFIG.sheets.fixed}'!$A$${fixedFirst}:$A$${fixedLast},$A${row})`
      );
  });
  setCurrency(sheet.getRange(categoryFirst, 2, categories.length, 1));
  markFormula(sheet.getRange(categoryFirst, 2, categories.length, 1));
  protectFormula(sheet.getRange(categoryFirst, 2, categories.length, 1), 'נתוני עזר לתרשים - שדה מחושב');

  sheet.hideRows(helperRow, categoryLast - helperRow + 1);

  // ---- Chart 1: 90-day cash balance trend (data lives on Tab 7) ----
  const cashChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dashboardSheet.getRange(forecastStart - 1, 1, forecastEnd - forecastStart + 2, 1))
    .addRange(dashboardSheet.getRange(forecastStart - 1, 8, forecastEnd - forecastStart + 2, 1))
    .setOption('title', 'תזרים מזומנים - יתרת בנק צפויה ל-90 יום')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [CONFIG.colors.sectionBg])
    .setOption('hAxis', { title: 'תאריך', slantedText: true })
    .setOption('vAxis', { title: 'יתרה (₪)' })
    .setOption('width', 620)
    .setOption('height', 360)
    .setPosition(5, 1, 0, 0)
    .build();
  sheet.insertChart(cashChart);

  // ---- Chart 2: monthly closing-balance outlook (data lives on Tab 7) ----
  const monthlyChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dashboardSheet.getRange(monthlyHeaderRow, 1, monthlyLastRow - monthlyHeaderRow + 1, 1))
    .addRange(dashboardSheet.getRange(monthlyHeaderRow, 5, monthlyLastRow - monthlyHeaderRow + 1, 1))
    .setOption('title', 'תחזית חודשית - יתרת סגירה צפויה')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#38761D'])
    .setOption('vAxis', { title: 'יתרה (₪)' })
    .setOption('width', 620)
    .setOption('height', 360)
    .setPosition(5, 8, 0, 0)
    .build();
  sheet.insertChart(monthlyChart);

  // ---- Chart 3: expected revenue vs. break-even overhead (data lives on Tab 6) ----
  const beChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(revHeaderRow, 1, 3, 2))
    .setOption('title', 'הכנסות צפויות מול תקורה (נקודת איזון)')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#0B5394'])
    .setOption('vAxis', { title: '₪ לחודש' })
    .setOption('width', 620)
    .setOption('height', 360)
    .setPosition(24, 1, 0, 0)
    .build();
  sheet.insertChart(beChart);

  // ---- Chart 4: fixed/variable expense breakdown by category (data lives on Tab 4) ----
  const pieChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(categoryHeaderRow, 1, categories.length + 1, 2))
    .setOption('title', 'התפלגות ההוצאות הקבועות/משתנות לפי קטגוריה')
    .setOption('pieSliceText', 'percentage')
    .setOption('width', 620)
    .setOption('height', 360)
    .setPosition(24, 8, 0, 0)
    .build();
  sheet.insertChart(pieChart);
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
