/**
 * ====================================================================
 * MatFinance — Personal Finance Tracker (Google Apps Script Backend)
 * ====================================================================
 */

const SHEET_CATEGORIES = 'Categories';
const SHEET_TRANSACTIONS = 'Transactions';
const SHEET_REPORT = 'Monthly Report';

const DEFAULT_INCOME_CATEGORIES = [
  'Salary', 'Freelance', 'Investment', 'Gift', 'Other Income'
];

const DEFAULT_EXPENSE_CATEGORIES = [
  'Food', 'Transport', 'Housing', 'Bills', 'Entertainment',
  'Health', 'Education', 'Shopping', 'Other Expenses'
];

// ====================================================================
// SETUP
// ====================================================================
function setupSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    _setupCategoriesSheet_(ss);
    _setupTransactionsSheet_(ss);
    _setupMonthlyReportSheet_(ss);

    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }

    SpreadsheetApp.getUi().alert('MatFinance setup complete! Sheets created: Categories, Transactions, Monthly Report.');
  } catch (err) {
    Logger.log('setupSpreadsheet error: ' + err.message);
    try {
      SpreadsheetApp.getUi().alert('Setup failed: ' + err.message);
    } catch (uiErr) {}
  }
}

function _setupCategoriesSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_CATEGORIES);
  if (!sheet) sheet = ss.insertSheet(SHEET_CATEGORIES);
  sheet.clear();

  sheet.getRange(1, 1, 1, 2).setValues([['Income', 'Expense']]);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');

  const maxLen = Math.max(DEFAULT_INCOME_CATEGORIES.length, DEFAULT_EXPENSE_CATEGORIES.length);
  const rows = [];
  for (let i = 0; i < maxLen; i++) {
    rows.push([
      DEFAULT_INCOME_CATEGORIES[i] || '',
      DEFAULT_EXPENSE_CATEGORIES[i] || ''
    ]);
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  sheet.setColumnWidths(1, 2, 180);
  sheet.setFrozenRows(1);
}

function _setupTransactionsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_TRANSACTIONS);
  if (!sheet) sheet = ss.insertSheet(SHEET_TRANSACTIONS);
  sheet.clear();

  const headers = ['ID', 'Date', 'Type', 'Category', 'Description', 'Amount', 'Month'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');

  sheet.getRange('G2').setFormula(
    '=ARRAYFORMULA(IF(LEN(B2:B)=0,"",MID(B2:B,6,2)))'
  );

  sheet.getRange(1, 6, sheet.getMaxRows(), 1).setNumberFormat('#,##0');
  sheet.setColumnWidths(1, 1, 90);
  sheet.setColumnWidths(2, 1, 100);
  sheet.setColumnWidths(3, 1, 90);
  sheet.setColumnWidths(4, 1, 130);
  sheet.setColumnWidths(5, 1, 220);
  sheet.setColumnWidths(6, 1, 120);
  sheet.setColumnWidths(7, 1, 70);
  sheet.setFrozenRows(1);
}

function _setupMonthlyReportSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_REPORT);
  if (!sheet) sheet = ss.insertSheet(SHEET_REPORT);
  sheet.clear();

  const headers = ['Month', 'Income', 'Expense', 'Net Balance'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');

  const months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
  const rows = [];
  for (let i = 0; i < months.length; i++) {
    const r = i + 2;
    rows.push([
      months[i],
      '=SUMIFS(Transactions!$F$2:$F,Transactions!$G$2:$G,A' + r + ',Transactions!$C$2:$C,"Income")',
      '=SUMIFS(Transactions!$F$2:$F,Transactions!$G$2:$G,A' + r + ',Transactions!$C$2:$C,"Expense")',
      '=B' + r + '-C' + r
    ]);
  }
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  sheet.getRange(2, 2, rows.length, 3).setNumberFormat('#,##0');
  sheet.setColumnWidths(1, 4, 130);
  sheet.setFrozenRows(1);
}

// ====================================================================
// WEB APP ENTRY POINT
// ====================================================================
function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('MatFinance')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<h3>MatFinance failed to load</h3><p>' + _escapeHtml_(err.message) + '</p>'
    );
  }
}

function _getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found. Run setupSpreadsheet() first.');
  return sheet;
}

function _escapeHtml_(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ====================================================================
// API
// ====================================================================
function getInitialData() {
  try {
    const catSheet = _getSheet_(SHEET_CATEGORIES);
    const catValues = catSheet.getDataRange().getValues();
    const income = [];
    const expense = [];
    for (let i = 1; i < catValues.length; i++) {
      if (catValues[i][0]) income.push(catValues[i][0]);
      if (catValues[i][1]) expense.push(catValues[i][1]);
    }

    const txSheet = _getSheet_(SHEET_TRANSACTIONS);
    const lastRow = txSheet.getLastRow();
    let transactions = [];
    if (lastRow > 1) {
      const values = txSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      transactions = values
        .filter(function (row) { return row[0] !== '' && row[0] !== null; })
        .map(function (row) {
          return {
            id: String(row[0]),
            date: row[1],
            type: row[2],
            category: row[3],
            description: row[4],
            amount: Number(row[5]) || 0
          };
        });
    }

    return {
      success: true,
      data: {
        categories: { income: income, expense: expense },
        transactions: transactions
      }
    };
  } catch (err) {
    Logger.log('getInitialData error: ' + err.message);
    return { success: false, error: err.message };
  }
}

function addTransaction(tx) {
  try {
    if (!tx || !tx.date || !tx.type || !tx.category || tx.amount === undefined) {
      throw new Error('Transaction details are incomplete.');
    }
    const amount = Number(tx.amount);
    if (isNaN(amount) || amount <= 0) {
      throw new Error('Invalid amount.');
    }
    if (tx.type !== 'Income' && tx.type !== 'Expense') {
      throw new Error('Invalid transaction type.');
    }

    const sheet = _getSheet_(SHEET_TRANSACTIONS);
    const id = Utilities.getUuid();
    
    // FIX FOR SYNC BUG: The ARRAYFORMULA in Column G makes getLastRow() return max sheet rows.
    // We must manually find the first empty row by checking Column A.
    const aValues = sheet.getRange("A:A").getValues();
    let newRow = 2; // Start scanning after the header
    while (newRow <= aValues.length && aValues[newRow - 1][0] !== "") {
      newRow++;
    }

    sheet.getRange(newRow, 1, 1, 6).setValues([[
      id,
      tx.date,
      tx.type,
      tx.category,
      tx.description || '',
      amount
    ]]);

    return { success: true, data: { id: id, clientId: tx.clientId || null } };
  } catch (err) {
    Logger.log('addTransaction error: ' + err.message);
    return { success: false, error: err.message, data: { clientId: tx ? tx.clientId : null } };
  }
}

function deleteTransaction(id) {
  try {
    if (!id) throw new Error('Transaction ID is missing.');
    const sheet = _getSheet_(SHEET_TRANSACTIONS);
    
    // Scan Column A explicitly to avoid getLastRow() overreach
    const aValues = sheet.getRange("A:A").getValues();
    for (let i = 1; i < aValues.length; i++) { // Skip header at i=0
      if (String(aValues[i][0]) === String(id)) {
        sheet.deleteRow(i + 1); // Row numbers are 1-indexed
        return { success: true };
      }
    }
    return { success: true, warning: 'Transaction not found (it may have been deleted already).' };
  } catch (err) {
    Logger.log('deleteTransaction error: ' + err.message);
    return { success: false, error: err.message };
  }
}

function getMonthlyReport() {
  try {
    const sheet = _getSheet_(SHEET_REPORT);
    const lastRow = sheet.getLastRow();
    const raw = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const report = raw.map(function (row) {
      return { month: row[0], income: row[1], expense: row[2], net: row[3] };
    });
    return { success: true, data: report };
  } catch (err) {
    Logger.log('getMonthlyReport error: ' + err.message);
    return { success: false, error: err.message };
  }
}