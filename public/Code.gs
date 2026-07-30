/**
 * MatFinance Backend - Google Apps Script
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MatFinance')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Fetch initial transactions and dynamic categories from Google Sheets
 */
function getInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Get Transactions Sheet
    var txSheet = getOrCreateSheet(ss, 'Transactions', ['ID', 'Date', 'Type', 'Category', 'Description', 'Amount']);
    var txData = fetchTransactions(txSheet);
    
    // 2. Get Categories Sheet
    var categories = fetchCategories(ss);
    
    return {
      success: true,
      data: {
        transactions: txData,
        categories: categories
      }
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Add a single transaction row to the Sheet
 */
function addTransaction(tx) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, 'Transactions', ['ID', 'Date', 'Type', 'Category', 'Description', 'Amount']);
    
    // Generate server ID if it's a temp local ID
    var finalId = (tx.id && !tx.id.startsWith('local-')) ? tx.id : 'tx-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
    
    sheet.appendRow([
      finalId,
      tx.date || '',
      tx.type || 'Expense',
      tx.category || '',
      tx.description || '',
      Number(tx.amount) || 0
    ]);
    
    return { success: true, data: { id: finalId } };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Delete a transaction row by ID
 */
function deleteTransaction(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Transactions');
    if (!sheet) return { success: true };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1); // 1-based index
        return { success: true };
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Helper: Reads categories from the "Categories" tab
 */
function fetchCategories(ss) {
  var sheet = ss.getSheetByName('Categories');
  
  // Default fallbacks if Categories tab doesn't exist yet
  var defaultIncome = ['Salary', 'Freelance', 'Gifts', 'Family Payment', 'Taken from Saving'];
  var defaultExpense = ['Foods & Drinks', 'Car', 'Internet & Sim', 'Entertainment', 'Healthcare', 'Clothing', 'House Maintenance', 'Personal', 'Loan'];
  
  if (!sheet) {
    // Create the tab automatically with defaults
    sheet = ss.insertSheet('Categories');
    sheet.appendRow(['Income', 'Expense']);
    var maxRows = Math.max(defaultIncome.length, defaultExpense.length);
    for (var r = 0; r < maxRows; r++) {
      sheet.appendRow([defaultIncome[r] || '', defaultExpense[r] || '']);
    }
    return { income: defaultIncome, expense: defaultExpense };
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { income: defaultIncome, expense: defaultExpense };
  }
  
  var incomeList = [];
  var expenseList = [];
  
  // Header detection (Row 1)
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var incomeCol = headers.indexOf('income');
  var expenseCol = headers.indexOf('expense');
  
  // If headers "Income" and "Expense" exist in Row 1:
  if (incomeCol !== -1 || expenseCol !== -1) {
    for (var i = 1; i < data.length; i++) {
      if (incomeCol !== -1 && data[i][incomeCol]) {
        var incVal = String(data[i][incomeCol]).trim();
        if (incVal) incomeList.push(incVal);
      }
      if (expenseCol !== -1 && data[i][expenseCol]) {
        var expVal = String(data[i][expenseCol]).trim();
        if (expVal) expenseList.push(expVal);
      }
    }
  } else {
    // Alternative format: Column A = Category Name, Column B = Type ('Income' or 'Expense')
    for (var i = 1; i < data.length; i++) {
      var catName = String(data[i][0]).trim();
      var catType = String(data[i][1]).trim().toLowerCase();
      if (catName) {
        if (catType === 'income') {
          incomeList.push(catName);
        } else {
          expenseList.push(catName);
        }
      }
    }
  }
  
  return {
    income: incomeList.length > 0 ? incomeList : defaultIncome,
    expense: expenseList.length > 0 ? expenseList : defaultExpense
  };
}

/**
 * Helper: Read transactions from sheet
 */
function fetchTransactions(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // Skip empty rows
    
    list.push({
      id: String(row[0]),
      date: String(row[1] || ''),
      type: String(row[2] || 'Expense'),
      category: String(row[3] || ''),
      description: String(row[4] || ''),
      amount: Number(row[5]) || 0
    });
  }
  return list;
}

/**
 * Helper: Get sheet or create if missing
 */
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}