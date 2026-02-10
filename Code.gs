//--GLOBALS--
var A_ASCII_CODE = 'A'.charCodeAt(0);

// Template Constants
var TEMPLATE_SHEET_ID = '1Fse7LA-frH8CnIXmCmpVwg9_1TRLxHj2r_5XDqrvd4E';
var TEMPLATE_CATEGORIES_SHEET_NAME = 'Categories Dropdown';
var TEMPLATE_TRANSACTION_SHEET_NAME = 'Transactions';
var TEMPLATE_EXPENSES_SHEET_NAME = '2018 Expenses';
var TEMPLATE_INCOME_SHEET_NAME = '2018 Income';
var TEMPLATE_SUMMARY_SHEET_NAME = '2018 Summary';

// Column Constants
var A_COL = colNameToNum('A');
var B_COL = colNameToNum('B');
var C_COL = colNameToNum('C');
var R_COL = colNameToNum('R');
var CATEGORY_COL = colNameToNum('A');
var SUB_CATEGORY_COL = colNameToNum('C');
var MONTHLY_TOTALS_COL = SUB_CATEGORY_COL;
var JAN_COL = colNameToNum('D');
var DEC_COL = colNameToNum('O');
var TOTAL_COL = colNameToNum('P');
var AVG_COL = colNameToNum('Q');
var BUDGET_COL = colNameToNum('R');
var TRANSACTION_CAT_COL = colNameToNum('F');
var TRANSACTION_SUBCAT_COL = colNameToNum('G');
var DATE_ROW = 3;

// Menu Names
var MAIN_MENU = "SpendTab";
var SETUP_BUDGET = 'Setup SpendTab';
var ENTER_TRANSACTIONS_MENU = 'Enter Transactions';
var FIX_FORMULAS_MENU = 'Fix Formulas';
var FIX_CATEGORY_DROPDOWNS_MENU = 'Fix Category Dropdowns';

// Sheet Names
var TRANSACTION_SHEET_NAME = 'Transactions';
var EXPENSES_SHEET_NAME_SUFFIX = ' Expenses';
var INCOME_SHEET_NAME_SUFFIX = ' Income';
var SUMMARY_SHEET_NAME_SUFFIX = ' Summary';
var EXPENSES_SHEET_NAME = getCurrentYear() + EXPENSES_SHEET_NAME_SUFFIX;
var INCOME_SHEET_NAME = getCurrentYear() + INCOME_SHEET_NAME_SUFFIX;
var SUMMARY_SHEET_NAME = getCurrentYear() + SUMMARY_SHEET_NAME_SUFFIX;
var CATEGORIES_SHEET_NAME = 'Categories Dropdown';


///////////////////////////////////////////////////////////
// Public methods
///////////////////////////////////////////////////////////

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var menu;
  var installTriggers = false;
  if (e && e.authMode == "NONE") {
    // Add a normal menu item (works in all authorization modes).
    menu = SpreadsheetApp.getUi().createAddonMenu();
    menu.addItem(SETUP_BUDGET, 'enableBudgetTracker');
  } else {
    // Add a menu item based on properties (doesn't work in AuthMode.NONE).
    menu = SpreadsheetApp.getUi().createMenu(MAIN_MENU);
    var properties = PropertiesService.getDocumentProperties();
    var budgetTrackerVersion = properties.getProperty('budgetTrackerVersion');
    if (budgetTrackerVersion !== null) {
      installTriggers = true;
      addMenuItems(menu);
    } else {
      menu.addItem(SETUP_BUDGET, 'enableBudgetTracker');
    }
  }
  menu.addToUi();
  
  if (installTriggers) {
    ensureTriggersInstalled();
  }
}

function enableBudgetTracker() {
  var properties = PropertiesService.getDocumentProperties();
  properties.setProperty('budgetTrackerVersion', 1);
  
  // Create any missing sheets
  setupBudgetSheets();
  addMenuItems(SpreadsheetApp.getUi().createMenu(MAIN_MENU)).addToUi();
  
  ensureTriggersInstalled();
}

function addMenuItems(menu) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var missingCurrentSheets = 
      (   spreadsheet.getSheetByName(EXPENSES_SHEET_NAME) == null 
       && spreadsheet.getSheetByName(INCOME_SHEET_NAME) == null
       && spreadsheet.getSheetByName(SUMMARY_SHEET_NAME_SUFFIX) == null);
  var needToCreateCurrentYearSheets = 
      (   missingCurrentSheets
       && spreadsheet.getSheetByName(getPreviousYear() + EXPENSES_SHEET_NAME_SUFFIX) != null
       && spreadsheet.getSheetByName(getPreviousYear() + INCOME_SHEET_NAME_SUFFIX) != null
       && spreadsheet.getSheetByName(getPreviousYear() + SUMMARY_SHEET_NAME_SUFFIX) != null);
  if (needToCreateCurrentYearSheets) {
    menu.addItem('Create ' + getCurrentYear() + ' Budget Sheets', 'createCurrentYearSheets');
  }
  
  return menu.addItem(ENTER_TRANSACTIONS_MENU, 'openEnterTransactionsDialog')
             .addSeparator()
             .addItem(FIX_FORMULAS_MENU, 'fixAllSheetFormulas')
             .addItem(FIX_CATEGORY_DROPDOWNS_MENU, 'fixCategoryAndSubCategory');
}

function ensureTriggersInstalled() {
  var properties = PropertiesService.getDocumentProperties();
  var onEditTriggerId = properties.getProperty('onEditTriggerId');
  if (onEditTriggerId == null) {
    Logger.log('Installing trigger...');
    onEditTriggerId = ScriptApp.newTrigger('onEdit')
        .forSpreadsheet(SpreadsheetApp.getActive().getId())
        .onEdit()
        .create()
        .getUniqueId();
    properties.setProperty('onEditTriggerId', onEditTriggerId);
  } else {
    Logger.log('Trigger already installed');
  }
}

///////////////////////////////////////////////////////////
// TRIGGER: onEdit
///////////////////////////////////////////////////////////

function onEdit(e) {
  try {
    var spreadsheet = e.source;
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    if (!sheetName.endsWith(TRANSACTION_SHEET_NAME)) {
      return;
    }
    var categoryAffected = (e.range.getColumn() <= TRANSACTION_CAT_COL && e.range.getLastColumn() >= TRANSACTION_CAT_COL);
    if (!categoryAffected) {
      return;
    }
    
    var categorySheetName = CATEGORIES_SHEET_NAME
    if (sheetName != TRANSACTION_SHEET_NAME) {
      var prefix = sheetName.replace(TRANSACTION_SHEET_NAME, '');
      categorySheetName = prefix + CATEGORIES_SHEET_NAME;
    }
    
    fixSubCategory(spreadsheet, sheet, e.range, categorySheetName);
  } catch(err) {
    Logger.log('Unexpected exception in onEdit: ' + err);
  }
}

///////////////////////////////////////////////////////////
// Support for setting up initial sheets
///////////////////////////////////////////////////////////

function setupBudgetSheets() {
  var templateSpreadsheet = SpreadsheetApp.openById(TEMPLATE_SHEET_ID);
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Create 'Setup' Sheet
  var setupSheet = copySheet(templateSpreadsheet, currentSpreadsheet, 'Setup').activate();
  currentSpreadsheet.setNamedRange('StartingDate', setupSheet.getRange('C13'));
  currentSpreadsheet.setNamedRange('StartingBalance', setupSheet.getRange('C14'));
  setupSheet.protect().setWarningOnly(true).setUnprotectedRanges([setupSheet.getRange('C14')]);
  
  // Create 'Categories' Sheet
  var categoriesSheet = copySheet(templateSpreadsheet, currentSpreadsheet, TEMPLATE_CATEGORIES_SHEET_NAME).hideSheet();
  categoriesSheet.protect().setWarningOnly(true);
  
  // Create 'Transactions' Sheet
  var transactionsSheet = copySheet(templateSpreadsheet, currentSpreadsheet, TEMPLATE_TRANSACTION_SHEET_NAME);
  setupSheet.getRange('C13').setFormula('=iferror(if(filter(Transactions!B:B,Transactions!B:B<>"")<>"",eomonth(min(Transactions!B:B), -1) + 1), today())');
  
  // Create 'Expenses' Sheet
  var expensesSheet = copySheet(templateSpreadsheet, currentSpreadsheet, TEMPLATE_EXPENSES_SHEET_NAME);
  setupExpensesSheet(expensesSheet);
  
  // Create 'Income' Sheet
  var incomeSheet = copySheet(templateSpreadsheet, currentSpreadsheet, TEMPLATE_INCOME_SHEET_NAME);
  setupIncomeSheet(incomeSheet);
  fixCategory();
  
  // Create 'Summary' Sheet
  var summarySheet = copySheet(templateSpreadsheet, currentSpreadsheet, TEMPLATE_SUMMARY_SHEET_NAME);
  setupSummarySheet(summarySheet, '=(StartingBalance+D28)-D29');
  
  // NamedRanges don't copy over, so all those formulas need to be fixed up
  // 1st: Fix Expenses sheet Avg column
  var expenseCategories = findCategoryRows(expensesSheet);
  var range = expensesSheet.getDataRange();
  expenseCategories.forEach(function(category) {
    category.subCategories.forEach(function(subcategory) {
      updateAverageFormula(range, subcategory);
    });
  });
  
  // 2nd: Fix Income sheet Avg column
  var incomeCategories = findCategoryRows(incomeSheet);
  var range = incomeSheet.getDataRange();
  incomeCategories.forEach(function(category) {
    category.subCategories.forEach(function(subcategory) {
      updateAverageFormula(range, subcategory);
    });
  });
  SpreadsheetApp.flush();
}

function copySheet(srcSpreadsheet, destSpreadsheet, sheetName) {
  var destSheet = destSpreadsheet.getSheetByName(sheetName);
  if (destSheet == null) {
    var srcSheet = srcSpreadsheet.getSheetByName(sheetName);
    destSheet = srcSheet.copyTo(destSpreadsheet);
    destSheet.setName(sheetName);
  }
  
  // Remove NamedRanges
  var namedRanges = destSheet.getNamedRanges();
  for (var i = 0; i < namedRanges.length; ++i) {
    namedRanges[i].remove();
  }
  
  return destSheet;
}

function setupExpensesSheet(newExpenseSheet) {
  newExpenseSheet.setName(EXPENSES_SHEET_NAME)
  newExpenseSheet.getRange('D2').setValue(getCurrentYear());
}

function setupIncomeSheet(newIncomeSheet) {
  newIncomeSheet.setName(INCOME_SHEET_NAME)
  newIncomeSheet.getRange('D2').setValue(getCurrentYear());
}

function setupSummarySheet(newSummarySheet, startBalanceFormula) {
  newSummarySheet.setName(SUMMARY_SHEET_NAME)
  newSummarySheet.getRange('E21').setValue(getCurrentYear());
  newSummarySheet.getRange('E22').setFormula(startBalanceFormula);
  newSummarySheet.getRange('E23').setValue(EXPENSES_SHEET_NAME);
  newSummarySheet.getRange('E24').setValue(INCOME_SHEET_NAME);
  newSummarySheet.getRange('D31').setFormula('=(E22+D28)-D29');
  newSummarySheet.getRange('C35').setFormula('=unique({0}!A:A)'.format("'" + INCOME_SHEET_NAME + "'"));
  newSummarySheet.getRange('B36').setFormula('=ArrayFormula(iferror(match(C36:C45,{0}!A:A,0)))'.format("'" + INCOME_SHEET_NAME + "'"));
  newSummarySheet.getRange('C48').setFormula('=unique({0}!A:A)'.format("'" + EXPENSES_SHEET_NAME + "'"));
  newSummarySheet.getRange('B49').setFormula('=ArrayFormula(iferror(match(C49:C64,{0}!A:A,0)))'.format("'" + EXPENSES_SHEET_NAME + "'"));
  newSummarySheet.protect().setWarningOnly(true);
}

///////////////////////////////////////////////////////////
// Support for archiving Previous Year
///////////////////////////////////////////////////////////

function createCurrentYearSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Copy the Expenses Sheet
  var oldExpenseSheetName = getPreviousYear() + EXPENSES_SHEET_NAME_SUFFIX;
  var oldExpenseSheet = spreadsheet.getSheetByName(oldExpenseSheetName);
  if (oldExpenseSheet != null) {
    var newExpenseSheet = oldExpenseSheet.copyTo(spreadsheet);
    setupExpensesSheet(newExpenseSheet);
    newExpenseSheet.activate();
    spreadsheet.moveActiveSheet(3);
  }
  
  // Copy the Income Sheet
  var oldIncomeSheetName = getPreviousYear() + INCOME_SHEET_NAME_SUFFIX;
  var oldIncomeSheet = spreadsheet.getSheetByName(oldIncomeSheetName);
  if (oldIncomeSheet != null) {
    var newIncomeSheet = oldIncomeSheet.copyTo(spreadsheet);
    setupIncomeSheet(newIncomeSheet);
    newIncomeSheet.activate();
    spreadsheet.moveActiveSheet(4);
  }
  
  // Copy the Summary Sheet
  var oldSummarySheetName = getPreviousYear() + SUMMARY_SHEET_NAME_SUFFIX;
  var oldSummarySheet = spreadsheet.getSheetByName(oldSummarySheetName);
  if (oldSummarySheet != null) {
    var newSummarySheet = oldSummarySheet.copyTo(spreadsheet);
    setupSummarySheet(newSummarySheet, '={0}!O31'.format("'" + oldSummarySheetName + "'"));
    newSummarySheet.activate();
    spreadsheet.moveActiveSheet(5);
  }
  
  // Rename the Transactions Sheet
  var transactionSheet = spreadsheet.getSheetByName(TRANSACTION_SHEET_NAME);
  if (transactionSheet != null) {
    var oldTransactionSheetName = getPreviousYear() + ' ' + TRANSACTION_SHEET_NAME;
    transactionSheet.setName(oldTransactionSheetName);
    var newTransactionSheet = transactionSheet.copyTo(spreadsheet);
    newTransactionSheet.setName(TRANSACTION_SHEET_NAME);
    newTransactionSheet.activate();
    spreadsheet.moveActiveSheet(2);
    
    // Clear the previous year's data and data validations
    newTransactionSheet.getRange('G:G').clearDataValidations();
    var data = newTransactionSheet.getRange(2, 1, newTransactionSheet.getLastRow() - 1, newTransactionSheet.getLastColumn());
    data.clearContent();
  }
  
  // Rename the Categories Dropdown Sheet
  var categoriesSheet = spreadsheet.getSheetByName(CATEGORIES_SHEET_NAME);
  if (categoriesSheet != null) {
    var oldCategoriesSheetName = getPreviousYear() + ' ' + CATEGORIES_SHEET_NAME;
    categoriesSheet.setName(oldCategoriesSheetName);
    var newCategoriesSheet = categoriesSheet.copyTo(spreadsheet);
    newCategoriesSheet.setName(CATEGORIES_SHEET_NAME);
    newCategoriesSheet.protect().setWarningOnly(true);
    newCategoriesSheet.hideSheet();
  }
  
  // Fix up all the formulas for the new sheets
  fixCategoryAndSubCategory();
  fixAllSheetFormulas();
  SpreadsheetApp.flush();
}

///////////////////////////////////////////////////////////
// Fix Category & SubCategory dropdown support
///////////////////////////////////////////////////////////

function fixCategoryAndSubCategory() {
  fixCategory();
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var transactionSheet = spreadsheet.getSheetByName(TRANSACTION_SHEET_NAME);
  var transactionRange = transactionSheet.getDataRange();
  fixSubCategory(spreadsheet, transactionSheet, transactionRange, CATEGORIES_SHEET_NAME);
}

function fixCategory() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var categorySheet = spreadsheet.getSheetByName(CATEGORIES_SHEET_NAME);
  if (categorySheet == null) {
    SpreadsheetApp.getUi().alert('Missing ' + CATEGORIES_SHEET_NAME + ' sheet.  Please add it and try again.');
    return;
  }
  var range = categorySheet.getRange(1, 1, 60, 25);
  
  range.clear();
  
  var startRow = 1;
  setCategoryCells(range, startRow, EXPENSES_SHEET_NAME);
  range.getCell(startRow + 27, B_COL).setValue('---');
  range.getCell(startRow + 27, C_COL).setValue('---');
  setCategoryCells(range, startRow + 27 + 1, INCOME_SHEET_NAME);
}

function fixSubCategory(spreadsheet, transactionSheet, transactionRange, categorySheetName) {
  // Map Category names to the row found in the Category sheet
  var categorySheet = spreadsheet.getSheetByName(categorySheetName);
  if (categorySheet == null) {
    return;
  }
  
  var categoryToRow = {};
  var categories = categorySheet.getDataRange().getValues();
  for (var row = 0, lastRow = categories.length; row < lastRow; ++row) {
    var catName = categories[row][1];
    if (catName != undefined && catName.length > 0) {
      categoryToRow[catName] = row + 1;
    }
  }
  
  // Loop over rows in Transactions sheet
  var row = transactionRange.getRow();
  do {
    var subCategoryCell = transactionSheet.getRange(row, TRANSACTION_SUBCAT_COL);
    var categoryRow = categoryToRow[transactionSheet.getRange(row, TRANSACTION_CAT_COL).getValue()];
    if (categoryRow != undefined) {
      var validationRange = categorySheet.getRange(categoryRow, 3, 1, 20);
      var rule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      subCategoryCell.setDataValidation(rule);
    } else {
      subCategoryCell.clearDataValidations();
    }
  } while(row++ < transactionRange.getLastRow());
}

function setCategoryCells(range, startRow, refSheetName) {
  refSheetName = "'" + refSheetName + "'";
  var NUM_ROWS = 25;
  
  // Update "Row #" row
  var curRow = startRow;
  range.getCell(curRow, A_COL).setValue('Row #');
  range.getCell(curRow, B_COL).setFormula('=uniQue({0}!A:A)'.format(refSheetName));
  range.getCell(++curRow, A_COL)
  .setFormula('=ArrayFormula(iferror(match(B{0}:B{1},{2}!$A:$A,0)))'.format(curRow, curRow + 25, refSheetName))
  .setNote('This formula matches the categories in column C with their locations in the {0} tab, then displays the resulting row number. The formula in column D uses it to calculate the values in columns D:Q.'.format(refSheetName));
  
  for (var row = curRow, l = row + 25; row <= l; ++row) {
    var formula = '=if(not(isblank(A{0})), Transpose(indirect("{2}!C"&A{0}+1&":C"&if(isblank(A{1}), A{0}+10, A{1}-1))),"")'.format(row, row + 1, refSheetName);
    range.getCell(row, C_COL).setFormula(formula);
  }
}


///////////////////////////////////////////////////////////
// Fix Formulas in all sheets
///////////////////////////////////////////////////////////

function fixAllSheetFormulas() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var expensesSheet = spreadsheet.getSheetByName(EXPENSES_SHEET_NAME);
  if (expensesSheet == null) {
    SpreadsheetApp.getUi().alert('Missing ' + EXPENSES_SHEET_NAME + ' sheet.  Please add it and try again.');
  } else {
    fixSheetFormulas(expensesSheet, true);
  }
  
  var incomeSheet = spreadsheet.getSheetByName(INCOME_SHEET_NAME);
  if (incomeSheet == null) {
    SpreadsheetApp.getUi().alert('Missing ' + INCOME_SHEET_NAME + ' sheet.  Please add it and try again.');
  } else {
    fixSheetFormulas(incomeSheet, false);
  }
}

function fixSheetFormulas(sheet, expenses) {
  // FIXME- Remove this write-repair when all old spreadsheets have been updated
  if (sheet.getRange('C3').getValue() == 'Transaction Sheet') {
    sheet.deleteRow(3);
  }
  
  var categories = findCategoryRows(sheet);
  
  categories.forEach(function(category) {
    updateCategoryRowFormulas(sheet, category);
    category.subCategories.forEach(function(subcategory) {
      updateSubCategoryRowFormulas(sheet, subcategory, expenses);
    });
  });
}

function findCategoryRows(sheet) {
  var categories = [];
  
  var values = sheet.getDataRange().getValues();
  var curCategory = null;
  for (var rowIdx = 0; rowIdx < values.length; ++rowIdx) {
    var row = values[rowIdx];
    var rowNum = rowIdx + 1;
    if (row[0] != null && row[0].length > 0) {
      // Found new Category
      curCategory = {name: row[0], rowNum: rowNum, endRowNum: rowNum, subCategories: []};
      categories.push(curCategory);
      continue;
    }
    
    if (curCategory != null) {
      if (row[2] != null && row[2].length > 0) {
        // Found new SubCategory
        var subCategory = {name: row[2], rowNum: rowNum, categoryRowNum: curCategory.rowNum};
        curCategory.subCategories.push(subCategory);
      }
      
      curCategory.endRowNum = rowNum
    }
  }
  
  return categories;
}

function updateCategoryRowFormulas(sheet, category) {
  var range = sheet.getDataRange();
  
  if (range.getCell(category.rowNum, CATEGORY_COL).getValue() != category.name) {
    throw "Category name does not match cell: " + category.name;
  }
  
  range.getCell(category.rowNum, MONTHLY_TOTALS_COL).setValue('Monthly Totals:');
  
  for (var col = JAN_COL; col <= BUDGET_COL; ++col) {
    var colName = colNumToName(col);
    
    var formula = '=SuM({0}{1}:{0}{2})'.format(colName, category.rowNum + 1,category.endRowNum);
    range.getCell(category.rowNum, col).setFormula(formula);
  }
}

function updateSubCategoryRowFormulas(sheet, subcategory, expenses) {
  var range = sheet.getDataRange();
  
  if (range.getCell(subcategory.rowNum, CATEGORY_COL).getValue() != '') {
    throw "Category name column must be blank on SubCategory row: " + subcategory.name;
  }
  if (range.getCell(subcategory.rowNum, SUB_CATEGORY_COL).getValue() != subcategory.name) {
    throw "SubCategory name does not match cell: " + subcategory.name;
  }
  
  // Fix Month formulas
  for (var col = JAN_COL; col <= DEC_COL; ++col) {
    var colName = colNumToName(col);
    
    var sign = (expenses ? '-' : '');
    var transAmtRange = 'Transactions!$E:$E';
    var transCategoryRange = 'Transactions!$F:$F';
    var categoryCell = '$A$' + subcategory.categoryRowNum;
    var transSubCategoryRange = 'Transactions!$G:$G';
    var subCategoryCell = '$C' + subcategory.rowNum;
    var transDateRange = 'Transactions!$B:$B';
    var curColDateCell = colName + '$' + DATE_ROW;
    var nextColDateCell = colNumToName(col + 1) + '$' + DATE_ROW;
    if (col == DEC_COL) {
      nextColDateCell = curColDateCell + '+31';
    }
    var formula = 
        '={0}SuMIFS({1}, {2}, {3}, {4}, {5}, {6}, ">="&{7}, {6}, "<"&{8})'
        .format(sign, transAmtRange, transCategoryRange, categoryCell, transSubCategoryRange, subCategoryCell, transDateRange, curColDateCell, nextColDateCell);
    range.getCell(subcategory.rowNum, col).setFormula(formula);
  }
  
  // Fix Total formula
  var totalFormula = '=SuM(D{0}:O{0})'.format(subcategory.rowNum);
  range.getCell(subcategory.rowNum, TOTAL_COL).setFormula(totalFormula);
  
  // Fix Average formula
  updateAverageFormula(range, subcategory);
}

function updateAverageFormula(range, subcategory) {
  var avgFormula = '=IFerror(averageifs(D{0}:O{0}, $D${1}:$O${1}, ">="&StartingDate, $D${1}:$O${1}, "<="&TODAY()))'.format(subcategory.rowNum, DATE_ROW);
  range.getCell(subcategory.rowNum, AVG_COL).setFormula(avgFormula);
}

function colNameToNum(colName) {
  return colName.charCodeAt(0) - A_ASCII_CODE + 1;
}

function colNumToName(colNum) {
  return String.fromCharCode(A_ASCII_CODE + colNum - 1);
}


///////////////////////////////////////////////////////////
// "Enter Transactions" menu handling
///////////////////////////////////////////////////////////

function openEnterTransactionsDialog() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(TRANSACTION_SHEET_NAME);
  if (sheet == null) {
    SpreadsheetApp.getUi().alert('Missing ' + TRANSACTION_SHEET_NAME + ' sheet.  Please add it and try again.');
    return;
  }
  
  //Call the HTML file and set the width and height
  var html = HtmlService.createHtmlOutputFromFile("EnterTrans")
    .setWidth(650)
    .setHeight(300);
  
  //Display the dialog
  var dialog = SpreadsheetApp.getUi().showModalDialog(html, "Enter new Transactions");
}

/**
 * Callback from the dialog.
 */
function enterTransactions(transactionData) {
  // Display the values submitted from the dialog box in the Logger. 
  //Logger.log(transactionData.transSource);
  //Logger.log('suppressDups: ' + transactionData.suppressDups);
  var isChecking = transactionData.transSource == 'checking';
  //Logger.log('isChecking: ' + isChecking);
  var values = transactionData.values;
  //Logger.log(values);
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(TRANSACTION_SHEET_NAME);
  var currentRow = sheet.getLastRow() + 1;
  
  var existingTransactions = {};
  if (transactionData.suppressDups) {
    var existingData = sheet.getDataRange().getValues();
    for (var row = 0; row < existingData.length; ++row) {
      var key = createKey(existingData[row][1], existingData[row][3], existingData[row][4]);
      //Logger.log('existing Key: ' + key + ' --> ' + existingData[row][1]);
      existingTransactions[key] = existingData[row];
    }
  }
  
  var suppressedTransactions = [];
  for (var idx = 0; idx < values.length; ++idx) {
    var entry = values[idx];
    entry.type = entry.type.trim();
    entry.desc = entry.desc.trim();
    var range = sheet.getRange(currentRow + ':' + currentRow);
    var amt = fixAmount(entry.amt, isChecking);
    var key = createKey(entry.transDate, entry.desc, amt);
    //Logger.log('Incoming Key: ' + key);
    if (existingTransactions.hasOwnProperty(key)) {
      //Logger.log('Found duplicate row: ' + key);
      suppressedTransactions.push(entry);
      continue;
    }
    range.getCell(1, 1).setValue(entry.type);
    range.getCell(1, 2).setValue(entry.transDate);
    range.getCell(1, 3).setValue(entry.postDate);
    range.getCell(1, 4).setValue(entry.desc);
    range.getCell(1, 5).setValue(amt);
    ++currentRow;
  }
  
  if (suppressedTransactions.length > 0) {
    var details = '<u>These transactions appeared to be duplicates</u>:';
    details += '<table>'
    for (var i = 0; i < suppressedTransactions.length; ++i) {
      var tran = suppressedTransactions[i];
      details += '<tr><td>' + escapeHtml(tran.transDate) + '</td><td>' + escapeHtml(tran.desc) + '</td><td><strong>' + escapeHtml(tran.amt) + '</strong></td></tr>';
    }
    details += '</table>';
    
    var html =   '<html><head><link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">'
               + details
               + '<input type="button" value="OK" onClick="google.script.host.close();" />';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), 'Transactions NOT added');
    //SpreadsheetApp.getUi().alert(details);
  }
}

function createKey(transDate, desc, amt) {
  transDate = new Date(transDate);
  desc = desc.toLowerCase().replace(/[^a-zA-Z]/g, '').trim();
  amt = ('' + amt).replace(/\.00/, ''); // Remove trailing '.00'
  amt = amt.replace(/\.[0-9]0/, function(val) {return val.substring(0, val.length - 1)}); // Remove trailing '0'
  if (amt.endsWith('.')) {
    amt = amt.substring(0, amt.length - 1);
  }
  amt = amt.replace(',', '');
  return transDate.getUTCMonth() + '-' + transDate.getUTCDate() + '-' + transDate.getUTCFullYear() + '#' + desc + '#' + amt;
}

function fixAmount(amt, isChecking) {
  if (!amt) {
    return amt;
  }
  amt = amt.trim();
  if (amt.length == 0) {
    return amt;
  }
  amt = amt.replace('$', '');
  if (amt.charCodeAt(0) == 8722 /* unicode minus */) {
    amt = '-' + amt.substring(1);
  }
  else if (amt.charCodeAt(0) == 10133 /* unicode plus */) {
    amt = '+' + amt.substring(1);
  }
  
  if (!isChecking) {
    if (amt.charAt(0) == '-') {
      amt = amt.substr(1);
    } else if (amt.charAt(0) == '+') {
      amt = '-' + amt.substr(1);
    } else {
      amt = '-' + amt;
    }
  }
  
  return amt;
}


///////////////////////////////////////////////////////////
// Utility functions
///////////////////////////////////////////////////////////

function escapeHtml(text) {
  if (text == null) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function getCurrentYear() {
  return new Date().getFullYear();
}

function getPreviousYear() {
  return new Date().getFullYear() - 1;
}

if (!String.prototype.format) {
  String.prototype.format = function() {
    var args = arguments;
    return this.replace(/{(\d+)}/g, function(match, number) { 
      return typeof args[number] != 'undefined'
        ? args[number]
        : match
      ;
    });
  };
}

if (!String.prototype.startsWith) {
  String.prototype.startsWith = function(prefix) {
      return this.indexOf(prefix) === 0;
  }
}

if (!String.prototype.endsWith) {
  String.prototype.endsWith = function(suffix) {
    return this.match(suffix+"$") == suffix;
  }
}
