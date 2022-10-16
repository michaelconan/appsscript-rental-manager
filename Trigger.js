/**
 * Edit trigger to filldown date and lookup formula columns when speadsheet updated directly
 */
function onEdit(e) {
  var rg = e.range;
  var sheet = rg.getSheet();
  // Filldown detail formulas
  if (sheet.getName() == 'Transaction Detail') {
    if (sheet.getRange(rg.getRow(), 4, 1, 7).getValues()[0].filter(r => r != '').length) {
      sheet.getRange(rg.getRow()-1, 1, 1, 3).autoFill(sheet.getRange(rg.getRow()-1, 1, 2, 3), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      sheet.getRange(rg.getRow()-1, 11).autoFill(sheet.getRange(rg.getRow()-1, 11, 2), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    } else {
      sheet.getRange(rg.getRow(), 1, 1, 11).clear();
    }
  }
}

/**
 * Submission trigger of expense form to post entry
 */
function onFormSubmit(e) {
    
  // Get account and form header mappings
  var accounts = CRM.getSheetByName('Accounts').getDataRange().getValues();
  var headers = CRM.getSheetByName("Expenses").getDataRange().getValues().slice(0,2);
  Logger.log('Read mapping data');

  // Get column headers and number of rows from ledger
  var data = LEDGER.getDataRange();
  var columns = data.getValues().slice(0)[0];
  var row = data.getNumRows() + 1;
  Logger.log('Read transaction data');

  // Create new object of form values based on mapping
  var formVals = {};
  for (let i=0; i< headers[0].length; i++) {
    if (headers[0][i]) {
      formVals[headers[0][i]] = e.namedValues[headers[1][i]]
    }
  }
  Logger.log('Mapped form values to transaction detail columns');

  // Identify cell to populate each form value, convert as needed
  for (let p in formVals) {
    Logger.log(p);
    var col = columns.indexOf(p) + 1;
    if (col) {
      var value = formVals[p];
      if (p == 'Account Name') {
        value = accounts.filter(r => r[2] == value)[0][1];
        col = columns.indexOf('Account') + 1;
      } else if (p == 'Amount' && formVals['Account Name'].indexOf('Income') == -1) {
        value = parseFloat(value) * -1;
      }

      LEDGER.getRange(row, col).setValue(value);
      Logger.log('Populated '+p+' value');
    }
  }

  // Trigger autofill functionality
  let evt = {
    range: LEDGER.getRange(row, 4)
  };
  onEdit(evt);

}