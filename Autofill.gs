function test() {
  populateEventDependentDropdowns(3);
}

function populateColumnDropdowns() {
  var lastRow = getLastRow(SHEET_INVOICES) + 5;
  (new Autofill())
    .setDestination(SHEET_INVOICES, colData(INVOICES_COMPANY, lastRow))
    .setSource(SHEET_COMPANIES, 1)
    .buildDropdown();
  (new Autofill())
    .setDestination(SHEET_INVOICES, colData(INVOICES_EVENT, lastRow))
    .setSource(SHEET_EVENTS, 1)
    .buildDropdown();
  var lastRow = getLastRow(SHEET_ITEMS) + 5;
  (new Autofill())
    .setDestination(SHEET_ITEMS, colData(ITEMS_EVENT, lastRow))
    .setSource(SHEET_EVENTS, 1)
    .buildDropdown();
}

function populateEventDependentDropdowns(row) {
  var event = getRange(SHEET_INVOICES, a1Cell(INVOICES_EVENT, row)).getValue();
  (new Autofill())
    .setDestination(SHEET_INVOICES, a1Cell(INVOICES_EVENTDATE, row))
    .setSource(SHEET_EVENTS, COL_EVENTS_EVENTDATE)
    .addFilter(COL_EVENTS_EVENT, event)
    .buildDropdown(true);

  (new Autofill())
    .setDestination(SHEET_INVOICES, a1Cell(INVOICES_ITEM, row))
    .setSource(SHEET_ITEMS, COL_ITEMS_ITEM)
    .addFilter(COL_ITEMS_EVENT, event)
    .buildDropdown(true);
  
  populateItemDependentCells(row);
}

function populateItemDependentCells(row) {
  var event = getRange(SHEET_INVOICES, a1Cell(INVOICES_EVENT, row)).getValue();
  var item = getRange(SHEET_INVOICES, a1Cell(INVOICES_ITEM, row)).getValue();
  (new Autofill())
    .setDestination(SHEET_INVOICES, a1Cell(INVOICES_PRICE, row))
    .setSource(SHEET_ITEMS, COL_ITEMS_PRICE)
    .addFilter(COL_ITEMS_EVENT, event)
    .addFilter(COL_ITEMS_ITEM, item)
    .buildDefaultValue();
  var price = getRange(SHEET_INVOICES, a1Cell(INVOICES_PRICE, row)).getValue();
  var cellQuantity = getRange(SHEET_INVOICES, a1Cell(INVOICES_QUANTITY, row));
  if (cellQuantity.getValue() == "") {
    cellQuantity.setValue(price != "" ? 1 : "");
  }
}

function populateCompanyDependentCells(row) {
  var company = getRange(SHEET_INVOICES, a1Cell(INVOICES_COMPANY, row)).getValue();
  if (company == "") {
    return;
  }
  var cellInvoiceNo = getRange(SHEET_INVOICES, a1Cell(INVOICES_INVOICE_NO, row));
  if (cellInvoiceNo.getValue() == "" && row > 2) {
    var prevInvoiceNo = getRange(SHEET_INVOICES, a1Cell(INVOICES_INVOICE_NO, row-1)).getValue();
    if (prevInvoiceNo != "") {
      var nextInvoiceNo = getNextInvoiceNo(prevInvoiceNo);
      cellInvoiceNo.setValue(nextInvoiceNo);
    }
  }
  
  var cellInvoiceDate = getRange(SHEET_INVOICES, a1Cell(INVOICES_INVOICE_DATE, row));
  if (cellInvoiceDate.getValue() == "") {
    var now = new Date();
    cellInvoiceDate.setValue((new Date(now.getYear(), now.getMonth(), now.getDate())));
  }
}

/**
 * Set dropdown values for , using values sheetNameValues.colValues
 */
var Autofill = function() {
  this.destSheet = null;
  this.destRange = null;
  this.sourceSheet = null;
  this.sourceCol = null;

  this.filters = {};
  
  this.setDestination = function(sheet, range) {
    this.destSheet = sheet;
    this.destRange = range;
    return this;
  };
  
  this.setSource = function(sheet, col) {
    this.sourceSheet = sheet;
    this.sourceCol = col;
    return this;
  };
  
  this.addFilter = function(col, value) {
    this.filters[col] = value;
    return this;
  };
  
  this.buildDefaultValue = function() {
    var range = getRange(this.destSheet, this.destRange);
    var values = this.getSourceValues();
    if (values.length) {
        range.setValue(values[0]);
    } else {
      range.setValue("");
    }
  };
  
  this.buildDropdown = function(setFirstValue) {
    var range = getRange(this.destSheet, this.destRange);
    var values = this.getSourceValues();
    if (values.length) {
      var rule = this.createRule(values);
      range.setDataValidation(rule);
      if (setFirstValue) {
        range.setValue(values[0]);
      }
    } else {
      range.clearDataValidations();
      range.setValue("");
    }
  };

  this.createRule = function(values) {
      return SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(true)
      //.setHelpText('value not found on "' + this.sourceSheet + '" sheet')
      .build();
  }
  
  this.getSourceValues = function () {
    var sheet = getSheet(this.sourceSheet);
    var rows = this.getSourceRows();
    var values = [];
    for (var i in rows) {
      var value = sheet.getRange(rows[i], this.sourceCol).getValue();
      if (value == "") {
        //no more rows to process
        break;
      }
      // don't add duplicates
      if (value in values) {
        continue;
      }
      values.push(value);
    }
    return values;
  };
  
  this.getSourceRows = function() {
    var sheet = getSheet(this.sourceSheet);
    var rows = [];
    for (var r = ROW_DATA_START; r <= sheet.getLastRow(); r++) {
      if (sheet.getRange(r, this.sourceCol).getValue() == "") {
        //no more rows to process
        break;
      }
      // are we filtering results by cell[row, col] == filterValue?
      if (!this.passesFilters(sheet, r)) {
        continue;
      }
      rows.push(r);
    }
    return rows;  
  }
  
  this.passesFilters = function(sheet, row) {
    for (var filterCol in this.filters) {
      var expectedValue = this.filters[filterCol];
      var actualValue = sheet.getRange(row, filterCol).getValue();
      if (actualValue != expectedValue) {
        return false;
      }
    }
    return true;
  }
}

