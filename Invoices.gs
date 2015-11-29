/**
 * Generates invoices for rows that don't yet have a value in the COLUMN_PDF_URL column.
 */
function generateInvoices() {
  // Get the active spreadsheet and the active sheet
  var sheet = getSheet(SHEET_INVOICES);

  for (var r = ROW_DATA_START; r <= sheet.getLastRow(); r++) {
    if (sheet.getRange(r, COLUMN_INVOICE_NO).getValue() == "") {
      //no more rows to process
      break;
    }
    if (sheet.getRange(r, COLUMN_PDF_URL).getValue() == "") {
      var invoiceValues = getInvoiceValues(sheet, r);
      var companyName = invoiceValues[HEADER_COMPANY];
      var fileId = createNewInvoice(companyName);
      fillInInvoiceTemplate(fileId, invoiceValues);
      sheet.getRange(r, COLUMN_DOC_URL).setValue(String(URL_DOC).replace("<<file_id>>",fileId))
      sheet.getRange(r, COLUMN_PDF_URL).setValue(String(URL_PDF).replace("<<file_id>>",fileId))
    }
  }
}

function testInvoiceValues() {
  var sheet = getSheet(SHEET_INVOICES);
  var r = 2;
  var values = getInvoiceValues(sheet, r);
  Logger.log(values);
}

function getInvoiceValues(sheet, r) {
  var invoiceDict = new InvoiceDictionary();
  invoiceDict.addValues(sheet, r);
  var company = invoiceDict.get(HEADER_COMPANY);
  invoiceDict.extractValuesFromSheet(SHEET_COMPANIES, COL_COMPANIES_COMPANY, company);
  var event = invoiceDict.get(HEADER_EVENT);
  invoiceDict.extractValuesFromSheet(SHEET_EVENTS, COL_EVENTS_EVENT, event);
  
  //special formatting here
  var invoiceDate = invoiceDict.values[HEADER_INVOICE_DATE];
  var date = new Date(invoiceDate);
  invoiceDict.values[HEADER_INVOICE_DATE] = date.getMonth()+1 + "/" + date.getDate() + "/" + date.getYear();
  
  return invoiceDict.values;
}


function InvoiceDictionary() {
  this.values = {};
  
  this.addValues = function(sheet, r) {
    for (var c = 1; c <= sheet.getLastColumn(); c++) {
      var key = sheet.getRange(ROW_HEADER, c).getValue();
      var value = sheet.getRange(r, c).getValue();
      this.values[key] = value;
    }
    return this;
  };

  this.extractValuesFromSheet = function(sheetName, filterCol, filterValue) {
    var rows = (new Autofill())
    .setSource(sheetName, filterCol)
    .addFilter(filterCol, filterValue)
    .getSourceRows();
    if (rows.length > 0) {
      this.addValues(getSheet(sheetName), rows[0]);
    }
  }
  
  this.get = function(key) {
    return this.values[key];
  }
}

function fillInInvoiceTemplate(fileId, values) {
  var doc = DocumentApp.openById(fileId);
  
  // Use editAsText to obtain a single text element containing
  // all the characters in the document.
  var text = doc.getBody().editAsText();
  
  for(var key in values) {
    text.replaceText("<<" + key + ">>", values[key]);
  }
}

function createNewInvoice(companyName) {
  var template = DriveApp.getFileById(TEMPLATE_ID);
  var templateFolder = template.getParents().next();
  var existingCompanyFolders = templateFolder.getFoldersByName(companyName);
  var companyFolder = null;
  if (existingCompanyFolders.hasNext()) {
    companyFolder = existingCompanyFolders.next();
  } else {
    companyFolder = templateFolder.createFolder(companyName);
  }
  
  var d = new Date();
  var filename = "Invoice " + d.getYear() + "-" + (d.getMonth() + 1);
  var file = template.makeCopy(filename, companyFolder);
  return file.getId();
}
