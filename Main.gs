var ROW_HEADER = 1;
var ROW_DATA_START = 2;

function onOpen() {
  populateColumnDropdowns();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wscMenu = [ 
    {name: "Refresh Dropdowns", functionName: "populateColumnDropdowns"},
    {name: "Generate Invoices", functionName: "generateInvoices"}
  ];
  ss.addMenu("Run...", wscMenu);
}

function onEdit(e){
  invoiceSheetTriggers.onEdit(e);
}

