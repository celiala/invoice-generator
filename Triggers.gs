var COLUMN_COMPANY = 5;
var COLUMN_EVENT = 8;
var COLUMN_ITEM = 10;

var invoiceSheetTriggers = new Triggers();
invoiceSheetTriggers.addTrigger(COLUMN_EVENT, populateEventDependentDropdowns);
invoiceSheetTriggers.addTrigger(COLUMN_ITEM, populateItemDependentCells);
invoiceSheetTriggers.addTrigger(COLUMN_COMPANY, populateCompanyDependentCells);

function Triggers() {
  this.triggers = {};
  this.addTrigger = function(col, fxn) {
    this.triggers[col.toString()] = function(a,b,c,d,e) { fxn(a,b,c,d,e); }
  };
  this.onEdit = function(e) {
    var triggerKey = e.range.getColumn().toString();
    Logger.log(triggerKey);
    if (triggerKey in this.triggers) {
      (this.triggers[triggerKey])(e.range.getRow());
    }
  };
};
