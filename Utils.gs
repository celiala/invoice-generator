function wholeCol(col) {
  return col + ":" + col;
}

function colData(col, lastRow) {
  return col + "2:" + col + lastRow;
}

function a1Cell(a1Col, row) {
  return a1Col + row;
}

function getRange(sheetName, a1Notation) {
  return getSheet(sheetName).getRange(a1Notation);
}

function getSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

function getLastRow(sheetName) {
  return getSheet(sheetName).getLastRow();
}

function getNextInvoiceNo(prevInvoiceNo) {
  var nextNum = parseInt(prevInvoiceNo, 10)+1;
  return padNumber(nextNum, prevInvoiceNo.length);
}

function padNumber(num, numPadding) {
  var str = num.toString();
  while (str.length < numPadding) {
    str = "0" + str;
  }
  return str;
}

function testDate() {
  var d = new Date();
  Logger.log(d.getYear() + " " + d.getMonth());
  //var dateParts = e.parameter.dataFromField.split(',');
  Logger.log(d.toString('yyyy-MM'));
}

function testArray() {
  var rows = [1,2,3,6];
  for(var i in rows) {
    Logger.log(rows[i]);
  }
}
/*
var foo = function() {
  this.bar = null;
  this.a = null;
  this.b = null;
  
  this.setA = function(a) {
    this.a = a;
    return this;
  }

  this.setB = function(b) {
    this.b = b;
    return this;
  }
  
  this.doIt = function() {
    Logger.log(this.bar + "|" + this.a + "|" + this.b);
  }
}

function testFoo() {
  var f = new foo();
  f.bar = 3;
  f.setA('a').setB('b');
  f.doIt();
  f = new foo();
  f.doIt();
}


var INVOICES = new function() {
  this.name = "invoices";
  this.COMPANY = "E";
  this.EVENT = "H";
  this.EVENTDATE = "I";
  this.ITEM = "J";
  this.getRangeColumn = function(a1Col) {
    return getRange(this.name, a1Col + ":" + a1Col);
  }
  this.getRangeCell = function(a1Col, row) {
    return getRange(this.name, a1Col + row);
  }
}
*/
