var ctx = new Excel.RequestContext();
var sheet = ctx.workbook.worksheets.getItem("Sheet1");

var categoriesRange = sheet.getRange("A2:A5");
categoriesRange.values = [
	[ "Category 1" ], 
	[ "Category 2" ], 
	[ "Category 3" ], 
	[ "Category 4" ]
];
  
var seriesRange = sheet.getRange("B1:D1");
seriesRange.values = [ 
	[ "Series 1", "Series 2", "Series 3", ] 
];  

var dataRange = sheet.getRange("B2:D5");
dataRange.formulas = "=RAND()*17";
dataRange.numberFormat = "#0";

ctx.executeAsync().then();