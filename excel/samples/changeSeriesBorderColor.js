var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	
chart.series.getItemAt(0).lineFormat.color = "#FF0000";
ctx.executeAsync().then();