// After it is working, check video at 1:40:41
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Custom Menu')
        .addItem('Execute', 'execute')
        .addToUi();
}
function execute() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('Main Sheet');
    var resultSheet = ss.getSheetByName('Results');
    var url = mainSheet.getRange('B1').getValue();
    if (!url) {
        ss.toast('No API endpoint specified', 'Error', 5);
    }
    else {
        var response = UrlFetchApp.fetch(url);
        var str = response.getContentText();
        var obj = JSON.parse(str);
        var matrix = [];
        for (var _i = 0, _a = obj.data.objects; _i < _a.length; _i++) {
            var elem = _a[_i];
            matrix.push([elem.date, elem.problem, elem.user, elem.language, elem.result, elem.points]);
        }
        var numRows = matrix.length;
        var numCols = matrix[0].length;
        resultSheet.getRange(2, 1, numRows, numCols).setValues(matrix);
    }
}
