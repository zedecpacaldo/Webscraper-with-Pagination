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
    var user = mainSheet.getRange('B2').getValue();
    var lang = mainSheet.getRange('B3').getValue();
    var prob = mainSheet.getRange('B4').getValue();
    if (!url) {
        ss.toast('No API endpoint specified', 'Error', 5);
    }
    else {
        var clearOptions = {
            contentsOnly: true
        };
        if (user || lang || prob) {
            url += "?";
        }
        if (user) {
            url += "user=" + user;
            if (lang || prob) {
                url += '&';
            }
        }
        if (lang) {
            url += "language=" + lang;
            if (prob) {
                url += '&';
            }
        }
        if (prob) {
            url += "problem=" + prob;
        }
        resultSheet.getRange('A2:F2').clear(clearOptions);
        if (resultSheet.getMaxRows() > 2) {
            resultSheet.deleteRows(3, resultSheet.getMaxRows() - 2);
        }
        var response = UrlFetchApp.fetch(url);
        var str = response.getContentText();
        var obj = JSON.parse(str);
        var matrix = [];
        for (var _i = 0, _a = obj.data.objects; _i < _a.length; _i++) {
            var elem = _a[_i];
            matrix.push([elem.date, elem.problem, elem.user, elem.language, elem.result, elem.points]);
        }
        if (matrix.length > 0) {
            var numRows = matrix.length;
            var numCols = matrix[0].length;
            resultSheet.getRange(2, 1, numRows, numCols).setValues(matrix);
        }
    }
}
