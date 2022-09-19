// After it is working, check video at 1:40:41

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Execute', 'execute')
      .addToUi();
}

function execute() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName('Main Sheet');
  const resultSheet = ss.getSheetByName('Results');

  var url = mainSheet.getRange('B1').getValue();

  if (!url) {
    ss.toast('No API endpoint specified', 'Error', 5);
  } 
  else
  {
    const response = UrlFetchApp.fetch(url);
    const str = response.getContentText();
    const obj = JSON.parse(str);
  
    const matrix: string[][] = [];
  
    for(let elem of obj.data.objects) {
      matrix.push([elem.date, elem.problem, elem.user, elem.language, elem.result, elem.points]);
    }
  
    const numRows = matrix.length;
    const numCols = matrix[0].length;
    resultSheet.getRange(2, 1, numRows, numCols).setValues(matrix);
  }

  

}