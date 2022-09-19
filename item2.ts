// After it is working, check video at 1:40:41

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Fetch Data', 'execute')
      .addItem('Clear Results', 'clearRes')
      .addToUi();
}

function clearRes() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Clear Results', 'Are you sure you want to clear the results?', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.NO) {
    return;
  } 

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('Results');
  
  const clearOptions = {
    contentsOnly: true
  };

  resultSheet.getRange('A2:F2').clear(clearOptions);
  if(resultSheet.getMaxRows() > 2) {
    resultSheet.deleteRows(3, resultSheet.getMaxRows()-2)
  }
}

function execute() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName('Main Sheet');
  const resultSheet = ss.getSheetByName('Results');

  var url = mainSheet.getRange('B1').getValue();
  var user = mainSheet.getRange('B2').getValue();
  var lang = mainSheet.getRange('B3').getValue();
  var prob = mainSheet.getRange('B4').getValue();


  if (!url) {
    ss.toast('No API endpoint specified', 'Error', 5);
  } 
  else
  {
    const clearOptions = {
      contentsOnly: true
    };
    
    resultSheet.getRange('A2:F2').clear(clearOptions);
    if(resultSheet.getMaxRows() > 2) {
      resultSheet.deleteRows(3, resultSheet.getMaxRows()-2)
    }

    if(user || lang || prob) {
      url += "?";
    }
  
    if(user) {
      url += "user=" + user;
      if(lang || prob) {
        url += '&';
      }
    }
  
    if(lang) {
      url += "language=" + lang;
      if(prob) {
        url += '&';
      }
    }
  
    if(prob) {
      url += "problem=" + prob;
    }
  
    const response = UrlFetchApp.fetch(url);
    const str = response.getContentText();
    const obj = JSON.parse(str);
  
    const matrix: string[][] = [];
  
    for(let elem of obj.data.objects) {
      matrix.push([elem.date, elem.problem, elem.user, elem.language, elem.result, elem.points]);
    }
    if (matrix.length > 0) {
      const numRows = matrix.length;
      const numCols = matrix[0].length;
      resultSheet.getRange(2, 1, numRows, numCols).setValues(matrix);
    }
    var res = 'Results found: ' + matrix.length;
    ss.toast(res, 'Results', 5);
  }
}