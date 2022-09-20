
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Fetch Data', 'execute')                                                                               // Fetch Data
      .addItem('Clear Results', 'clearRes')                                                                           // Clear Results
      .addToUi();
}

function clearRes() {                                                                                                 // Clear Results implementation
  var ui = SpreadsheetApp.getUi();  
  var response = ui.alert('Clear Results', 'Are you sure you want to clear the results?', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.NO) {                                                                                     // Exit function if choice is No
    return;
  } 

  let ss = SpreadsheetApp.getActiveSpreadsheet();                                                                       
  const resultSheet = ss.getSheetByName('Results')                                                                    
  
  const clearOptions = {                                                                                               // Preserve Formatting
    contentsOnly: true
  };

  resultSheet.getRange('A2:F2').clear(clearOptions);                                                                   // Clears second row
  if(resultSheet.getMaxRows() > 2) {                                                                                   // Checks if is a third row so we can delete 
    resultSheet.deleteRows(3, resultSheet.getMaxRows()-2)
  }
}

function execute() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName('Main Sheet');
  const resultSheet = ss.getSheetByName('Results');

  var url = mainSheet.getRange('B1').getValue();                                                                       // Extracting API endpoint, user, language, 
  var user = mainSheet.getRange('B2').getValue();                                                                      // and problem fields from Main Sheet
  var lang = mainSheet.getRange('B3').getValue();
  var prob = mainSheet.getRange('B4').getValue();


  if (!url) {                                                                                                          // Toast for empty API endpoint field
    ss.toast('No API endpoint specified', 'Error', 5);
  } 
  else
  {
    const clearOptions = {                                                                                             // Preserve Formatting
      contentsOnly: true
    };
    
    resultSheet.getRange('A2:F2').clear(clearOptions);                                                                 // Automatically clears field every time we
    if(resultSheet.getMaxRows() > 2) {                                                                                 // fetch data
      resultSheet.deleteRows(3, resultSheet.getMaxRows()-2)
    }

    if(user || lang || prob) {                                                                                         // Appends '?' in the API endpoint
      url += "?";                                                                                                      // if there is a criteria in user,
    }                                                                                                                  // language, and problem field
  
    if(user) {                                                                                                         // appends data in user to url
      url += "user=" + user;
      if(lang || prob) {                                                                                               // appends '&' if there is a language
        url += '&';                                                                                                    // or problem criteria
      }                                         
    }
  
    if(lang) {                                                                                                         // appends data in language to url
      url += "language=" + lang;
      if(prob) {                                                                                                       // appends '&' if there is a problem criteria
        url += '&';                                                                                                    
      }
    }
  
    if(prob) {                                                                                                         // appends data in problem to url
      url += "problem=" + prob;
    }
  
    var response = UrlFetchApp.fetch(url);                                                                             // fetching data from url
    var str = response.getContentText();
    var obj = JSON.parse(str);                                                                                         // parsing str data
  
    const matrix: string[][] = [];                                                                                    
  
    for(let elem of obj.data.objects) {
      matrix.push([elem.date, elem.problem, elem.user, elem.language, elem.result, elem.points]);                      // generating of matrix
    }
    if (matrix.length > 0) {                                                                                           // populating Results sheet with rows
      const numRows = matrix.length;                                                                                   // based on the size of matrix
      const numCols = matrix[0].length;
      resultSheet.getRange(2, 1, numRows, numCols).setValues(matrix);
    }

    if(user || lang || prob) {                                                                                         // Pagination
      url += "&page=2";                                                                                                // If there is a user/lang/prob
    }                                                                                                                  // criteria, no need to add '?'
    else                                                                                                               // otherwise, add '?'
    {
      url += "?page=2";
    }

    if(UrlFetchApp.fetch(url, {muteHttpExceptions: true}).getResponseCode() == 200)                                    // attempt to add data from next page
    {                                                                                                                  // if the response code is 200 (success),
      response = UrlFetchApp.fetch(url);                                                                               // populate matrix with the data on the second page
      str = response.getContentText();
      obj = JSON.parse(str);
      
      for(let elem of obj.data.objects) {
        matrix.push([elem.date, elem.problem, elem.user, elem.language, elem.result, elem.points]);
      }
      if (matrix.length > 0) {
        const numRows = matrix.length;
        const numCols = matrix[0].length;
        resultSheet.getRange(2, 1, numRows, numCols).setValues(matrix);
      }  
      
    }

    var res = 'Submissions retrieved: ' + matrix.length;                                                                // Tally of submissions retrieved
    ss.toast(res, 'Results', 5);
  }
}