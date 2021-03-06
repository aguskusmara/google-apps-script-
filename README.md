### move data to another sheet
this is function to copy and delete data from 1 sheet to destination sheet

copy this file gs [movedata.gs](https://github.com/aguskusmara/google-apps-script-/blob/main/movedata.gs)


```
//move data
function moveData() {
  const sheet = SpreadsheetApp.getActive()
  const wsFrom = sheet.getSheetByName('PUT SHEET NAME FROM') //Enter the sheet name that will be moved
  const lastRow = getLastRowSpecial(wsFrom.getRange('put range').getValues()) //see helper function
  const dataFrom = wsFrom.getRange(2, 1, lastRow, 7).getValues() // This is the range of data that will be printed
  const wsTo = sheet.getSheetByName('PUT DESTINATION SHEET NAME') //Enter the sheet name destination
  const range = wsTo.getRange(wsTo.getLastRow(), 1, dataFrom.length, dataFrom[0].length) // This is the destination range
  
      range.setValues(dataFrom)
  
  //after the data is successfully transferred
  //the original sheet will be cleaned
  wsFrom.getRange(2, 1, lastRow, 7).clear()

}


//helper function
//this to get the last row without blank data
function getLastRowSpecial(range){
  var rowNum = 0;
  var blank = false;
  for(var row = 0; row < range.length; row++){
 
    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;
    }else if(range[row][0] !== ""){
      blank = false;
    };
  };
  return rowNum;
};

```


### print selection range from google sheet
this is function to print selected range

copy this file gs [printdata.gs](https://github.com/aguskusmara/google-apps-script-/blob/main/print%20data.gs)

```
//print data
function printData() {
  const sheet = SpreadsheetApp.getActive()
  const ws = sheet.getSheetByName('PUT SHEET NAME HERE')
  const lastRows = getLastRowSpecial(ws.getRange("put range").getDisplayValues()) //see helper function
  const range = ws.getRange(7, 11, lastRows, 4)
  exportCurrentSheetAsPDF(range, ws) //see exportCurrentSheetAsPDF function
  console.log(range)
}

//helper function
//this to get the last row without blank data
function getLastRowSpecial(range){
  var rowNum = 0;
  var blank = false;
  for(var row = 0; row < range.length; row++){
 
    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;
    }else if(range[row][0] !== ""){
      blank = false;
    };
  };
  return rowNum;
};

//export range of data from google sheet to PDF
function exportCurrentSheetAsPDF(range,currentSheet) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet,range)
  _exportBlob(blob, currentSheet.getName(), spreadsheet)
}

//create blob from range google sheet
function _getAsBlob(url, sheet, range) {
  var rangeParam = ''
  var sheetParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }

  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=LETTER'
      + '&portrait=true'
      + '&fitw=true'       
      + '&top_margin=0.75'              
      + '&bottom_margin=0.75'          
      + '&left_margin=0.7'             
      + '&right_margin=0.7'           
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + '&gridlines=true'
      + '&fzr=FALSE'      
      + sheetParam
      + rangeParam
      
  Logger.log('exportUrl=' + exportUrl)
  var response
  var i = 0
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000)
    } else {
      break
    }
  }
  
  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.')
  }
  
  return response.getBlob()
}

//generate link to open the pdf and printed
// By default, PDFs are saved in your Drive Root folder
// To save in the same folder as the spreadsheet, change the value to 'false' without the single quote pair
// You must have EDIT permission to the same folder
var saveToRootFolder = true
function _exportBlob(blob, fileName, spreadsheet) {
  blob = blob.setName(fileName)
  var folder = saveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next()
  var pdfFile = folder.createFile(blob)
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
    .setWidth(300)
    .setHeight(80)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Print')
}
```


### install
create menu button in google sheet

copy this file gs [install.gs](https://github.com/aguskusmara/google-apps-script-/blob/main/install.gs)
```
//create custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('APP') //name of menu
    .addItem('PRINT', 'printData') //PRINT is name of item & printData is name of function
    .addItem('MOVE DATA','moveData')
    .addToUi()
}
```
