//move data
function moveData() {
  const sheet = SpreadsheetApp.getActive()
  const wsFrom = sheet.getSheetByName('PUT SHEET NAME FROM') //Enter the sheet name that will be moved
  const lastRow = getLastRowSpecial(wsFrom.getRange('A2:A').getValues()) //see helper function
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

