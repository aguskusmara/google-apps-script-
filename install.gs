//create custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('APP') //name of menu
    .addItem('PRINT', 'printData') 
    .addItem('MOVE DATA','moveData')
    .addToUi()
}
