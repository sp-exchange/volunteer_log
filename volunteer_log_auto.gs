
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Update Sheets');
  menu.addItem('Add New Project', 'createNewProject')
      .addToUi();
}

function createNewProject(){
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var gridId = spreadSheet.getSheetId();
  var mainSheet = spreadSheet.getSheetByName("Full Project List");
  var templateSheet = spreadSheet.getSheetByName("Sample Project Block");
  var templateRange = templateSheet.getRange(2,1,9, mainSheet.getLastColumn());
  var templateData = templateRange.getValues();
  var targetRange = mainSheet.getRange(mainSheet.getLastRow()+1, 1, 9,mainSheet.getLastColumn());

  templateRange.copyTo(targetRange);

}

function onEdit() {
  // assumes source data in sheet named main
  // target sheet of move to named Completed
  // getColumn with check-boxes is currently set to column 4 or D
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = SpreadsheetApp.getActiveSheet();
  var r = SpreadsheetApp.getActiveRange();

  if(s.getName() == "Full Project List" && r.getColumn() == 6 && r.getValue() == true) {
    var row = r.getRow();
    var numColumns = s.getLastColumn();
    var targetSheet = ss.getSheetByName("Active Project Dashboard");
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    
    s.getRange(row, 1, 9, numColumns).copyTo(target);
  
  } 
  
  else if(s.getName() == "Full Project List" && r.getColumn() == 6 && r.getValue() == false) {
    var row = r.getRow();
    var projName = s.getRange(row,1).getValue(); 
    var numColumns = s.getLastColumn();
    var targetSheet = ss.getSheetByName("Active Project Dashboard");
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);

    // get the values in an array
    var range = targetSheet.getRange("A2:A");
    var values = range.getValues();

    
    let Row = range.createTextFinder(projName).matchEntireCell(true).findNext().getRow();
    
    targetSheet.deleteRows(Row, 9);
    //targetSheet.getRange(rowNum, 1, 1, numColumns).clear();
  
  }
}


