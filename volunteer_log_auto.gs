
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Update Sheets');
  menu.addItem('Add New Project', 'createNewProject')
      .addItem('Update Dashboard', 'updateDashboard')
      .addToUi();
}

//This function is to add Sample Project Block into the 'Full Project List' sheet to add new projects
function createNewProject(){
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = spreadSheet.getSheetByName("Full Project List");
  var templateSheet = spreadSheet.getSheetByName("Sample Project Block");
  var templateRange = templateSheet.getRange(2,1,9, mainSheet.getLastColumn());
  var targetRange = mainSheet.getRange(mainSheet.getLastRow()+1, 1, 9,mainSheet.getLastColumn());

  templateRange.copyTo(targetRange); //copies template to target sheet

}

function updateDashboard() {
  var sSheet = SpreadsheetApp.getActiveSpreadsheet();
  var srcSheet = sSheet.getSheetByName("Full Project List");
  var tarSheet = sSheet.getSheetByName("Active Project Dashboard");
  var lastRow = srcSheet.getLastRow();
  var lastColumn =srcSheet.getLastColumn();

  var start = 2; //Hard coded row number from where to start deleting

  var howManyToDelete = tarSheet.getLastRow() - start + 1;//How many rows to delete -
      //The blank rows after the last row with content will not be deleted

  tarSheet.deleteRows(start, howManyToDelete);

  for (var i = 2; i <= lastRow; i+=9) {
    var cell = srcSheet.getRange("F" + i);
    var val = cell.getValue();
    if (val == true) {
      
      var srcRange = srcSheet.getRange(i, 1, 9, lastColumn);
      
      var tarRow = tarSheet.getLastRow();
      tarSheet.insertRowAfter(tarRow);
      var tarRange = tarSheet.getRange(tarRow+1, 1);
      
      srcRange.copyTo(tarRange);
    }
  }
};




/*
//This function will copy the project information from 'Full Project List' sheet to 'Active Project Dashboard' when the tickbox is clicked
//The function will also delete any projects that are unclicked from the 'Active Project Dashboard'
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

*/