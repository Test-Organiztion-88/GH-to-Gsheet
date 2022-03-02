//Forked from main GH-to-Gsheet//
//eliminating all functions and making into one//
//Date 3.1.22//

//Next eliminate duplicate 'var spreadsheet = ...'//

function Test1(){

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getActiveRange().getDataRegion().activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setFontFamily('Arial')
  .setFontSize(11)
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('M:M').activate();
  spreadsheet.getActiveSheet().moveColumns(spreadsheet.getRange('M:M'), 5);
  spreadsheet.getRange('R:R').activate();

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G:G').activate();
  spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
  spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
  spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
  spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
  spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
  spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
  spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
  spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
  spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
  spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().setFrozenRows(1);
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  .setHorizontalAlignment('center');

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G1').activate();
  spreadsheet.getCurrentCell().setValue('Status');
  spreadsheet.getRange('H1').activate();
  spreadsheet.getCurrentCell().setValue('Notes');
  spreadsheet.getRange('H:H').activate();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getRange('G1:H1').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center')
  .setFontWeight('bold');

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.PREVIOUS).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBackground('#d0e0e3');
  spreadsheet.getRange('H1').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBackground('#d0e0e3');
  spreadsheet.getRange('G1').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ff9900')
  .setBackground('#666666');
  spreadsheet.getRange('G2').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBackground('#efefef');
  spreadsheet.getRange('H970').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getRange('H2').activate();

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 5, ascending: true});
  spreadsheet.getRange('F2').activate();

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I9').activate();
  spreadsheet.getActiveSheet().setColumnWidth(8, 321);
  spreadsheet.getActiveSheet().setColumnWidth(7, 144);
  spreadsheet.getRange('A2').activate();
};
