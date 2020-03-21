/** 
    @OnlyCurrentDoc
    Last edited: 21 March 2020
    Managed by: Bappy 
*/

var titleRowColor = '#d9ead3'; //light green
var currentSpreadsheet = SpreadsheetApp.getActive();

//update the color of the top row (the title row)
function ApplyTopRowColor() {
  currentSpreadsheet.getRange('1:1').activate();
  currentSpreadsheet.getActiveRangeList().setBackground(titleRowColor);  
}; 

//make the top row bold
function Untitledmacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveRangeList().setFontWeight('bold');
};