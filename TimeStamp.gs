/*
Author: Dhawal Jaiswal
Date: 2nd March 2020

To Do:
1. Check if cell is updated then add timestamp in Added Time if empty,
else add timestamp to Last Updated Time. DONE
2. Format Text after timestamp is added. DONE
*/
function onEdit(e) {
  
  var yearStr = Utilities.formatDate(new Date(),"Australia/Sydney","yyyy");
  var monthStr = Utilities.formatDate(new Date(),"Australia/Sydney","MM");
  var dateStr = Utilities.formatDate(new Date(),"Australia/Sydney","dd");
  var now = new Date(yearStr, Number(monthStr) - 1, dateStr); // Date based on Sydney, Australian Time Zone
  var year = now.getFullYear();
  var month = now.getMonth();
  var date = now.getDate();
  var monthsStringArray = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul","Aug", "Sep", "Oct", "Nov", "Dec"]; // Array of months to create name
  var sheetName = monthsStringArray[month] + " - " + year; // Example of sheet name Mar - 2020
  Logger.log(sheetName);
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Get active sheet
  var monthlySheet = ss.getSheetByName(sheetName); // Get active sheet by name
  var headerRow = 1;
  var day = now.getDay(); // Will return integer value of present day
  var daysStringArray = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]; // Array of days to create column title
  var hours = Utilities.formatDate(new Date(),"Australia/Sydney","HH");
  var min = Utilities.formatDate(new Date(),"Australia/Sydney","mm");
  var isPM = Number(hours) >= 12;
  var isMidday = Number(hours) == 12;
  var time = [Number(hours) - (isPM && !isMidday ? 12 : 0), 
              Number(min) > 9 ? Number(min) : "0" + Number(min)].join(":") +
           (isPM ? "pm" : "am");
  var timeStamp = daysStringArray[day] + " " + date + " " + monthsStringArray[month] + ", " + time;
  var fontSize = 8; // Font size of timestamp
  var fontFamily = "Arial"; // Font family used in timestamp
  var color = "#666666"; // Font color used in timestamp
  var fontSizeMain = 10; // Font size of tasks
  var fontFamilyMain = "Arial"; // Font family of tasks
  var colorMain = "#000000"; // Font color of tasks
  var dayColumnTitle = daysStringArray[day] + ", " + date + " " + monthsStringArray[month];
  
  if (!monthlySheet) { 
    // do nothing 
  } else {
    for (var col = 1; col < 31*3; col = col + 3) {
      var value = "";
      if (monthlySheet.getRange(headerRow, col).getValue() != "") {
        var gotDate = new Date(monthlySheet.getRange(headerRow, col).getValue());
        // Workaround to fix time issue , some issue with time zone of google script
        gotDate.setHours(gotDate.getHours() + 11);
        
        var gotYearStr = Utilities.formatDate(gotDate,"Australia/Sydney","yyyy");
        var gotMonthStr = Utilities.formatDate(gotDate,"Australia/Sydney","MM");
        var gotDateStr = Utilities.formatDate(gotDate,"Australia/Sydney","dd");
        var gotNow = new Date(gotYearStr, Number(gotMonthStr) - 1, gotDateStr);
        var gotDay = gotNow.getDay();
        var gotDate = gotNow.getDate();
        var gotMonth = gotNow.getMonth();
        value = daysStringArray[gotDay] + ", " + gotDate + " " + monthsStringArray[gotMonth];
      }
      
      if (value != "") {
        if (value == dayColumnTitle) {
          Logger.log("EXIT 1");
          break;
        } 
      } else {
        for (var dateCounter = col; dateCounter < 31*3; dateCounter = dateCounter + 3) {
          var newDate = dateCounter > 3 ? Math.floor(dateCounter/3) + 1 : dateCounter;
          var newNow = new Date(year, month, newDate);
          if (month == newNow.getMonth()) {
            var newDay = newNow.getDay();
            var newDayColumnTitle = daysStringArray[newDay] + ", " + newDate + " " + monthsStringArray[month];
            
            monthlySheet.setActiveRange(monthlySheet.getRange(headerRow+1, dateCounter));
            monthlySheet.getRange(headerRow+1, dateCounter).activate();
            
            if (dateCounter == 1) { monthlySheet.setFrozenRows(1); }
            if (newDayColumnTitle == dayColumnTitle) { Logger.log("EXIT 3"); break; }
          } else { Logger.log("EXIT 4"); break; }
        }
        Logger.log("EXIT 2");
        break;
      }
    }
    
    var activeCell = monthlySheet.getActiveCell();
    if (monthlySheet.getRange(headerRow, activeCell.getColumn()+1).getValue() == "Added Time"
       && activeCell.getValue() != "") {
      monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()).setFontSize(fontSizeMain);
      monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()).setWrap(true);
      monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()).setFontFamily(fontFamilyMain);
      monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()).setFontColor(colorMain);
      if (monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+1).getValue() == "") {
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+1).setValue(timeStamp);
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+1).setFontSize(fontSize);
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+1).setWrap(true);
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+1).setFontFamily(fontFamily);
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+1).setFontColor(color);
        monthlySheet.hideColumns(activeCell.getColumn()+1); // Hide column
      } else {
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+2).setValue(timeStamp);
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+2).setFontSize(fontSize);
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+2).setWrap(true);
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+2).setFontFamily(fontFamily);
        monthlySheet.getRange(activeCell.getRowIndex(), activeCell.getColumn()+2).setFontColor(color);
        monthlySheet.hideColumns(activeCell.getColumn()+2); // Hide column
      }  
    }
  }
}
