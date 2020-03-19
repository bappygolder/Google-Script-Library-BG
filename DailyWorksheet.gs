/*
Author: Dhawal Jaiswal
Date: 2nd March 2020

To Do:
1. Create month's sheet if does not exists. DONE
2. Check date in last created section, and create section for all dates till now. DONE
3. Add protection and release protection on task columns. DONE
4. Vertical align, wrap text and insert sheet at beginning (left side). DONE
*/
function onOpen(e) {
  
  // Create month's sheet if does not exists.
  var yearStr = Utilities.formatDate(new Date(),"Australia/Sydney","yyyy");
  var monthStr = Utilities.formatDate(new Date(),"Australia/Sydney","MM");
  var dateStr = Utilities.formatDate(new Date(),"Australia/Sydney","dd");
  var now = new Date(yearStr, Number(monthStr) - 1, dateStr); // Date based on Sydney, Australian Time Zone
  var year = now.getFullYear();
  var month = now.getMonth();
  var monthsStringArray = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul","Aug", "Sep", "Oct", "Nov", "Dec"]; // Array of months to create name
  var sheetName = monthsStringArray[month] + " - " + year; // Example of sheet name Mar - 2020
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Get active sheet
  var monthlySheet = ss.getSheetByName(sheetName); // Get active sheet by name
  if (!monthlySheet) {
    monthlySheet = ss.insertSheet(sheetName, 0); // Create new sheet if doesn't exists
  } 
  
  // Vertical align whole sheet
  monthlySheet.getRange(1, 1, monthlySheet.getMaxRows(), monthlySheet.getMaxColumns()).setVerticalAlignment('middle');
  
  // Wrap Text whole sheet
  monthlySheet.getRange(1, 1, monthlySheet.getMaxRows(), monthlySheet.getMaxColumns()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Add protection to whole sheet, later release protection on range
  // var protection = monthlySheet.protect();
  
  // Check date in last created section, and create section for all dates till now.
  var day = now.getDay(); // Will return integer value of present day
  var daysStringArray = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]; // Array of days to create column title
  var date = now.getDate();
  var dayColumnTitle = daysStringArray[day] + ", " + date + " " + monthsStringArray[month];
  var addedColumnTitle = "Added Time";
  var updatedColumnTitle = "Last Update Time";
  var headerRow = 1;
  var dayColumnWidth = 222;
  var addedColumnWidth = 100;
  var updatedColumnWidth = 100;
  var fontSize = 8;
  var fontWeight = 800;
  var fontFamily = "Arial";
  var backgroundColor = "#d9ead3";
  
  for (var col = 1; col < 31*3; col = col + 3) {
    var value = "";
    if (monthlySheet.getRange(headerRow, col).getValue() != "") {
      var gotDate = new Date(monthlySheet.getRange(headerRow, col).getValue());
      // Workaround to fix time issue , some issue with time zone of google script
      gotDate.setHours(gotDate.getHours() + 10);
      
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
        
        for (var row = 2; row < monthlySheet.getMaxRows(); row++) {
          if (monthlySheet.getRange(row, col).getValue() == "") {
            monthlySheet.setActiveRange(monthlySheet.getRange(row, col));
            break;
          }
        }
        
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
          monthlySheet.getRange(headerRow, dateCounter).setValue(newDayColumnTitle);
          monthlySheet.setColumnWidth(dateCounter, dayColumnWidth);
          monthlySheet.getRange(headerRow, dateCounter).setFontSize(fontSize);
          monthlySheet.getRange(headerRow, dateCounter).setFontWeight(fontWeight);
          monthlySheet.getRange(headerRow, dateCounter).setFontFamily(fontFamily);
          monthlySheet.getRange(headerRow, dateCounter).setBackground(backgroundColor);
          monthlySheet.getRange(headerRow, dateCounter).setHorizontalAlignment("center");
          monthlySheet.getRange(headerRow, dateCounter).setVerticalAlignment("center");
         
          monthlySheet.getRange(headerRow+1, dateCounter, monthlySheet.getMaxRows(), 1).setWrap(true);
        //  protection.setUnprotectedRanges([monthlySheet.getRange(headerRow+1, dateCounter, monthlySheet.getMaxRows(), 1)]);
          monthlySheet.setActiveRange(monthlySheet.getRange(headerRow+1, dateCounter));
          monthlySheet.getRange(headerRow+1, dateCounter).activate();
         /*
          var me = Session.getEffectiveUser();
          protection.addEditor(me);
          protection.removeEditors(protection.getEditors());
          if (protection.canDomainEdit()) {
            protection.setDomainEdit(false);
          }
          */
          
          monthlySheet.getRange(headerRow, dateCounter+1).setValue(addedColumnTitle);
          monthlySheet.setColumnWidth(dateCounter+1, addedColumnWidth);
          monthlySheet.getRange(headerRow, dateCounter+1).setFontSize(fontSize);
          monthlySheet.getRange(headerRow, dateCounter+1).setFontWeight(fontWeight);
          monthlySheet.getRange(headerRow, dateCounter+1).setFontFamily(fontFamily);
          monthlySheet.getRange(headerRow, dateCounter+1).setBackground(backgroundColor);
          monthlySheet.getRange(headerRow, dateCounter+1).setHorizontalAlignment("center");
          monthlySheet.getRange(headerRow, dateCounter+1).setVerticalAlignment("center");
          monthlySheet.getRange(headerRow, dateCounter+2).setValue(updatedColumnTitle);
          monthlySheet.setColumnWidth(dateCounter+2, updatedColumnWidth);
          monthlySheet.getRange(headerRow, dateCounter+2).setFontSize(fontSize);
          monthlySheet.getRange(headerRow, dateCounter+2).setFontWeight(fontWeight);
          monthlySheet.getRange(headerRow, dateCounter+2).setFontFamily(fontFamily);
          monthlySheet.getRange(headerRow, dateCounter+2).setBackground(backgroundColor);
          monthlySheet.getRange(headerRow, dateCounter+2).setHorizontalAlignment("center");
          monthlySheet.getRange(headerRow, dateCounter+2).setVerticalAlignment("center");
          monthlySheet.getRange(headerRow, dateCounter, monthlySheet.getMaxRows(), 3).setBorder(false, true, false, true, false, false, null, null);
          if (dateCounter == 1) { monthlySheet.setFrozenRows(1); }
          if (newDayColumnTitle == dayColumnTitle) { Logger.log("EXIT 3"); break; }
        } else { Logger.log("EXIT 4"); break; }
      }
      Logger.log("EXIT 2");
      break;
    }
  }
}
