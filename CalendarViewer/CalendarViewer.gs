function initializeCalendarViewer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Open or create Setup sheet and clear it
  var setupSheet = ss.getSheetByName("Setup");
  if (setupSheet == null) {
    setupSheet = ss.insertSheet("Setup");
  }
  setupSheet.clear();
  
  // Add headers, freeze top row and make it bold
  setupSheet.getRange(1, 1, 1, 3)
    .setFontWeight("bold")
    .setValues([["Calendar", "Calendar ID","Show (Y/N)"]]);
  setupSheet.setFrozenRows(1);
  
  // Hide Calendar ID column
  setupSheet.hideColumns(2);
  
  // Add all subscribed calendars unselect them
  var calendars = CalendarApp.getAllCalendars();
  var calendarValues = [];
  for (var i = 0; i < calendars.length; i++) {
    calendarValues[i] = [calendars[i].getName(), calendars[i].getId(),"N"];
  }
  setupSheet.getRange(2, 1, calendars.length, 3).setValues(calendarValues);
  
  // Set data validation for the Show column
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["Y", "N"], true).build();
  setupSheet.getRange(2, 3, calendars.length, 1).setDataValidation(rule);
}

function updateMonthlyView(month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var monthlySheet = ss.getSheetByName("Monthly View");
  if (monthlySheet == null) {
    monthlySheet = ss.insertSheet("Monthly View");
  }
  monthlySheet.clear();
  
  // Set reference date to today
  var date = new Date();
  var firstDay = (date.getDay() - ((date.getDate() - 1) % 7)) % 7;
  var monthStart = date.getTime() - ((date.getDay() - 1) * 1000 * 60 * 60 * 24) - ((date.getHours() - 1) * 1000 * 60 * 60);
  var monthEnd = monthStart + (30 * 1000 * 60 * 60 * 24);
  
  var monthlyRange = monthlySheet.getRange(1, 1, 7, 7);
  var monthlyValues = [];
  monthlyValues[0] = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
  for (var i = 1; i < 7; i++) {
    monthlyValues[i] = ["","","","","","",""];
  }
  
  // Get events from selected calendars in the setup sheet
  var setupSheet = ss.getSheetByName("Setup");
  var calendarList = setupSheet.getRange(2, 2, setupSheet.getLastRow() - 1, 2).getValues();
  
  for (var i = 1; i < 31; i++)
  {
    var week = Math.floor((i-1+firstDay)/7) + 2;
    var day = (firstDay + i - 1) % 7;
    if (day < 0) day += 7;
    monthlyValues[week][day] = i;
    
    var dayEvents = [];
    
    for (var c = 0; c < calendarList.length; c++)
    {
      if (calendarList[c][1] == "Y") {
        var d = new Date();
        d.setDate(i);
        dayEvents = dayEvents.concat(CalendarApp.getCalendarById(calendarList[c][0]).getEventsForDay(d));
      }
    }
    
    dayEvents.sort({
    
    });
    
    for (var eventIndex = 0; eventIndex < dayEvents.length; eventIndex++) {
      monthlyValues[week][day] += "\n" + dayEvents[eventIndex].getTitle();
    }
  }
  
  monthlyRange.setValues(monthlyValues)
    .setVerticalAlignment("top")
    .setHorizontalAlignment("left")
    .setWrap(true);
}

// Add a custom menu
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Calendar Viewer")
    .addItem("Initialize", "initializeCalendarViewer")
    .addItem("Update calendars", "updateMonthlyView")
    .addToUi();
}
