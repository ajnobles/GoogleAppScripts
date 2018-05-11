// IMPORT CODE
// NEEDS TO IMPLEMENT
//     ALARM FUNCTIONALITY
//     REMOVE GOOGLE HANGOUT INFORMATION
//     ADD PROJECT ADDRESS

// SPREADSHEET COLUMN NUMBER GLOBAL  
// 0 BASED
var ID = 0;
var TITLE = 1;
var LOCATION = 2;
var DATE = 3;
var TIME = 4;
var PREBID_DATE = 5;
var PREBID_TIME = 6;
var ADDENDUM = 7;
var HVAC = 8;
var PLUMBING = 9;
var SHAREFILE = 10;
var GOOGLE_CAL_EVENT_ID = 11;
var GOOGLE_CAL_EVENT_ID_PREBID = 12;

// OTHER GLOBALS
var GOOGLE_CAL_ID = ""  // INSERT CALENDAR ID BETWEEN QUOTES
var SPREADSHEET = "";  // INSERT SPREADSHEET SHEETNAME BETWEEN QUOTES
var SLEEP = 1000;
var SLEEP_MOD_OCCUR = 10;
var BID_ALARM = 24;   // IN HOURS
var PREBID_ALARM = 6; // IN HOURS
var BID_COLOR = 9;
var PREBID_COLOR = 11;
var DEFAULT_TIME = "17:00:00";
var DELETE_START = "01/01/2017 12:00 AM"
var DELETE_STOP = "12/31/2018 11:59 PM"


// MAIN FUNCTION TO RUN CALENDAR UPDATES
function refreshCalendar() {
  var cal_Id = GOOGLE_CAL_ID;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SPREADSHEET);
  var dataRange = ss.getDataRange();
  var values = dataRange.getValues();
  
  var cal = CalendarApp.getCalendarById(cal_Id);
  var numEventsAdded = 0;
  
  Logger.log("NUMBER OF ROWS: " + values.length);
  
  var i = values.length-1;
  while(values[i--][TITLE] == null);
  
  // ITERATE OVER ITEMS IN SPREADSHEET
//  for (var i = 1; i < values.length && values[ID]; i++) {
  for (; i >= 1 && values[ID]; i--) {
    var value = values[i];
    
    Logger.log(value[TITLE]);
    
    // BID EVENTS
    if (value[GOOGLE_CAL_EVENT_ID]) {
      if (value[DATE]) {
        updateBid(ss, cal, value, i);
      }
      
      else {
        removeBid(ss, cal, value, i);
        numEventsAdded++;
      }
    }
    
    else if(value[TITLE] && value[DATE]) {
      addBid(ss, cal, value, i);
      numEventsAdded++;
    }
    
    // PRE-BID EVENTS
    if (value[GOOGLE_CAL_EVENT_ID_PREBID]) {
      if (value[DATE]) {
        updatePreBid(ss, cal, value, i);
      }
      
      else {
        removePreBid(ss, cal, value, i);
        numEventsAdded++;
      }
    }
    
    else if(value[TITLE] && value[PREBID_DATE]) {
      addPreBid(ss, cal, value, i);
      numEventsAdded++;
    }
    
    // USED FOR DEBUGGING
    Logger.log("i: " + i);
    
    if (numEventsAdded % SLEEP_MOD_OCCUR == 0) {
      Utilities.sleep(SLEEP)
    }
  }
}


function addBid(ss, cal, value, i) {
  
  var bid = {id:0, title:0, location:0, bidTime:0, description:0};

  bid.id = value[ID] || 0;
  bid.title = value[TITLE] || 0;
  bid.location = value[LOCATION] || 0;
    
  bid.bidTime = new Date(parseDate(value[DATE],value[TIME]))

  var event = cal.createEvent(bid.title, bid.bidTime, bid.bidTime);

  // ADD CALENDAR ID TO SPREADSHEET
  ss.getRange(i+1,GOOGLE_CAL_EVENT_ID+1).setValue(event.getId());
  
  // SET REMINDERS
  event.removeAllReminders();
  
  // SET GUEST PERMISSIONS
  event.setGuestsCanInviteOthers(false);
  event.setGuestsCanModify(false);
  event.setGuestsCanSeeGuests(false);
  
  // SET COLOR
  event.setColor(BID_COLOR);
  
  // SET ALARM  
  
  // SET DESCRIPTION
  event.setDescription(buildDescription(value));
}



function updateBid(ss, cal, value, i) {

  var currentEvent = cal.getEventById(value[GOOGLE_CAL_EVENT_ID]);
  var currentValues = {id:0, title:0, description:0, bidTime:0};
  var newValues = {id:0, title:0, description:0, bidTime:0};
  
  // COMPILE VALUES
  currentValues.title = currentEvent.getTitle();
  currentValues.bidTime = currentEvent.getStartTime();
  currentValues.description = currentEvent.getDescription();
  // GET CURRENT ALARM
  
  newValues.title = value[TITLE];
  newValues.bidTime = new Date(parseDate(value[DATE],value[TIME]));
  newValues.description = buildDescription(value);
  // GET 'NEW' ALARM
  

  // TEST VALUES  
  if (currentValues.title != newValues.title) {
    currentEvent.setTitle(newValues.title);
  }
  
  if (currentValues.bidTime != newValues.bidTime){
    currentEvent.setTime(newValues.bidTime, newValues.bidTime);
  }
  
  if (currentValues.description != newValues.description) {
    currentEvent.setDescription(newValues.description);
  }
  
  // TEST ALARM
}


function removeBid(ss, cal, value, i) {
  var currentEvent = cal.getEventById(value[GOOGLE_CAL_EVENT_ID]);
  
  // DELETE CALENDAR EVENT
  currentEvent.deleteEvent();
  
  // DELETE CALENDAR ID FROM SPREADSHEET
  ss.getRange(i+1,GOOGLE_CAL_EVENT_ID+1).setValue(null);
}



function addPreBid(ss, cal, value, i) {
  
  var bid = {id:0, title:0, location:0, bidTime:0, description:0};

  bid.id = value[ID] || 0;
  bid.title = value[TITLE] || 0;
  bid.location = value[LOCATION] || 0;
    
  bid.bidTime = new Date(parseDate(value[PREBID_DATE],value[PREBID_TIME]))

  var event = cal.createEvent(bid.title, bid.bidTime, bid.bidTime);
  
  // SET REMINDERS
  event.removeAllReminders();
  
  // SET GUEST PERMISSIONS
  event.setGuestsCanInviteOthers(false);
  event.setGuestsCanModify(false);
  event.setGuestsCanSeeGuests(false);
  
  // SET COLOR
  event.setColor(PREBID_COLOR);
  
  
  // SET DESCRIPTION
  event.setDescription(buildDescription(value));

  // ADD CALENDAR ID TO SPREADSHEET
  ss.getRange(i+1,GOOGLE_CAL_EVENT_ID_PREBID+1).setValue(event.getId());
}


function updatePreBid(ss, cal, value, i) {

  var currentEvent = cal.getEventById(value[GOOGLE_CAL_EVENT_ID_PREBID]);
  var currentValues = {id:0, title:0, description:0, bidTime:0};
  var newValues = {id:0, title:0, description:0, bidTime:0};
  
  // COMPILE VALUES
  currentValues.title = currentEvent.getTitle();
  currentValues.bidTime = currentEvent.getStartTime();
  currentValues.description = currentEvent.getDescription();
  
  newValues.title = value[TITLE];
  newValues.bidTime = new Date(parseDate(value[PREBID_DATE],value[PREBID_TIME]));
  newValues.description = buildDescription(value);
  

  // TEST VALUES  
  if (currentValues.title != newValues.title) {
    currentEvent.setTitle(newValues.title);
  }
  
  if (currentValues.bidTime != newValues.bidTime){
    currentEvent.setTime(newValues.bidTime, newValues.bidTime);
  }
  
  if (currentValues.description != newValues.description) {
    currentEvent.setDescription(newValues.description);
  }
}


function removePreBid(ss, cal, value, i) {
  var currentEvent = cal.getEventById(value[GOOGLE_CAL_EVENT_ID_PREBID]);
  
  // DELETE CALENDAR EVENT
  currentEvent.deleteEvent();
  
  // DELETE CALENDAR ID FROM SPREADSHEET
  ss.getRange(i+1,GOOGLE_CAL_EVENT_ID_PREBID+1).setValue(null);
}



function buildDescription(value)
{
  var description = "";
  description += "Project ID: " + value[ID] + "\n";
  description += "Bid Location: " + value[LOCATION] + "\n";
  description += "Addenda: " + value[ADDENDUM] + "\n";
  
  description += "HVAC: ";
  description += (value[HVAC] || "No Bid");
  
  description += "\nPLUMBING: ";
  description += (value[PLUMBING] || "No Bid");
  
  description += "\nSHAREFILE: " + value[SHAREFILE] + "\n";
  
  return description;
}


function parseDate(date, time) {
//  Logger.log("date: " + date);
//  Logger.log("\ntime: " + time);
  
  var dateString = "";
  dateString += date;
  var dateParsed = ""  
  if (date) {
    for (var j = 0; j < 11; j++) {
      dateParsed += dateString[j+4];
      
      // REFORMAT MONTH
      if (dateParsed === "Jan") {
        dateParsed = "January";
      }
      
      else if (dateParsed === "Feb") {
        dateParsed = "February";
      }
      
      else if (dateParsed === "Mar") {
        dateParsed = "March";
      }
      
      else if (dateParsed === "Apr") {
        dateParsed = "April";
      }
      
      else if (dateParsed === "May") {
        dateParsed = "May";
      }
      
      else if (dateParsed === "Jun") {
        dateParsed = "June";
      }
      
      else if (dateParsed === "Jul") {
        dateParsed = "July";
      }
      
      else if (dateParsed === "Aug") {
        dateParsed = "August";
      }
      
      else if (dateParsed === "Sep") {
        dateParsed = "September";
      }
      
      else if (dateParsed === "Oct") {
        dateParsed = "October";
      }
      
      else if (dateParsed === "Nov") {
        dateParsed = "November";
      }
      
      else if (dateParsed === "Dec") {
        dateParsed = "December";
      }
      
      
    }
  }  
    
  if (time) {
    var timeString = "";
    timeString += time;
    
    var timeParsed = "";
    for (var j = 0; j < 8; j++) {
     
      timeParsed += timeString[j+16];
    }
    
    // ADJUST FOR TIME ZONE
    timeParsed += " ";
    timeParsed += "-";
    timeParsed += "0";
    timeParsed += "5";
    timeParsed += "0";
    timeParsed += "0";
    
    // DO NOT DELETE LOOP
// USE FOR DEBUGGING IF TIME ZONE MESSES UP
//    for (var j = 0; j < 5; j++) {
//      timeParsed += timeString[j+28];
//    }
  }
  
  // ELSE NO TIME LISTED
  else {
    timeParsed = DEFAULT_TIME;
  }
  
//  Logger.log(timeParsed);
  
  var newDateString = dateParsed + " " + timeParsed;
  
  return newDateString;
}

function deleteAll()
{
  var cal_Id = GOOGLE_CAL_ID;
  var cal = CalendarApp.getCalendarById(cal_Id);
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SPREADSHEET);
  
  var start = new Date(DELETE_START);
  var end = new Date (DELETE_STOP);

  var events = cal.getEvents(start, end);
  
  for (var i = events.length - 1; i >= 0; i--) {
    var event = events[i];
    event.deleteEvent();
    
    if (i % SLEEP_MOD_OCCUR == 0) {
      Utilities.sleep(SLEEP);
    }
  }
    
  var values = ss.getDataRange().getValues();
  
  for (var i = 2; i < values.length; i++) {
    ss.getRange(i,GOOGLE_CAL_EVENT_ID+1).setValue(null);
    ss.getRange(i,GOOGLE_CAL_EVENT_ID_PREBID+1).setValue(null);
  }
}
