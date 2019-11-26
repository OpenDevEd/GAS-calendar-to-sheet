/*
Enhancement needed: Simple UI for:
- Timerange
- calendar id
- which fields to display.
- destination sheet
- select whether script is run manually or 'onOpen'.
*/

var cals = ["calendarid@group.calendar.google.com"];

function onOpen() {
  launch();
}

function launch() {
  info(1);
  for (i=0;i<=cals.length-1;i++) {
    listEvents(cals[i],i+1);
  };
  info(0);
}

function makeActive(number) {
  var outputid = "xxxxxx"; // current sheet
  var spsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/'+outputid+'/edit#');
  var setsheet = spsheet.getSheets()[number];
  spsheet.setActiveSheet(setsheet);
  sheet = spsheet.getActiveSheet();
  return sheet;
}

function info(action) {
  var sheet = makeActive(0);
  var today = new Date();
  var infoarr = [];
  if (action==1) {
     infoarr = [[ "Please wait - Sheet is updating", today ]];
  } else {
   infoarr = [[ "Sheet is up to date.", today ]];
  };
  sheet.getRange(1, 1, 1, infoarr[0].length ).setValues(infoarr);
};

function listEvents(mycal,sheetno) {
  var sheet = makeActive(sheetno);
  var today = new Date();
  var Calendar  =  CalendarApp.getCalendarById(mycal);
  //var Calendar  =  CalendarApp.getDefaultCalendar();

  // Calendar name and other IDs:
  var infoarr = [
    [ Calendar.getName(),
     Calendar.getId() ,
     Calendar.getTimeZone() ,
     Calendar.getColor() 
    ]
  ];
  sheet.getRange(1, 1, 1, infoarr[0].length ).setValues(infoarr);

  // Get events: Thisyear-1 to thisyear+1 
  var events = Calendar.getEvents(new Date(today.getFullYear()-1,1,1), new Date(today.getFullYear()+1,12,31));
  
  var eventarray = new Array();
  for (var i = 0; i<events.length; i++) {
    var allDayStart = "";
    var allDayEnd = "";
    if (events[i].isAllDayEvent == "TRUE") {
      allDayStart = events[i].getAllDayStartDate();
      allDayEnd = events[i].getAllDayEndDate();
    };
    var designation = "";
    var title = events[i].getTitle();
    var result = /^\[?(travel|lodge|office|visit)\]?/i.exec(title);
    var destination = "";
    var origin = "";
    if (result != null) {
      designation = result[1];
      var result = /^\[?(travel)\]?.* to (.*)/i.exec(title);
      if (result != null) {
        destination = result[2];
      };
      var result = /^\[?(travel)\]?\s*from (.*?)\s*to/i.exec(title);
      if (result != null) {
        origin = result[2];
      } else {
        origin = events[i].getLocation();
      };
    };
    var iscomplete = 0;
    if (events[i].getEndTime() < today) {
      iscomplete = 1;
    }
    var duration = (events[i].getEndTime() - events[i].getStartTime())/3600/1000;
    var multiday = false;
    if (duration > 24) {
      multiday = true;
    };
    var line = [
      title,
      events[i].getLocation(),
      events[i].getStartTime(),
      events[i].getEndTime(),
      duration,
      events[i].getDescription(),
      events[i].getCreators().join("; ")      ,
      iscomplete,
      events[i].getLastUpdated()	,      
      events[i].getDateCreated()	,
      events[i].isAllDayEvent(),
      multiday /*,
      designation,
      origin,
      destination,
      events[i].isOwnedByMe()	,
      events[i].isRecurringEvent(),
      events[i].getVisibility(),
      allDayStart,
      allDayEnd,
      events[i].getAllTagKeys().join("; "),	
      events[i].getOriginalCalendarId()	,
      events[i].getId(),
      events[i].getEventSeries()	*/
    ];
    /*    
getDateCreated()	
getOriginalCalendarId()	

getEventSeries()	
getLastUpdated()	
    */    
    eventarray.push(line);
  }
  infoarr = [
    ["Title","Location","Start Time","End Time","Duration (hours)","Description","Created by","complete","Last Updated on","Created on","All Day Event",
     "Multi Day Event"
    /* ,"designation","origin","destination",
    "isOwnedByMe","isRecurringEvent","getVisibility",
    "getAllDayStartDate","getAllDayEndDate","getAllTagKeys[]",
    "getOriginalCalendarId","getId","getEventSeries"
    */  
     ,  eventarray.length
    ]
    ]
  sheet.getRange(2, 1, 1, infoarr[0].length ).setValues(infoarr);
  sheet.getRange(3, 1, eventarray.length, eventarray[0].length).setValues(eventarray);
}

