// Keys of object fields are values from range A5:A28 on sheet 'Configuration'
var fields = {
	
	Title : function(e) {
		return e.getTitle();
	},
	
	Location : function(e) {
		return e.getLocation();
	},
	
	StartTime : function(e) {
		return e.getStartTime();
	},
	
	EndTime : function(e) {
		return e.getEndTime();
	},
	
	Durationhours : function(e) {
		return (e.getEndTime() - e.getStartTime())/3600/1000;;
	},
	
	Description : function(e) {
		return e.getDescription();
	  },
	  
	Createdby : function(e) {
		return e.getCreators().join("; ");
	  },
	  
	Complete : function(e) {
		var today = new Date();
		var iscomplete = 0;
		if (e.getEndTime() < today) {
		  iscomplete = 1;
		}
		return iscomplete;
	  },
	  
	LastUpdatedon : function(e) {
		return e.getLastUpdated();
	  },
	  
	Createdon : function(e) {
		return e.getDateCreated();
	  },
	  
	AllDayEvent : function(e) {
		return e.isAllDayEvent();
	  },
	  
	MultiDayEvent : function(e) {
		var duration = (e.getEndTime() - e.getStartTime())/3600/1000;
		var multiday = false;		
		if (duration > 24) {
		  multiday = true;
		};
		return multiday;
	  },
	  
	designation : function(e) {
		    var designation = "";
			var title = e.getTitle();
			var result = /^\[?(travel|accommodation|office|visit)\]?/i.exec(title);
			if (result != null) {
			  designation = result[1];
			};
		return designation;
	  },
	  
	origin : function(e) {
			var title = e.getTitle();
			var result = /^\[?(travel|accommodation|office|visit)\]?/i.exec(title);
			var origin = "";
			if (result != null) {
			  var result = /^\[?(travel)\]?\s*from (.+?)( to)?/i.exec(title);
			  if (result != null) {
				origin = result[2];
			  } else {
				origin = e.getLocation();
			  };
			};		
		return origin;
	  },
	  
	destination : function(e) {
			var title = e.getTitle();
			var result = /^\[?(travel|accommodation|office|visit)\]?/i.exec(title);
			var destination = "";
			if (result != null) {
			  var result = /^\[?(travel)\]?.* to (.+)/i.exec(title);
			  if (result != null) {
				destination = result[2];
			  };
			};		
		return destination;
	  },
	  
	isOwnedByMe : function(e) {
		return e.isOwnedByMe();
	  },
	  
	isRecurringEvent : function(e) {
		return e.isRecurringEvent();
	  },
	  
	getVisibility : function(e) {
		return e.getVisibility();
	  },
	  
	getAllDayStartDate : function(e) {
		    var allDayStart = "";
			if (e.isAllDayEvent == "TRUE") {
			  allDayStart = e.getAllDayStartDate();
			};
		return allDayStart;
	  },
	  
	getAllDayEndDate : function(e) {
			var allDayEnd = "";
			if (e.isAllDayEvent == "TRUE") {
			  allDayEnd = e.getAllDayEndDate();
			};
		return allDayEnd;
	  },
	  
	getAllTagKeys : function(e) {
		return e.getAllTagKeys().join("; ");
	  },
	  
	getOriginalCalendarId : function(e) {
		return e.getOriginalCalendarId();
	  },
	  
	getId : function(e) {
		return e.getId();
	  },
	  
	getEventSeries : function(e) {
		return e.getEventSeries();
	  }  
}; 


var blueColor = '#C9DAF8';
var yellowColor = '#FFF2CC';

var configurationTemplate = [
	['Status', '','',''],					
	['Sheet status:', '', '(Do not edit.)', ''],				
	['Date:', '', '(Do not edit.)', ''],			
	['','','',''],					
	['Configuration', '(Rows 6 and 7: enter an arbitrary number of pairs of calendar ids and sheet names.)', '', ''],
	['Calendar id', '', '', ''],				
	['Sheet', '', '', ''],
	['Match', '', '', ''],
	['Timerange', '', '', '(Start date and end date for the range to be extracted.)'],			
	['Run', 'Manually', '(Whether the script runs manually or when the document is opened.)', ''],				
	['Title', 'TRUE', '(Extract calendar field \'Title\'.)', ''],	
	['Location', 'TRUE', '(Extract calendar field \'Location\'.)', ''],				
	['Start Time', 'TRUE', '(Extract calendar field \'Start Time\'.)', ''],			
	['End Time', 'TRUE', '(Extract calendar field \'End Time\'.)', ''],				
	['Description', 'TRUE', '(Extract calendar field \'Description\'.)', ''],				
	['Created by', 'TRUE', '(Extract calendar field \'Created by\'.)', ''],				
	['Last Updated on', 'TRUE', '(Extract calendar field \'Last Updated on\'.)', ''],				
	['Created on', 'TRUE', '(Extract calendar field \'Created on\'.)', ''],			
	['All Day Event', 'TRUE', '(Extract calendar field \'All Day Event\'.)', ''],				
	['Multi Day Event', 'TRUE', '(Extract calendar field \'Multi Day Event\'.)', ''],			
	['isOwnedByMe	', 'TRUE', '(Extract calendar field \'isOwnedByMe\'.)', ''],				
	['isRecurringEvent	', 'TRUE', '(Extract calendar field \'isRecurringEvent\'.)', ''],				
	['getVisibility	', 'TRUE', '(Extract calendar field \'getVisibility\'.)', ''],				
	['getAllDayStartDate', 'TRUE', '(Extract calendar field \'getAllDayStartDate\'.)', ''],				
	['getAllDayEndDate', 'TRUE', '(Extract calendar field \'getAllDayEndDate\'.)', ''],				
	['getAllTagKeys', 'TRUE', '(Extract calendar field \'getAllTagKeys\' - resulting tags are joined with \'; \'.)', ''],				
	['getOriginalCalendarId', 'TRUE', '(Extract calendar field \'getOriginalCalendarId\'.)', ''],				
	['getId', 'TRUE', '(Extract calendar field \'getId\'.)', ''],				
	['getEventSeries', 'TRUE', '(Extract calendar field \'getEventSeries\'.)', ''],				
	['Complete', 'TRUE', 'This is a custom field: Set to TRUE if the event is in the past.', ''],				
	['Duration (hours)', 'TRUE', 'This is a custom field: The calculated duration of the event in hours.', ''],				
	['designation', 'TRUE', 'This is a custom field: If the event title matches (travel|accommodation|office|visit), the matching item is returned.', ''],				
	['origin', 'TRUE', 'This is a custom field: If the designation is travel, and the title contains from ... to ..., the from field is returned.', ''],				
	['destination', 'TRUE', 'This is a custom field: If the designation is travel, and the title contains from ... to ..., the to field is returned.', ''],			
	['', '', 'Note: You can reorder the above rows to adjust the output in the sheets.', '']
];

// Create custom menu
function onOpen() {  
  SpreadsheetApp.getUi()
  .createMenu('Calendar')
  .addItem('Launch', 'launch')
  .addToUi();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = ss.getSheetByName('Configuration');
  if (sheetConfig != null){
	  var howToRun = sheetConfig.getSheetValues(10, 2, 1, 1);
	  if (howToRun == 'OnOpen'){
		launch();
	  }
  }
}


// Main function
function launch() {	
  var config = getConfiguration();	
  if (config === false) return false;
  
  info(1);  
  for (i=0;i<=config.cals.length-1;i++) {
    listEvents(config.cals[i], config.sheets[i], config.matches[i], config.indexesTLD, config.dateFrom, config.dateTo, config.fieldNames, config.fieldKeys);
  };

  info(0);
}


// Retrieve configuration from sheet 'Configuration'
function getConfiguration(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = ss.getSheetByName('Configuration');
  
  if (sheetConfig == null){	  

	  var ui = SpreadsheetApp.getUi();

	  var result = ui.alert(
		 'Please confirm',
		 'A configuration sheet needs to be created.',
		  ui.ButtonSet.YES_NO);

	  // Process the user's response.
	  if (result == ui.Button.YES){
		// User clicked "Yes".
		sheetConfig = createSheetConfiguration(ss);
	  }else{
		// User clicked "No" or X in the title bar.
		ui.alert('Script can\'t work without a configuration sheet.');
		return false;
	  }	
  }
  
  var lastRowConfig = sheetConfig.getLastRow();
  var lastColConfig = sheetConfig.getLastColumn();	

  var calendarIdsSheetNames = sheetConfig.getSheetValues(6, 2, 3, lastColConfig);

  var calendarIds = [];
  var sheetNames = [];
  var matches = [];
  var errorMessage = '';
  var checkMatches = false;
  
  for (var i in calendarIdsSheetNames[0]){
    if (calendarIdsSheetNames[0][i] != ''){
      calendarIds.push(calendarIdsSheetNames[0][i]);
      sheetNames.push(calendarIdsSheetNames[1][i]);
	  matches.push(calendarIdsSheetNames[2][i]);
	  if (calendarIdsSheetNames[2][i]) checkMatches = true;
    }
  }
  
  if (calendarIds.length == 0) 
	  errorMessage = 'Enter calendar ids. ';
  
  var dates = sheetConfig.getSheetValues(9, 2, 1, 2);
  if (dates[0][0] == '' || dates[0][1] == '') 
	  errorMessage += 'Enter start date and end date for the range to be extracted. ';
  
  var fieldsToDisplay = sheetConfig.getSheetValues(11, 1, lastRowConfig - 10, 2);
  
  var fieldKeys = [];
  var fieldNames = [];
  for (var i in fieldsToDisplay){
    if (fieldsToDisplay[i][1] === true){
      fieldNames.push(fieldsToDisplay[i][0]);
      fieldKeys.push(fieldsToDisplay[i][0].replace(/[^A-Za-z]+/g,''));
    }
  }
  if (fieldKeys.length == 0) 
	  errorMessage += 'Mark at least one calendar field. '; 

  if (errorMessage != ''){
	  if (!ui) var ui = SpreadsheetApp.getUi();
	  ui.alert(errorMessage);
	  return false;
  }


if (checkMatches){
	var indexTitle = fieldKeys.indexOf('Title');
	var indexLocation = fieldKeys.indexOf('Location');
	var indexDescription = fieldKeys.indexOf('Description');
	if (indexTitle == -1 && indexLocation == -1 && indexDescription == -1) {
		stringMatch = '';
		if (!ui) var ui = SpreadsheetApp.getUi();
		ui.alert('Warning: You marked neither Title nor Location nor Description. Setting Match won\'t affect on results.');
		var indexesTLD = false;
	}else{
		var indexesTLD = [indexTitle, indexLocation, indexDescription];
	}
}

  var config = {
    cals: calendarIds,
    sheets: sheetNames,
	matches: matches,
	indexesTLD: indexesTLD,
    dateFrom: dates[0][0],
    dateTo: dates[0][1],
    fieldNames: fieldNames,
    fieldKeys: fieldKeys
  };
  
  return config;
}

function createSheetConfiguration(ss){
	sheet = ss.insertSheet('Configuration');
	sheet.getRange(1, 1, configurationTemplate.length, configurationTemplate[0].length).setValues(configurationTemplate);

	setFontBold(sheet, 'A1');
	setFontBold(sheet, 'A5');
	setFontBold(sheet, 'B6:C9');
	setFontBold(sheet, 'C35');

	setFontItalic(sheet, 'C2:C3');
	setFontItalic(sheet, 'C10:C35');
	setFontItalic(sheet, 'D9');
	setFontItalic(sheet, 'B5');

	setBackground(sheet, 'A1:G1', blueColor);
	setBackground(sheet, 'A5:G5', blueColor);
	setBackground(sheet, 'B6:B34', yellowColor);
	setBackground(sheet, 'C6:C9', yellowColor);
	setBackground(sheet, 'D6:F8', yellowColor);

	setBorders(sheet, 'A6:B34'); 
	setBorders(sheet, 'C6:C9');  
	setBorders(sheet, 'D6:F8');  

	var cell = sheet.getRange('B10');
	var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Manually', 'OnOpen'], true).build();
	cell.setDataValidation(rule);

	var range = sheet.getRange('B9:C9');
	var rule = SpreadsheetApp.newDataValidation().requireDate().build();
	range.setDataValidation(rule);

	var range = sheet.getRange('B11:B34');
	var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
	range.setDataValidation(rule);

	var textNote = [
      ['Required. Enter a row of an arbitrary number of ids and sheet names.'],
      ['Optional. Enter a row of a sheet names, one for each calendar. If you leave this blank, the calendar name is used.'],
      ['Optional. Enter a row of stringss. Events from the calendar in this column be included if Title, Location or Description matches this string.'],
    ];
	var range = sheet.getRange("A6:A8");
	range.setNotes(textNote);

	var colsToDelete = sheet.getMaxColumns() - 7;
	if (colsToDelete > 0){
		sheet.deleteColumns(8, colsToDelete);
	}

	var rowsToDetele = sheet.getMaxRows() - 35;
	if (rowsToDetele > 0){
		sheet.deleteRows(36, rowsToDetele);
	}
	
return sheet;
}

function setFontBold(sheet, range){
	sheet.getRange(range).setFontWeight('bold');
}

function setFontItalic(sheet, range){
	sheet.getRange(range).setFontStyle('italic');
}

function setBackground(sheet, range, color){
	sheet.getRange(range).setBackground(color);	
}

function setBorders(sheet, range){
	sheet.getRange(range).setBorder(true, true, true, true, true, true);	
}

function makeActive(sheetName) {
  var spsheet = SpreadsheetApp.getActiveSpreadsheet();
  var setsheet = spsheet.getSheetByName(sheetName);
  if (setsheet == null) {
	  if (sheetName != 'Configuration')
		  setsheet = spsheet.insertSheet(sheetName);
  }
  spsheet.setActiveSheet(setsheet);
  if (sheetName != 'Configuration'){
	setsheet.clear();
	setsheet.setFrozenRows(2);
  }
  return setsheet;
}

// Set status on sheet 'Configuration'
function info(action) {
  var sheet = makeActive('Configuration');
  var today = new Date();
  var infoarr = [];
  if (action==1) {
     infoarr = [[ "Please wait - Sheet is updating"], [today ]];
  } else {
   infoarr = [[ "Sheet is up to date."], [today ]];
  };
  sheet.getRange(2, 2, 2, 1).setValues(infoarr);
};

// Retrieve calendar events
// If sheet for calendar don't exist, create it
function listEvents(mycal, sheetName, stringMatch, indexesTLD, dateFrom, dateTo, fieldNames, fieldKeys) {

  var today = new Date();
  
  var Calendar  =  CalendarApp.getCalendarById(mycal);

  if (!Calendar) {
	var ui = SpreadsheetApp.getUi();
	ui.alert('Warning: Cannot access calendar ' + mycal + '.');
	return false;
  }
  
  var calendarName = Calendar.getName();
  
  if (sheetName == '') sheetName = calendarName;
  var sheet = makeActive(sheetName);
  
  
  // Calendar name and other IDs:
  
  var calendarColor = Calendar.getColor();
  
  var infoarr = [
    [ 
      calendarName,
      Calendar.getId() ,
      Calendar.getTimeZone() ,
      calendarColor 
    ]
  ];

  	var range = sheet.getRange(1, 1, 1, infoarr[0].length );
	range.setValues(infoarr);
	range.setBackground(calendarColor);

  // Get events: dateFrom, dateTo
  var events = Calendar.getEvents(dateFrom, dateTo);

	var checkMatch = false;
	var matchFound = false;
	if (stringMatch && indexesTLD !== false){
		var patt = new RegExp(stringMatch, "mi");
		checkMatch = true;
	}
	

  var eventarray = new Array();
   for (var i = 0; i<events.length; i++) {
	   
	var line = [];
	var matchFound = false;
	
	for (var j in fieldKeys){
		var fieldValue = this['fields'][fieldKeys[j]](events[i]);
		line.push(fieldValue);
		
		if (checkMatch && !matchFound){
			if (indexesTLD.indexOf(Number(j)) != -1){
				if (fieldValue.match(patt) != null){
					matchFound = true;
				}
			}
		}		
	}
	
	if (!checkMatch || (checkMatch && matchFound))
		eventarray.push(line);
  }
  
  var infoarr = [];
  
  for (var j in fieldKeys){
    infoarr.push(fieldNames[j]);			
  }
  
  infoarr.push(eventarray.length);
	
	sheet.getRange(2, 1, 1, infoarr.length).setValues([infoarr]);
  	sheet.getRange(1, 1, 2, infoarr.length).setBackground(calendarColor);
	sheet.setTabColor(calendarColor); 
	sheet.setCurrentCell(sheet.getRange('A3'));
  if (eventarray.length > 0){
    sheet.getRange(3, 1, eventarray.length, eventarray[0].length).setValues(eventarray);
  }
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
	var rowsToDetele = sheet.getMaxRows() - lastRow - 10;
	var colsToDetele = sheet.getMaxColumns() - lastCol;
	if (rowsToDetele > 0){
		sheet.deleteRows(lastRow + 11, rowsToDetele);
	}  
	if (colsToDetele > 0){
		sheet.deleteColumns(lastCol + 1, colsToDetele);
	}    
}
