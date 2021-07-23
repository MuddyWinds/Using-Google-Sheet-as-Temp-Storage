/** Open Google Form for storing user input */
function createSource() { 
  var widget = HtmlService.createHtmlOutputFromFile("importSource.html");
  widget.setWidth(800);
  widget.setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(widget, "Create New Source");  
}

/** Auto-sort Google Form result based on timestamp, then move inputs to "Source Data Files" */
function sortSourceReponse() {
  sform.sort(1, false);
  sform.getRange("A2:L2").clearFormat();

  /** Get user input in SForm, then assign source index */
  var [[sName, sURL, sDes, sUpdate, status, timeSheet, timeCell]] = sform.getRange("B2:H2").getValues();
  var nextIndex = "0000" + (parseInt((source_Storage[source_Storage.length-1]).substring(2,6)) + 1);
  var sIndex = "SD" + nextIndex.substr(-4);

  sources.appendRow([sIndex, "", sUpdate, sDes]);
  sources.getRange(sources.getLastRow(), 1, 1, sources.getLastColumn()).clearFormat();
  
  var updateTime = "Nil";
  if (checkSCondition(sName, sURL, status)) {
    var richValue = SpreadsheetApp.newRichTextValue().setText(sName).setLinkUrl(sURL).build();
    sources.getRange(sources.getLastRow(), 2).setRichTextValue(richValue); 

    /** Display Last Update Time if it is noted in the report */
    /** Extract the last update time from the same cell if the sheet is continuously updated. */
    /** Extract the last update time from the first sheet if a new sheet is appended every time */
    if (timeSheet != "" && timeCell != "" && SpreadsheetApp.openByUrl(sURL).getSheetByName(timeSheet)) {
      updateTime = SpreadsheetApp.openByUrl(sURL).getSheetByName(timeSheet).getRange(timeCell).getValue(); }
    else if (timeSheet == "00" && timeCell != "") {
      updateTime = SpreadsheetApp.openByUrl(sURL).getSheets()[0].getRange(timeCell).getValue(); }
  }
  else { sform.deleteRow(2); }
  sources.getRange(sources.getLastRow(),6).setValue(updateTime);

  /** If edit report, replaces the original user input */
  var sourceIDX = sIndex, sourceRow = sources.getLastRow();
  if (status == "Edit Source") {
    if (surl_Storage.indexOf(sURL) != -1) {
      sourceRow = surl_Storage.indexOf(sURL)+3;
      sourceIDX = sources.getRange("A"+sourceRow).getValue();

      sources.getRange(sources.getLastRow(),2,1,sources.getLastColumn()).moveTo(sources.getRange("B"+sourceRow));
      sources.deleteRow(sources.getLastRow());

      /** Set previous report row as red in "SForm" */
      var a = sform.getRange(3,3,sform.getLastRow()-2).getValues().flat().indexOf(sURL);
      if (a != -1)  { sform.getRange(a+3,1,1,9).setBackground("#fec3c3"); }
    }
    SpreadsheetApp.getActive().toast(sName + " is updated.", "✅ Action Successful:", 4);
  }

  /** Check raw data update status */
  if (sUpdate == "Auto") {
    if (updateTime != "Nil" && rawUpdate(sourceIDX,updateTime)) { 
      sources.getRange(sources.getLastRow(),5).setValue("Updated"); }
    else { sources.getRange(sources.getLastRow(),5).setValue("Not Updated").setBackground("#fec3c3"); }
  }

  /** Update Source Name / URL storage for later checking */
  source_Storage = sources.getRange(3, 1, sources.getLastRow()-2).getValues().flat();
  surl_Storage = sources.getRange(3, 2, sources.getLastRow()-2).getRichTextValues().map(r => r[0].getLinkUrl());

  /** Set text wrapping strategy as Clip in "Form" */
  sform.getRange("C2").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

/** Check whether source name / URL is duplicated */
function checkSCondition(sourceName, url, status) {
  var noError = true;
  if (status == "Create Source") {
    if (source_Storage.indexOf(sourceName) != -1) {
      noError = false;
      sources.getRange(sources.getLastRow(), 2).setBackground("#fec3c3").setValue("Duplicated Source Name.");
    }
    if (surl_Storage.indexOf(url) != -1) {
      noError = false;
      sources.getRange(sources.getLastRow(), 2).setBackground("#fec3c3").setValue("Duplicated Source URL.");
    }
  } return noError;
}

/** This function returns all sources to UI */
function returnSources()  { return sources.getRange("A3:B"+sources.getLastRow()).getValues(); }


/** Generate the google form URL with pre-filled source info */
function editSourcePrefill(sourceIDX) { 
  /** Check if source index exists in "Data source files" */
  if (source_Storage.indexOf(sourceIDX) != -1) {
    SpreadsheetApp.getActive().toast(sourceIDX + "  exists.", "✅ Action Successful:", 4); 

    /** Search and extract orginal info of that source */
    var sourceRow = sform.getRange(1,3,sform.getLastRow()).getValues().flat();
    return sourceRow.map(function(surl,i) {
      if (surl_Storage[source_Storage.indexOf(sourceIDX)] == sourceRow[i] && sform.getRange(i+1,3).getBackground() != "#fec3c3") {
        var [[sName,sURL,sDes,sUpdate,status,timeSheet,timeCell]] = sform.getRange(i+1,2,1,7).getValues();

        if (sDes != "")         { sDes = "[ PLEASE INPUT FORM FIELD ID ]" + sDes.replace(/\s/g,"+");            }
        if (timeSheet != "")    { timeSheet = "[ PLEASE INPUT FORM FIELD ID ]" + timeSheet.replace(/\s/g,"+");  }
        if (timeCell != "")     { timeCell = "[ PLEASE INPUT FORM FIELD ID ]" + timeCell.replace(/\s/g,"+");   }

        sformURL  = "[ PLEASE INPUT YOUR GOOGLE FORM URL ]";
        sformData = "[ PLEASE INPUT FORM FIELD ID ]" + sName.replace(/\s/g,"+") + "[ PLEASE INPUT FORM FIELD ID ]" + encodeURIComponent(sURL) + sDes
                    + timeSheet + timeCell + "[ PLEASE INPUT FORM FIELD ID ] =Edit+Source";

        return (sformURL + sformData); } 
      else { return null; }
    }).filter(value => value != null)[0];
  } SpreadsheetApp.getActive().toast(sourceIDX + " does not exist.", "❌ Action Unsuccessful:", 4);
    return null;
}

function editSource() {
  var widget = HtmlService.createHtmlOutputFromFile("editSource");
  widget.setWidth(800); 
  widget.setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(widget, "Edit Existing Source");
}


/** 
 * resetName function is triggered every time when the sheet is opened.
 * Report name and last update date will be updated automatically.
 */
function resetSName() {
  var timeSheet = [];

  surl_Storage.forEach(function(surl,num) {
    var updateTime = "Nil";
    if (surl != null) {
      /** Update source's name  */
      var sourceURL = sources.getRange(num+3, 2).getRichTextValue().getLinkUrl();
      if (sourceURL.indexOf("http") != -1)  { var sourceName = SpreadsheetApp.openByUrl(sourceURL).getName();}
      sources.getRange(num+3, 2).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(sourceName).setLinkUrl(sourceURL).build());

      /** Update source's last update date */
      var a = sform.getRange(2,3,sform.getLastRow()-1).getValues().flat().map(function(surl2,i) {
        if (sform.getRange(i+2,1).getBackground() !=  "#fec3c3" && surl == surl2)  { return i + 2; }
        else  { return null; }
      }).filter(value => value != null)[0];

      if (a) {
        var timeAll = sform.getRange("B"+a+":I"+a).getValues().flat(); 
        if (timeAll[5] != "" && timeAll[6] != "" && SpreadsheetApp.openByUrl(surl).getSheetByName(timeAll[5])) {
          updateTime = SpreadsheetApp.openByUrl(surl).getSheetByName(timeAll[5]).getRange(timeAll[6]).getValue(); }
        else if (timeAll[5] == "00" && timeAll[6] != "") {
          updateTime = SpreadsheetApp.openByUrl(surl).getSheets()[0].getRange(timeAll[6]).getValue(); }
      }
    } timeSheet.push(new Array(updateTime));
  }); sources.getRange(3,6,sources.getLastRow()-2).setValues(timeSheet).setHorizontalAlignment("left");
}

function autoSchedule() {
  source_Storage.forEach(function(sourceIndex, i) {
    var sourceRow = sources.getRange(i+3,1,1,6).getValues().flat();
    if (sourceRow[2] == "Auto" && sourceRow[5] != "Nil" && rawUpdate(sourceIndex, sources.getRange(i+3,6).getValue())) {
      sources.getRange(i+3,5).setValue("Updated").clearFormat(); }
    else { sources.getRange(i+3,5).setValue("Not Updated").setBackground("#fec3c3"); }
  })
}

/** Source: https://developers.google.com/apps-script/reference/calendar/calendar-event#getendtime */
function rawUpdate(sourceIndex, lastUpdate) {
  /** Schedule update time > Last update time => Not updated */
  lastUpdate = new Date(lastUpdate.substring(0,19));
  var past100days = new Date(new Date().getTime() - 100*(24*3600*1000));
  var today       = new Date(new Date().getTime() + 1000 );

  var event = CalendarApp.getCalendarById("[ PLEASE INPUT YOUR GOOGLE CALENDAR ID ]")
               .getEvents(past100days, today, { search:sourceIndex });

  if (event.length && event[event.length-1].getEndTime() <= lastUpdate)  { return true;  }
  else                                                                   { return false; }
}

