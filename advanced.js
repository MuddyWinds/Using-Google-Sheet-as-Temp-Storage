/** Source: https://stackoverflow.com/questions/2388115/get-locale-short-date-format-using-javascript */
/** Source: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Intl/DateTimeFormat */
/** Advanced option: Update timestamp. Format: 2000-06-31*/
function updateVersion() {
  var sheet = content, cell = "E1";
  sheet.getRange(cell).setValue("Last Update Date: " + new Intl.DateTimeFormat('se-SE').format(new Date()));
  SpreadsheetApp.getActive().toast("Timestamp is updated.", "✅ Action Successful:", 4);
}

/** Advanced option: Display all version History of a report */
function versionHist() {
  var widget = HtmlService.createHtmlOutputFromFile("vHist");
  widget.setWidth(800);
  widget.setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(widget, "Show Version History");  
}

/** Called by vHist.html to extract all version info of a report */
function displayR(name) {
  var reportURL = form.getRange("H1:H"+form.getLastRow()).getValues().flat();
  return reportURL.map(function(url,i) {
    if (url.includes("http") && name == SpreadsheetApp.openByUrl(url).getName()) {
      var getRow = form.getRange("A"+(i+1)+":L"+(i+1)).getValues().flat();
      var timeStamp = getRow[0].toLocaleDateString() + " " + getRow[0].toLocaleTimeString() + "&nbsp &nbsp &nbsp &nbsp";
      // var timeStamp = getRow[0].toString().substring(4,21) + "&nbsp &nbsp &nbsp &nbsp";
      var create_or_edit = getRow[10];
      if (create_or_edit == "Create Report")     { create_or_edit += "&nbsp &nbsp &nbsp &nbsp &nbsp"; }
      else                                       { create_or_edit += "&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp"; }
      return (timeStamp + create_or_edit + getRow[11]);
    } 
    else { return null; }
  }).filter(value => value != null).join("\n");
}

/** Advanced function: Delete all invalid inputs. */
function deleteInvalid() {
  var deleteR = false, deleteCount = 0;
  var reportColor = content.getRange(1,4,content.getLastRow(),5).getBackgrounds();
  reportColor.map(x => x[0]).map(function(name,i) {
    if (name =="#fec3c3" || reportColor.map(x => x[4])[i] == "#fec3c3") { 
      content.deleteRow(i+1-deleteCount);
      deleteR = true, deleteCount++; }
  });
  if (deleteR)  { SpreadsheetApp.getActive().toast(deleteCount + " report(s) deleted", "✅ Action Successful:", 4); }
  else          { SpreadsheetApp.getActive().toast("No reports are deleted.","❌ Action Unsuccessful:", 4);         }
}


/** Advanced function: Hide all invalid inputs. */
function hideInvalid() {
  hideR = false, hideCount = 0;
  var reportColor = content.getRange(1,4,content.getLastRow(),5).getBackgrounds();
  reportColor.map(x => x[0]).map(function(name,i) {
    if (name =="#fec3c3" || reportColor.map(x => x[4])[i] == "#fec3c3") { 
      content.hideRows(i+1);
      hideR = true, hideCount++; }
  });
  if (hideR) { SpreadsheetApp.getActive().toast(hideCount + " report(s) hidden", "✅ Action Successful:", 4);  }
  else       { SpreadsheetApp.getActive().toast("No reports are hidden.","❌ Action Unsuccessful:", 4);        }
}


/** Advanced function: Unhide all invalid inputs */
function unhideInvalid() {
  content.showRows(1,content.getLastRow());
  SpreadsheetApp.getActive().toast("All reports are unhidden", "✅ Action Successful:", 4);
}

/** Advanced function: User Guide */
function userGuide() {
  var widget = HtmlService.createHtmlOutputFromFile("guide");
  widget.setWidth(800);
  widget.setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(widget, "Complete User Guide");  
}

/** Send automatic email to teammates if the source update fails */
function sendEmail() {
  var teamlist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team List");
  var emails = teamlist.getSheetValues(2,1,teamlist.getLastRow()-1,2);
  var failedSources = sources.getSheetValues(3,1,sources.getLastRow()-2,5).map(function(sourceRow, i) {
    if (sourceRow[4] == "Not Updated")   { return (sourceRow[0] + " " + sourceRow[1]); }
    else { return null; }
  }).filter(value => value != null).join("<br>");

  emails.forEach(function(email, i) {
    /** Prevent sending depulicated emails to teammates */
    if (email[1] != "EMAIL SENT") {
      MailApp.sendEmail(email[0],"Failed Executions of Raw Data Update", "",
      { "htmlBody" : "<strong> This is an automatic email scheduled to be sent every Monday 9:00a.m.</strong> <br><br>" 
      + failedSources });
      teamlist.getRange("B"+(i+2)).setValue("EMAIL SENT");
      /** Make sure the cell is updated right away in case the script is interrupted */ 
      SpreadsheetApp.flush();
    }
  });

  /** Reset the status for next scheduled email */
  teamlist.getRange("B2:B"+teamlist.getLastRow()).clearContent();
  teamlist.getRange("A2:B"+teamlist.getLastRow()).clearFormat().setHorizontalAlignment("center");
}

// USER GUIDE DEPENDS ON USERS' NEEDS
/** Called by guide.html to display the guidance message */
function displayGuide(clicked_id) {
  switch (clicked_id) {
    case "01":
      var text = ""
      return text;
    case "02":
      return;
    case "03":
      return;
    case "04":
      return;
    case "05":
      return;
    case "06":
      return;
  }
}
