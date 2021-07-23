/** Determine whether the submitted form is for report or source */
function checkStatus() { 
  form.sort(1, false);
  if (form.getRange("A2").getValue().toString() != firstForm)     { sortResponses();     }
  else                                                            { sortSourceReponse(); }
  firstForm = form.getRange("A2").getValue().toString();
}

/** Open Google Form for storing user input */
function createReport() {
  var widget = HtmlService.createHtmlOutputFromFile("importForm.html");
  widget.setWidth(800);
  widget.setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(widget, "Create New Report");  
}

/** 
 * Source: https://spreadsheet.dev/how-to-automatically-sort-google-form-responses
 * Source: https://spreadsheet.dev/add-links-to-a-cell-in-google-sheets-using-apps-script
 * Source: https://www.august.com.au/blog/how-to-send-slack-alerts-from-google-sheets-apps-script/
 * Auto-sort Google Form result based on timestamp, then move inputs to "Content"
 */
function sortResponses() {
  form.getRange("A2:L2").clearFormat();
  var [[docType,cate,update,des,index,method,url,timeSheet,timeCell,status]] = form.getRange("B2:K2").getValues();

  var newRow = [[docType, cate, update, "", des, "", index, "", method]];
  content.insertRowBefore(3);
  content.getRange("3:3").clearFormat();
  content.getRange("A3:I3").setValues(newRow);

  /** If url is valid, then the hyperlink will be created in "Content" */
  var noError = true, reportName = "", updateTime = "Nil", reportError = "";
  if (status == "Create Report") {
    if (url_Storage.indexOf(url) != -1) {
      noError = false;
      reportError += "URL already exists \n"; 
    }
    if (url.indexOf("http") == -1) {
      noError = false;
      reportError += "Invalid URL \n";
    }
    if (url.indexOf(SpreadsheetApp.getActiveSpreadsheet().getUrl().substring(39,70)) != -1) {
      noError = false;
      reportError += "URL of current spreadsheet cannot be added";
    }
  }

  if (noError) {
    /** Attach URL to report name if it is valid */
    externalReport = SpreadsheetApp.openByUrl(url);
    reportName = externalReport.getName();
    var richValue = SpreadsheetApp.newRichTextValue().setText(reportName).setLinkUrl(url).build();
    content.getRange(3, 4).setRichTextValue(richValue);

    /** Display Last Update Time if it is noted in the report */
    if (timeSheet != "" && timeCell != "" && externalReport.getSheetByName(timeSheet)) {
      updateTime = externalReport.getSheetByName(timeSheet).getRange(timeCell).getValue().substring(0,10); }
  }

  else {
    form.deleteRow(2);
    content.getRange(3, 4).setBackground("#fec3c3").setValue(reportError); 
  }

  content.getRange(3, 6).setValue(updateTime).setHorizontalAlignment("left");
  sourceLinks(reportName, index.replace(/\s+/g,"\n"));

  /** If edit report, replaces the original user input */
  if (status == "Edit Report") {
    if (url_Storage.indexOf(url) != -1) {
      content.getRange("A3:K3").moveTo(content.getRange("A"+(4+url_Storage.indexOf(url))));
      content.deleteRow(3);

      /** Set previous report row as red in "Form" */
      var a = form.getRange(3,8,form.getLastRow()-2).getValues().flat().indexOf(url);
      if (a != -1)  { form.getRange(a+3,1,1,11).setBackground("#fec3c3"); } 
    }
    SpreadsheetApp.getActive().toast(reportName + " is updated.", "✅ Action Successful:", 4);
  }

  /** Update URL storage for later checking */
  url_Storage = content.getRange(3, 4, content.getLastRow()-2).getRichTextValues().map(r => r[0].getLinkUrl());

  /** Set text wrapping strategy as Clip in "Form" */
  form.getRange("H2").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}


/** 
 * Create hyperlinks for source Index: Position in "Source Data Files"
 * Create hyperlinks for source Names: Source spreadsheet (Applicable to both external and internal links)
 */
function sourceLinks(reportName, sourceIndex) {
  var invalidResult = false, sourceError = "";
  var lastSourceIndex = parseInt((sources.getRange(sources.getLastRow(), 1).getValue()).substring(2,6));
  var sourceLocation = "[ PLEASE INPUT THE CELL LOCATION. FORMAT: #gid=XXXXXXXX&range=A]";
  var richValue = SpreadsheetApp.newRichTextValue().setText(sourceIndex);
  var indexArray = sourceIndex.split(/\s+/g);

  /** Check whether source index is valid or not */
  if ((indexArray.filter(iDX => iDX.substring(0,2) != "SD" || iDX.length != 6 || isNaN(iDX.substring(2,6)))).length) {
    sourceError += "Invalid Source Index \n";
    invalidResult = true; }

  if ((indexArray.filter(iDX => !isNaN(iDX.substring(2,6)) && parseInt(iDX.substring(2,6)) > lastSourceIndex)).length) {
    sourceError += "Source Index Out of Range \n";
    invalidResult = true; }

  if ((indexArray.filter((iDX, temp) => indexArray.indexOf(iDX) != temp)).length) {
    sourceError += "Duplicated Source Index \n";
    invalidResult = true; }

  /** Check whether the required source indexes exist in "Source Data Files" */
  if (indexArray.map(sourceIx => source_Storage.indexOf(sourceIx) == -1).length) {
    sourceError += "Source Index does not exist \n";
    invalidResult = true; }

  if (invalidResult) { 
    content.getRange(3, 8).setBackground("#fec3c3").setValue(sourceError);
    return; }

  
  /** Locate the source index in "Source Data Files", then store source name in indexName */
  /** Attach cell location link for each source index */
  /** Append report name to the last row of its corresponding sources */
  var indexName = indexArray.map(function (iDX,i) {
    var temp = source_Storage.indexOf(iDX);
    if (temp != -1) {
      richValue.setLinkUrl(i+i*6, 6+(i*7), sourceLocation+(temp+3));
      if (reportName != "") { 
        var cellInfo = sources.getRange(temp+3,8);
        if (cellInfo.isBlank()) { cellInfo.setValue(reportName); }
        else                    { cellInfo.setValue(cellInfo.getValue() + "\n" + reportName); }
      }
      return sources.getRange((temp + 3),2).getValue(); }
    else { return ""; }
  });
  content.getRange(3, 7).setRichTextValue(richValue.build());

  /** Locate the source URL in "Source Data Files", then store source URL in indexURL */
  var indexURL = indexArray.map(function (iDX) {
    if (source_Storage.indexOf(iDX) != -1) {
      return sources.getRange(source_Storage.indexOf(iDX) + 3,2).getRichTextValue().getLinkUrl(); }
    else { return null; }
  });

  /** Attach URL to each source full name */
  var sourceNames = indexName.join("\n");
  var richValue2 = SpreadsheetApp.newRichTextValue().setText(sourceNames);

  indexName.forEach(function(name,i) {
    var start_substring = sourceNames.indexOf(name);
    if (i+1 != indexName.length) {
      var end_substring = sourceNames.indexOf(indexName[i+1]) - 1; }
    else {
      var end_substring = sourceNames.length; }
    richValue2.setLinkUrl(start_substring, end_substring, indexURL[i]);
  });
  content.getRange(3, 8).setRichTextValue(richValue2.build());
}


/** 
 * resetName function is triggered every time when the sheet is opened.
 * Report name and last update date will be updated automatically.
 */
// Last minute buffer : Regenerate all reports + regenerate URL
function resetName() {
  var timeSheet = [];

  url_Storage.forEach(function(url,num) {
    var updateTime = "Nil";
    if (url != null) {
      /** Update report's name. Prevent the user fils the current spreadsheet as a new report */
      var reportURL = content.getRange(num+3, 4).getRichTextValue().getLinkUrl();
      if (reportURL.indexOf("http") != -1)  { var reportName = SpreadsheetApp.openByUrl(reportURL).getName(); }
      content.getRange(num+3, 4).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(reportName).setLinkUrl(reportURL).build()); 

      /** Update report's last update date */
      var a = form.getRange(2,8,form.getLastRow()-1).getValues().flat().map(function(url2,i) {
        if (form.getRange(i+2,1).getBackground() != "#fec3c3" && url == url2) { return i+2 ; }
        else { return null; }
      }).filter(value => value != null)[0];

      if (a) {
        var timeAll = form.getRange("B"+a+":J"+a).getValues().flat();
        if (timeAll[7] != "" && timeAll[8] != "" && SpreadsheetApp.openByUrl(url).getSheetByName(timeAll[7])) {
          updateTime = SpreadsheetApp.openByUrl(url).getSheetByName(timeAll[7]).getRange(timeAll[8]).getValue().substring(0,10); }

        /** Update DocType, Category and  Report Schedule */
        content.getRange("A"+(num+3)+":C"+(num+3)).setValues(new Array(timeAll.slice(0,3)));
      }
    } timeSheet.push(new Array(updateTime));
  }); content.getRange(3,6,content.getLastRow()-2,1).setValues(timeSheet).setHorizontalAlignment("left");
}

function editReport() {
  var widget = HtmlService.createHtmlOutputFromFile("editForm");
  widget.setWidth(800); 
  widget.setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(widget, "Edit Existing Report");
}


/** This function returns the report names to editForm */
function returnReports() { return report_Storage; }


/** Generate the google form URL with pre-filled report info */
function editReportPrefill(reportName) {
  /** Check if report name exists in "Content" */
  if (report_Storage.indexOf(reportName) != -1)  {
    SpreadsheetApp.getActive().toast(reportName + " and its URL exist.", "✅ Action Successful:", 4);

    /** Search and extract original info of that report */
    var reportRow = form.getRange(1,8,form.getLastRow()).getValues().flat();
    return reportRow.map(function(url,i) {
      if (url_Storage[report_Storage.indexOf(reportName)] == reportRow[i] && form.getRange(i+1,8).getBackground() != "#fec3c3") {
        var [[docType,cate,update,des,index,method,url,timeSheet,timeCell]] = form.getRange(i+1,2,1,9).getValues();

        if (des != "")         { des = "[ PLEASE INPUT FORM FIELD ID ]" + des.replace(/\s/g,"+");            }
        if (timeSheet != "")   { timeSheet = "[ PLEASE INPUT FORM FIELD ID ]" + timeSheet.replace(/\s/g,"+"); }
        if (timeCell != "")    { timeCell = "[ PLEASE INPUT FORM FIELD ID ]" + timeCell.replace(/\s/g,"+");  }

        formURL  = "[ PLEASE INPUT YOUR GOOGLE FORM URL ]";
        formData = "[ PLEASE INPUT FORM FIELD ID ]" + docType.replace(/\s/g,"+") + "[ PLEASE INPUT FORM FIELD ID ]" + cate.replace(/\s/g,"+") +
                   "[ PLEASE INPUT FORM FIELD ID ]" + encodeURIComponent(url) + "[ PLEASE INPUT FORM FIELD ID ]" + update.replace(/\s/g,"+") +
                   des + timeSheet + timeCell + "[ PLEASE INPUT FORM FIELD ID ]" + index.replace(/\s/g,"+") +
                   "[ PLEASE INPUT FORM FIELD ID ]" + method.replace(/\s/g,"+") + "[ PLEASE INPUT FORM FIELD ID ] =Edit+Report"; 
        return (formURL + formData); } 
      else { return null; }
    }).filter(value => value != null)[0];
  } SpreadsheetApp.getActive().toast(reportName + " does not exist.", "❌ Action Unsuccessful:", 4);
    return null;
}

