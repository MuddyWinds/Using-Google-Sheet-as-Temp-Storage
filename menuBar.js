/** Global variables declared here */
var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
var form = activeSheet.getSheetByName("Form");
var content = activeSheet.getSheetByName("Content");
var sources = activeSheet.getSheetByName("Source Data Files");
var sform = activeSheet.getSheetByName("SForm");
var firstForm = form.getRange("A2").getValue().toString();


/** Store all report names & URL here. */ 
var report_Storage = [], url_Storage = [], source_Storage = [], surl_Storage = [];
if (!content.getRange("A3:K3").isBlank()) {
  report_Storage = content.getRange(3, 4, content.getLastRow()-2).getValues().flat();
  url_Storage = content.getRange(3, 4, content.getLastRow()-2).getRichTextValues().map( r => r[0].getLinkUrl());
}
if (!sources.getRange("A3:K3").isBlank()) {
  source_Storage = sources.getRange(3, 1, sources.getLastRow()-2).getValues().flat();
  surl_Storage = sources.getRange(3, 2, sources.getLastRow()-2).getRichTextValues().map( r => r[0].getLinkUrl());  
}


/**
 * Source: https://spreadsheet.dev/custom-menus-in-google-sheets
 * Menus are added to menu bar. Function is triggered every time when the sheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📄Document")
    .addItem("📝 Create Report", "createReport")
    .addItem("✍️ Edit Report", "editReport")
    .addSeparator()
    .addItem("🔗 Create Source","createSource")
    .addItem("✍️ Edit Source", "editSource")
    .addToUi();
  
  SpreadsheetApp.getUi()
    .createMenu("🔍Advanced")
    .addItem("🏴󠁧󠁢󠁥󠁮󠁧󠁿 Update Timestamp","updateVersion")
    .addItem("🔙 Version History","versionHist")
    // .addItem("🚫 Delete Invalid Inputs","deleteInvalid")
    .addItem("💻 Hide Invalid Inputs","hideInvalid")
    .addItem("🛠️ Show All Reports","unhideInvalid")
    .addItem("📓 User Guide","userGuide")
    .addToUi();
}
