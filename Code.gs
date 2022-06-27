/*
Contains some useful examples of:
1. Displaying a user prompt to get input
2. Writing a file to a  Google Drive folder
3. Reading in the file contents to a string
4. Passing the HTML string to HtmlService to create an output page.
*/

/**
 * Serve a HTML file stored in Google Drive as a HtmlService page.
 * Prompt for a file ID (the file ID from the shareable link and NOT the link from the preview),
 * read in its HTML as a string and display it.
 * IDs to try: 1L2tmKIkjEaYP_KBiKhuuV6vVlQIr7AiD 182F2dUqSA9JMAGc0U1C6JD9nv1E1Z9mQ
 * 
 * @return {void}
 */
function displayGuiFromDriveFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const htmlFileID = Browser.inputBox("Enter a HTML file ID", Browser.Buttons.OK_CANCEL); //"1L2tmKIkjEaYP_KBiKhuuV6vVlQIr7AiD"; // 182F2dUqSA9JMAGc0U1C6JD9nv1E1Z9mQ
  const file = DriveApp.getFileById(htmlFileID);
  const htmlText = file.getBlob().getDataAsString();
  const tablePage = HtmlService.createHtmlOutput(htmlText);
  tablePage.setWidth(600);
  tablePage.setHeight(400);
  tablePage.setTitle("Table Demo");
  ss.show(tablePage);
}

/**
 * Serve a HTML file from a a Google Apps Script source HTML file using a hard-coded file name.
 * The function displayGuiFromDriveFile() is a more flexible option.
 * 
 * @return {void}
 */
function displayGuiFromGasFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const html = HtmlService.createHtmlOutputFromFile('emp');
  html.setWidth(600);
  html.setHeight(400);
  html.setTitle("Table Demo");
  ss.show(html);
}

/**
 * Write the Sheets data in a hard-coded range name to a file.
 * Used as a one-off for testing.
 * 
 * @return {void}
 */
function createAdvancedFormattingHtml() {
  const rngName = "ADV_FORMAT";
  writeHtmlFromSheetToFile(rngName);
}

/**
 * Generate a HTML file from an input range and save the output to a specific folder using a hard-coded
 * folder ID. 
 * 
 * @param {string} rngName - A named range from the active spreadsheet.
 * 
 * @return {void}
 */
function writeHtmlFromSheetToFile(rngName) {
  const ss = SpreadsheetApp.getActive();
  const rng = ss.getRangeByName(rngName);
  const shName = rng.getSheet().getSheetName();
  const htmlGen = new HtmlGenerator(rng);
  const tableHtml = htmlGen.FullPageHtml;
  const htmlFileName = `${shName}.html`;
  const htmlFolder = DriveApp.getFolderById("10032EdsyUBh7uzUQ4Xzr35N-pUJEM3tG");
  htmlFolder.createFile(htmlFileName, tableHtml, MimeType.HTML)
}
