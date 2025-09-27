/**
 * =================================================================
 *              ** FINAL, CORRECTED SCRIPT V2 **
 * This is the final, corrected version of the script. You MUST
 * deploy this new version to fix all communication errors.
 * =================================================================
 *
 * To use this script:
 * 1. Open your Google Sheet.
 * 2. Go to Extensions > Apps Script.
 * 3. Delete all old code in the `Code.gs` file and paste this entire script.
 * 4. Save the project.
 * 5. Click "Deploy" > "New deployment".
 * 7. For "Select type", choose "Web app".
 * 8. In the "Configuration" section:
 *    - Give it a new description (e.g., "Gym Manager Backend FINAL").
 *    - For "Execute as", select "Me".
 *    - For "Who has access", select "Anyone".
 * 9. Click "Deploy".
 * 10. In the popup, click "Authorize access" (you may need to re-authorize).
 * 11. Choose your Google account, click "Advanced", then "Go to [Your Project Name] (unsafe)", and "Allow".
 * 12. After deployment, copy the new "Web app URL". You will need this for the final step.
 */

const SPREADSHEET_ID = "1PYPBDN1l4--LzbeLRPRb8DIvNcBdh0YrmfyFyN_0bBQ";
const SHEET_NAME = "Sheet1";

// This function is required to handle the CORS preflight request that the browser sends
// when the frontend makes a POST request with a 'Content-Type' header.
function doOptions(e) {
  return ContentService.createTextOutput()
    .addHeader('Access-Control-Allow-Origin', '*')
    .addHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .addHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const range = sheet.getRange("A1");
    const data = range.getValue();
    const content = data ? data.toString() : '{}';

    return ContentService.createTextOutput(content)
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader('Access-Control-Allow-Origin', '*');
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader('Access-Control-Allow-Origin', '*');
  }
}

function doPost(e) {
  try {
    const masterState = e.postData.contents;
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const range = sheet.getRange("A1");
    range.setValue(masterState);

    const response = { status: "success", message: "Data saved successfully." };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader('Access-control-Allow-Origin', '*');
  } catch (error) {
    const response = { status: "error", message: error.message };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader('Access-Control-Allow-Origin', '*');
  }
}