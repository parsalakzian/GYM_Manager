/**
 * =================================================================
 *                        ** IMPORTANT **
 * This is a new, corrected version of the script. You MUST deploy
 * this new version to fix the bug. The old Web App URL will not work.
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
 *    - Give it a new description (e.g., "Gym Manager Backend v2").
 *    - For "Execute as", select "Me".
 *    - For "Who has access", select "Anyone".
 * 9. Click "Deploy".
 * 10. In the popup, click "Authorize access" (you may need to re-authorize).
 * 11. Choose your Google account, click "Advanced", then "Go to [Your Project Name] (unsafe)", and "Allow".
 * 12. After deployment, copy the new "Web app URL". This is the URL you need to provide to me.
 */

const SPREADSHEET_ID = "1PYPBDN1l4--LzbeLRPRb8DIvNcBdh0YrmfyFyN_0bBQ";
const SHEET_NAME = "Sheet1";

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const range = sheet.getRange("A1");
    const data = range.getValue();
    const content = data ? data.toString() : '{}';

    return ContentService.createTextOutput(content)
      .setMimeType(ContentService.MimeType.JSON)
      .withSuccessCode(200);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .withSuccessCode(500);
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
      .withSuccessCode(200);
  } catch (error) {
    const response = { status: "error", message: error.message };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON)
      .withSuccessCode(500);
  }
}

// Note: A `doOptions` function is often needed for preflight requests when dealing with CORS.
// However, Apps Script's ContentService handles simple POSTs from a browser context surprisingly well
// without it, as long as the client-side fetch doesn't use a complex Content-Type that triggers a preflight.
// The new fetch call will be structured to avoid this, making `doOptions` unnecessary for this use case.
// The key is that the client can now properly await the response from `doPost`.