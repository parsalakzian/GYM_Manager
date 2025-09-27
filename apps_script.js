/**
 * To use this script:
 * 1. Open your Google Sheet.
 * 2. Go to Extensions > Apps Script.
 * 3. Delete any code in the `Code.gs` file and paste this entire script.
 * 4. IMPORTANT: Replace "YOUR_SPREADSHEET_ID_HERE" with the actual ID of your spreadsheet.
 *    You can find the ID in your sheet's URL (e.g., in "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit", it's the long string of characters).
 * 5. Save the project.
 * 6. Click "Deploy" > "New deployment".
 * 7. For "Select type", choose "Web app".
 * 8. In the "Configuration" section:
 *    - Give it a description (e.g., "Gym Manager Backend").
 *    - For "Execute as", select "Me".
 *    - For "Who has access", select "Anyone" (this is necessary for the app to call it, but only you can execute it).
 * 9. Click "Deploy".
 * 10. In the popup, click "Authorize access".
 * 11. Choose your Google account. You may see a "Google hasn't verified this app" warning. Click "Advanced", then "Go to [Your Project Name] (unsafe)".
 * 12. Click "Allow" to grant the script permission to access your spreadsheets.
 * 13. After deployment, copy the "Web app URL". This is the URL you need to provide to me.
 */

// !!! IMPORTANT !!!
// Replace this with your actual Google Sheet ID
const SPREADSHEET_ID = "1PYPBDN1l4--LzbeLRPRb8DIvNcBdh0YrmfyFyN_0bBQ";
const SHEET_NAME = "Sheet1"; // This is the default name. Change it if your sheet has a different name.

/**
 * Handles GET requests to the web app.
 * Reads the entire master state from cell A1 of the specified sheet.
 */
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
    }
    const range = sheet.getRange("A1");
    const data = range.getValue();

    // If the cell is empty, return an empty JSON object string.
    const content = data ? data.toString() : '{}';

    return ContentService.createTextOutput(content)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("Error in doGet:", error);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .setStatusCode(500);
  }
}

/**
 * Handles POST requests to the web app.
 * Saves the entire master state to cell A1 of the specified sheet.
 */
function doPost(e) {
  try {
    const masterState = e.postData.contents;

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
    }
    const range = sheet.getRange("A1");
    range.setValue(masterState);

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data saved successfully." }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("Error in doPost:", error);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .setStatusCode(500);
  }
}