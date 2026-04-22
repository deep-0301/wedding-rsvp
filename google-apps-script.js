// Wedding RSVP — Google Apps Script
// Paste this entire file into script.google.com, then deploy as a Web App.

const SHEET_NAME = 'RSVPs';

function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // Create the sheet with headers on first run
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Timestamp',
        'Full Name',
        'Email',
        'Phone',
        'Attendance',
        'Guest Count',
        'Message'
      ]);
      // Bold the header row
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.timestamp   || new Date().toLocaleString('en-CA'),
      data.full_name   || '',
      data.email       || '',
      data.phone       || '',
      data.attendance  || '',
      data.guest_count || 0,
      data.message     || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: lets you test the endpoint from a browser (GET request)
function doGet() {
  return ContentService
    .createTextOutput('RSVP endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
