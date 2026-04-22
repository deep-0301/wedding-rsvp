// Wedding RSVP — Google Apps Script
// Paste this into script.google.com and redeploy as Web App.
// It sends an email to you AND saves a copy to the sheet on every submission.

const NOTIFY_EMAIL = 'dkp99984@gmail.com';
const SHEET_NAME   = 'RSVPs';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const name       = data.full_name   || '(not provided)';
    const email      = data.email       || '(not provided)';
    const phone      = data.phone       || '(not provided)';
    const attendance = data.attendance  || '(not provided)';
    const guests     = data.guest_count != null ? data.guest_count : 'N/A';
    const message    = data.message     || '(no message)';
    const timestamp  = data.timestamp   || new Date().toLocaleString('en-CA');

    // ── Send email notification ──
    const subject = `New RSVP: ${name} — ${attendance}`;
    const body =
      `New RSVP received for Hinal & Deep's Wedding\n` +
      `${'─'.repeat(45)}\n\n` +
      `Name:        ${name}\n` +
      `Email:       ${email}\n` +
      `Phone:       ${phone}\n` +
      `Attendance:  ${attendance}\n` +
      `Guests:      ${guests}\n` +
      `Message:     ${message}\n\n` +
      `Submitted:   ${timestamp}\n`;

    MailApp.sendEmail({
      to:      NOTIFY_EMAIL,
      subject: subject,
      body:    body
    });

    // ── Also save to sheet ──
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Full Name', 'Email', 'Phone', 'Attendance', 'Guest Count', 'Message']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([timestamp, name, email, phone, attendance, guests, message]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput('RSVP endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
