/**
 * TessereX "Get In Touch" backend (Google Apps Script)
 * Paste into Extensions > Apps Script of your "Website Contact Submissions" Google Sheet,
 * then Deploy > Manage deployments > Edit > Version: New version > Deploy.
 * Execute as: Me.  Who has access: Anyone.
 */
var RECIPIENT = "contact.tesserex@gmail.com";   // where the notification emails go
var MIN_FILL_MS = 2500;                         // forms sent faster than this are treated as bots

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = e.parameter;

    // 1. Honeypot: real visitors never see this field. If it is filled, it is a bot.
    //    Reply "success" so the bot learns nothing, but save and send nothing.
    if (data.website) return ok_();

    // 2. Security check (simple sum question) and minimum fill time.
    var a = parseInt(data.cap_a, 10), b = parseInt(data.cap_b, 10), ans = parseInt(data.cap_answer, 10);
    if (isNaN(a) || isNaN(b) || a + b !== ans) return ok_();
    if (Number(data.elapsed) < MIN_FILL_MS) return ok_();

    // 3. Basic checks and length limits.
    var name = String(data.name || "").slice(0, 120);
    var email = String(data.email || "").slice(0, 200);
    var message = String(data.message || "").slice(0, 5000);
    var subject = String(data.subject || ("New Contact Form Message from " + name)).slice(0, 250);
    if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !message) {
      return ContentService.createTextOutput(JSON.stringify({ result: "error", error: "Invalid input" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. Save to the sheet (backup).
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([new Date(), name, email, subject, message]);

    // 5. Email notification; "Reply" in Gmail goes straight to the visitor.
    var body = "Name: " + name + "\nEmail: " + email + "\n\nMessage:\n" + message;
    MailApp.sendEmail(RECIPIENT, subject, body, { replyTo: email });

    return ok_();
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
