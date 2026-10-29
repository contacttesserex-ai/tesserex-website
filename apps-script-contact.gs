/**
 * TessereX "Get In Touch" backend (Google Apps Script)  -  VERSION v4
 *
 * HOW TO INSTALL / UPDATE  (all three steps are needed):
 *  1. In your "Website Contact Submissions" Google Sheet open Extensions > Apps Script.
 *     Select ALL the existing code, delete it, paste this whole file, press Save.
 *  2. Deploy > Manage deployments > pencil (Edit) icon > Version: "New version" > Deploy.
 *     (Pressing only Save does NOT update the live web app. The old code keeps running until you do this.)
 *  3. Open your web app URL (the one ending in /exec) in a browser. You must see
 *     {"status":"ok","service":"TessereX contact form","version":"v4"}.
 *     If you see "Script function not found: doGet" the old code is still live.
 *
 * ONE-TIME CLEAN-UP (optional): choose the function  setupSheet  in the editor toolbar and press Run.
 * It formats the header row and converts rows saved by the OLD script into the new columns.
 *
 * Email subject : TessereX Website - <Enquiry type> - <Subject typed by the visitor>
 * Email body    : Enquiry type, Name, Email, Message (the subject is not repeated in the body)
 * Sheet (first tab): Sr. No. | Timestamp | Enquiry Type | Subject | Name | Email | Message
 *                 (File > Download > Microsoft Excel (.xlsx) gives an Excel copy)
 * Spam          : submissions stopped by the spam checks are listed in the tab "Blocked" with the reason.
 */
var VERSION = "v4";
var RECIPIENT = "contact.tesserex@gmail.com";            // where notification emails go
var SUBJECT_FORMAT = "TessereX Website - {type} - {subject}";
var TYPES = ["Partner Enquiries", "Investor Relations", "Web Message"];   // keep in step with the website's enquiry types
var DEFAULT_TYPE = "Web Message";
var SHEET_NAME = "";                                     // "" = the FIRST tab of the spreadsheet (where your headers are)
var BLOCKED_NAME = "Blocked";
var MIN_FILL_MS = 1500;                                  // forms sent faster than this are treated as bots
var HEADERS = ["Sr. No.", "Timestamp", "Enquiry Type", "Subject", "Name", "Email", "Message"];
var DATE_FORMAT = "dd-mmm-yyyy hh:mm:ss";

function json_(obj) {
  obj.version = VERSION;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
// Stops text starting with = + - @ from being treated as a spreadsheet formula.
function safe_(v) {
  v = String(v == null ? "" : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function format_(sh) {
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#e8eefc");
  sh.getRange(2, 2, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat(DATE_FORMAT);
  sh.getRange(2, 7, Math.max(sh.getMaxRows() - 1, 1), 1).setWrap(true);
  sh.setColumnWidth(1, 70); sh.setColumnWidth(2, 160); sh.setColumnWidth(3, 150); sh.setColumnWidth(4, 260);
  sh.setColumnWidth(5, 160); sh.setColumnWidth(6, 220); sh.setColumnWidth(7, 420);
}
// Gets the submissions tab (the first tab unless SHEET_NAME is set); writes the header row if it is empty.
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = SHEET_NAME ? (ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME)) : ss.getSheets()[0];
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); format_(sh); }
  return sh;
}
// Records a submission that the spam checks stopped, so nothing disappears silently.
function blocked_(reason, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BLOCKED_NAME) || ss.insertSheet(BLOCKED_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(["Timestamp", "Reason", "Name", "Email", "Subject"]);
  sh.appendRow([new Date(), reason, safe_(String(data.name || "").slice(0, 120)), safe_(String(data.email || "").slice(0, 200)), safe_(String(data.subject || "").slice(0, 200))]);
}

// Open the /exec URL in a browser to see which version is live.
function doGet() {
  return json_({ status: "ok", service: "TessereX contact form" });
}

function doPost(e) {
  try {
    var data = e.parameter;

    // 1. Honeypot: real visitors never see this field. If it is filled, it is a bot.
    //    Reply "success" so the bot learns nothing, but save and send nothing.
    if (data.website) { blocked_("Honeypot filled", data); return json_({ result: "success" }); }

    // 2. Security check (simple sum question) and minimum fill time.
    var a = parseInt(data.cap_a, 10), b = parseInt(data.cap_b, 10), ans = parseInt(data.cap_answer, 10);
    if (isNaN(a) || isNaN(b) || a + b !== ans) { blocked_("Security check failed", data); return json_({ result: "success" }); }
    if (Number(data.elapsed) < MIN_FILL_MS) { blocked_("Sent too fast", data); return json_({ result: "success" }); }

    // 3. Read fields, apply limits, keep the enquiry type to the predefined list.
    var type = TYPES.indexOf(String(data.type)) >= 0 ? String(data.type) : DEFAULT_TYPE;
    var subject = String(data.subject || "").replace(/[\r\n]+/g, " ").trim().slice(0, 200) || "(no subject)";
    var name = String(data.name || "").slice(0, 120);
    var email = String(data.email || "").slice(0, 200);
    var message = String(data.message || "").slice(0, 5000);
    if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !message) {
      return json_({ result: "error", error: "Invalid input" });
    }

    // 4. Save to the sheet; Sr. No. counts entries (row 1 is the header).
    var sh = sheet_();
    var sr = Math.max(sh.getLastRow(), 1);
    sh.appendRow([sr, new Date(), type, safe_(subject), safe_(name), safe_(email), safe_(message)]);
    sh.getRange(sh.getLastRow(), 2).setNumberFormat(DATE_FORMAT);

    // 5. Email notification; "Reply" in Gmail goes straight to the visitor.
    var emailSubject = SUBJECT_FORMAT.split("{type}").join(type).split("{subject}").join(subject);
    var body = "Enquiry type: " + type + "\n" +
               "Name: " + name + "\n" +
               "Email: " + email + "\n\n" +
               "Message:\n" + message;
    MailApp.sendEmail(RECIPIENT, emailSubject, body, { replyTo: email });

    return json_({ result: "success" });
  } catch (error) {
    return json_({ result: "error", error: error.toString() });
  }
}

/**
 * Run ONCE from the editor (choose "setupSheet" > Run). Safe to run again.
 * - makes sure row 1 has the headers and formats the sheet
 * - converts rows saved by the OLD script (Timestamp | Name | Email | Message) into the new columns
 */
function setupSheet() {
  var sh = sheet_();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  var last = sh.getLastRow();
  for (var r = 2; r <= last; r++) {
    var row = sh.getRange(r, 1, 1, HEADERS.length).getValues()[0];
    if (row[0] instanceof Date) {   // old layout: A=Timestamp, B=Name, C=Email, D=Message
      sh.getRange(r, 1, 1, HEADERS.length).setValues([[r - 1, row[0], DEFAULT_TYPE, "(not recorded)", row[1], row[2], row[3]]]);
    }
  }
  format_(sh);
}
