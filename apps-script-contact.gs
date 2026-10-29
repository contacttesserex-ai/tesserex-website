/**
 * TessereX "Get In Touch" backend (Google Apps Script)  -  VERSION v3
 *
 * HOW TO UPDATE: in your "Website Contact Submissions" Google Sheet open Extensions > Apps Script,
 * replace ALL the code with this file, press Save, then Deploy > Manage deployments > (pencil icon)
 * > Version: "New version" > Deploy.  (If you skip "New version" the OLD code keeps running.)
 * Execute as: Me.   Who has access: Anyone.
 *
 * HOW TO CHECK THE NEW CODE IS LIVE: open your web app URL (the one ending in /exec) in a browser.
 * You should see {"status":"ok","version":"v3"...}.  If you see "Script function not found: doGet",
 * the old code is still deployed.
 *
 * Email subject : TessereX Website - <Enquiry type> - <Subject typed by the visitor>
 * Email body    : Enquiry type, Name, Email, Message (the subject is not repeated in the body)
 * Sheet         : tab "Contact Submissions" with the columns
 *                 Sr No | Date & Time | Enquiry Type | Subject | Name | Email | Message
 *                 (File > Download > Microsoft Excel (.xlsx) gives an Excel copy)
 * Spam          : submissions stopped by the spam checks are listed in the tab "Blocked" with the reason.
 */
var VERSION = "v3";
var RECIPIENT = "contact.tesserex@gmail.com";            // where notification emails go
var SUBJECT_FORMAT = "TessereX Website - {type} - {subject}";
var TYPES = ["Partner Enquiries", "Investor Relations", "Web Message"];   // keep in step with the website's enquiry types
var DEFAULT_TYPE = "Web Message";
var SHEET_NAME = "Contact Submissions";
var BLOCKED_NAME = "Blocked";
var MIN_FILL_MS = 1500;                                  // forms sent faster than this are treated as bots
var HEADERS = ["Sr No", "Date & Time", "Enquiry Type", "Subject", "Name", "Email", "Message"];

function json_(obj) {
  obj.version = VERSION;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
// Stops text starting with = + - @ from being treated as a spreadsheet formula.
function safe_(v) {
  v = String(v == null ? "" : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
// Gets (or creates) the submissions tab with a tidy header row.
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#e8eefc");
    sh.getRange("B:B").setNumberFormat("dd-mmm-yyyy hh:mm:ss");
    sh.getRange("G:G").setWrap(true);
    sh.setColumnWidth(2, 150); sh.setColumnWidth(4, 260); sh.setColumnWidth(6, 200); sh.setColumnWidth(7, 420);
  }
  return sh;
}
// Records a submission that the spam checks stopped, so nothing disappears silently.
function blocked_(reason, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BLOCKED_NAME) || ss.insertSheet(BLOCKED_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(["Date & Time", "Reason", "Name", "Email", "Subject"]);
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

    // 4. Save to the sheet; Sr No counts entries (row 1 is the header).
    var sh = sheet_();
    var sr = Math.max(sh.getLastRow(), 1);
    sh.appendRow([sr, new Date(), type, safe_(subject), safe_(name), safe_(email), safe_(message)]);

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
