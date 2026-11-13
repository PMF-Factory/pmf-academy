/**
 * Swiss PMF Academy – PMF Check results → Google Sheets  (v2: no duplicates)
 *
 * Each browser gets a random, anonymous "Visitor ID" (stored in the visitor's
 * browser, no personal data). If the same visitor completes the check again or
 * changes answers, their existing row is UPDATED instead of adding a new one.
 *
 * Update an existing deployment:
 * 1. Extensions → Apps Script: replace all code with this file. Save.
 * 2. Deploy → Manage deployments → pencil (Edit) → Version: "New version" → Deploy.
 *    (Execute as: Me · Who has access: Anyone). The URL stays the same.
 *
 * Answers: 3 = Very true, 2 = Somewhat true, 1 = Not yet true, 0 = Not true at all.
 */

var HEADERS = [
  'Submitted',
  'Score (/100)',
  'Points (/24)',
  'Result',
  'Q1 Growth repeatable & predictable',
  'Q2 Customers act fast / urgency',
  'Q3 CAC dropped significantly',
  'Q4 Sales process repeatable & scalable',
  'Q5 Significant problem without product',
  'Q6 Customers buy more or again',
  'Q7 Customers recommend you',
  'Q8 Organic growth major driver',
  'Page',
  'Visitor ID',
  'Updates'
];
var ID_COL = 14;      // column N: Visitor ID
var UPDATES_COL = 15; // column O: how often this visitor re-submitted

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    // Write / repair the header row (also adds the new columns to an existing sheet)
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    var d = JSON.parse(e.postData.contents);
    var a = Array.isArray(d.answers) ? d.answers.slice(0, 8) : [];
    if (a.length !== 8) throw new Error('Expected 8 answers');
    a = a.map(function (v) { v = Number(v); return (v >= 0 && v <= 3) ? v : ''; });
    var id = String(d.id || '').replace(/[^A-Za-z0-9-]/g, '').slice(0, 64);

    var row = [
      new Date(),
      Number(d.score) || 0,
      Number(d.sum) || 0,
      String(d.result || '').slice(0, 100)
    ].concat(a, [String(d.page || '').slice(0, 300), id]);

    // Same visitor already in the sheet? → update that row
    var target = 0, updates = 0;
    var last = sheet.getLastRow();
    if (id && last > 1) {
      var ids = sheet.getRange(2, ID_COL, last - 1, 2).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === id) { target = i + 2; updates = (Number(ids[i][1]) || 0) + 1; break; }
      }
    }
    row.push(updates);
    if (target) {
      sheet.getRange(target, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, updated: !!target }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Open the web app URL in a browser to check it is live.
function doGet() {
  return ContentService.createTextOutput('PMF Check endpoint is live (v2).');
}
