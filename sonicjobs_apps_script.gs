// SonicJobs Beta Interview — Google Apps Script Web App
// Extensions > Apps Script > paste this > Deploy as Web App
// Execute as: Me | Who has access: Anyone

const SHEET_ID    = '1JaKGezgFta4pM75f8zQm_PImuNQ9jZkqgmD6U2BK1KQ';
const SHEET_NAME  = 'Responses';
const TOTAL_SPOTS = 500;
const SECRET      = '8c0b19301a1cd999690fef8a1e8cb111'; // shared with index.html
const RETAIN_DAYS = 180; // 6 months

const HEADERS = [
  'Timestamp','Name','Email','Phone','Signed Up',
  'Q1: Typical job search day',
  'Q2: Most frustrating part',
  'Q3: Jobs per week (chips)','Q3: Jobs per week (text)',
  'Q_inc1: Platforms (chips)','Q_inc1: Platforms (text)',
  'Q_inc2: What keeps you coming back',
  'Q_inc3: Irritations (chips)','Q_inc3: Irritations (text)',
  'Q_inc4: Why stayed (chips)','Q_inc4: Why stayed (text)',
  'Q_inc5: What to switch (chips)','Q_inc5: What to switch (text)',
  'Q4: Abandoned app (chips)','Q4: Abandoned app (text)',
  'Q5: 2FA reaction (chips)','Q5: 2FA reaction (text)',
  'Q6: AI tailoring (chips)','Q6: AI tailoring (text)',
  'Q7: Human vs AI control',
  'Q7b: Auto-apply worries (chips)','Q7b: Auto-apply worries (text)',
  'Q7c: AI detection (chips)','Q7c: AI detection (text)',
  'Q8: Email mgmt trust (chips)','Q8: Email mgmt trust (text)',
  'Q8b: Relay reaction (chips)','Q8b: Relay reaction (text)',
  'Q9: Profile storage (chips)','Q9: Profile storage (text)',
  'Q9b: Switching cost (chips)','Q9b: Switching cost (text)',
  'Q10: WTP (chips)','Q10: WTP (text)',
  'Q11: Proof needed',
  'Q12: Ideal redesign',
  'Brand: Applify','Brand: Propel','Brand: Hirebound','Brand: Applai','Brand: Stride',
  'Brand: Most memorable','Brand: Comments',
  'Auto-flags',
  'Delete after', // ISO date 6 months from submission
];

// ── GET: return live count ────────────────────────────────────────────────────
function doGet(e) {
  try {
    purgeExpiredRows(); // clean up on each read too
    var sheet = getOrCreateSheet();
    var count = Math.max(0, sheet.getLastRow() - 1);
    return json({ count: count, spots_left: Math.max(0, TOTAL_SPOTS - count) });
  } catch(err) {
    return json({ count: 0, spots_left: TOTAL_SPOTS, error: err.toString() });
  }
}

// ── POST: write submission ────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Token check
    if (data.token !== SECRET) {
      return json({ success: false, error: 'Unauthorized' });
    }

    var sheet    = getOrCreateSheet();
    var answers  = data.answers      || {};
    var ratings  = data.brandRatings || {};
    var flags    = data.flags        || [];
    var signup   = answers._signup   || {};

    var now        = new Date();
    var deleteDate = new Date(now.getTime() + RETAIN_DAYS * 24 * 60 * 60 * 1000);

    function chips(k) { return (answers[k + '_chips'] || []).join(', '); }
    function text(k)  { var v = answers[k] || ''; return v === '[skipped]' ? '' : v; }

    var row = [
      signup.timestamp || now.toISOString(),
      signup.name  || '',
      signup.email || '',
      signup.phone || '',
      signup.name  ? 'YES' : 'No — skipped signup',
      text('q1'), text('q2'),
      chips('q3'), text('q3'),
      chips('q_inc1'), text('q_inc1'),
      text('q_inc2'),
      chips('q_inc3'), text('q_inc3'),
      chips('q_inc4'), text('q_inc4'),
      chips('q_inc5'), text('q_inc5'),
      chips('q4'), text('q4'),
      chips('q5'), text('q5'),
      chips('q6'), text('q6'),
      text('q7'),
      chips('q7b'), text('q7b'),
      chips('q7c'), text('q7c'),
      chips('q8'), text('q8'),
      chips('q8b'), text('q8b'),
      chips('q9'), text('q9'),
      chips('q9b'), text('q9b'),
      chips('q10'), text('q10'),
      text('q11'), text('q12'),
      ratings['applify']   || '',
      ratings['propel']    || '',
      ratings['hirebound'] || '',
      ratings['applai']    || '',
      ratings['stride']    || '',
      answers['q_brand_fav'] || '',
      text('q_brand'),
      flags.map(function(f){ return f.label; }).join(', '),
      deleteDate.toISOString().split('T')[0], // YYYY-MM-DD
    ];

    sheet.appendRow(row);
    styleNewRow(sheet, sheet.getLastRow(), !!signup.name, flags);

    purgeExpiredRows();

    return json({ success: true });

  } catch(err) {
    return json({ success: false, error: err.toString() });
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function getOrCreateSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    writeHeaders(sheet);
  } else if (sheet.getLastRow() === 0) {
    writeHeaders(sheet);
  }
  return sheet;
}

function writeHeaders(sheet) {
  sheet.appendRow(HEADERS);
  var r = sheet.getRange(1, 1, 1, HEADERS.length);
  r.setBackground('#201150');
  r.setFontColor('#ffffff');
  r.setFontWeight('bold');
  r.setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 210);
  sheet.setColumnWidth(4, 130);
  for (var i = 5; i <= HEADERS.length; i++) sheet.setColumnWidth(i, 230);
}

function styleNewRow(sheet, rowNum, signedUp, flags) {
  var flagText = flags.map(function(f){ return f.label; }).join(' ');

  // Signed up cell (col 5)
  sheet.getRange(rowNum, 5)
    .setBackground(signedUp ? '#d4f1e4' : '#fff3cd')
    .setFontColor(signedUp ? '#155724' : '#856404')
    .setFontWeight('bold');

  // Flags cell (second to last col)
  var flagCol = HEADERS.length - 1;
  var fc = sheet.getRange(rowNum, flagCol);
  if (/WTP: high|Beta signed-up/.test(flagText))       fc.setBackground('#d4f1e4');
  else if (/Pay-hesitant|Incumbent-loyal/.test(flagText)) fc.setBackground('#fff3cd');

  // Delete-after cell — highlight in light red so it's visible
  sheet.getRange(rowNum, HEADERS.length)
    .setBackground('#fde8e8')
    .setFontColor('#922b21')
    .setFontSize(9);

  // Zebra stripe
  if (rowNum % 2 === 0) {
    sheet.getRange(rowNum, 1, 1, HEADERS.length).setBackground('#f8f7f5');
  }
}

function purgeExpiredRows() {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return;

    // Find "Delete after" column index
    var headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var deleteCol = headers.indexOf('Delete after') + 1;
    if (deleteCol === 0) return;

    var today     = new Date();
    var lastRow   = sheet.getLastRow();
    var rowsToDelete = [];

    for (var r = 2; r <= lastRow; r++) {
      var cellVal = sheet.getRange(r, deleteCol).getValue();
      if (cellVal) {
        var deleteDate = new Date(cellVal);
        if (deleteDate < today) rowsToDelete.push(r);
      }
    }

    // Delete bottom-up so row indices stay valid
    for (var i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
  } catch(e) {
    // Fail silently
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
