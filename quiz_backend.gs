// ============================================================
// 스마트 문제집 - Google Apps Script Backend
// v1.0
// Sheets: "Sessions", "Questions"
// Deploy: 배포 → 웹 앱 → 실행 계정(나), 액세스(모든 사용자)
// ============================================================

const SHEET_SESSIONS  = 'Sessions';
const SHEET_QUESTIONS = 'Questions';

// ── CORS helper ──────────────────────────────────────────────
function addCors(output) {
  return output; // Apps Script auto-adds CORS for web app deployments
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Router ───────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheets(ss);
    let result;
    switch (body.action) {
      case 'saveSession':  result = saveSession(ss, body);  break;
      case 'saveQuestion': result = saveQuestion(ss, body); break;
      default: throw new Error('Unknown action: ' + body.action);
    }
    return jsonOut({ success: true, data: result });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheets(ss);
    let result;
    switch (action) {
      case 'getAnalysisData': result = getAnalysisData(ss); break;
      case 'getHistory':      result = getHistory(ss);      break;
      default: throw new Error('Unknown action: ' + action);
    }
    return jsonOut({ success: true, data: result });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

// ── Sheet setup ──────────────────────────────────────────────
function ensureSheets(ss) {
  // Sessions
  if (!ss.getSheetByName(SHEET_SESSIONS)) {
    const s = ss.insertSheet(SHEET_SESSIONS);
    s.getRange(1, 1, 1, 7).setValues([[
      'session_id','date','source_desc',
      'total_questions','correct','wrong_attempts_total','subject_type'
    ]]);
    s.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    s.setFrozenRows(1);
  }

  // Questions
  if (!ss.getSheetByName(SHEET_QUESTIONS)) {
    const s = ss.insertSheet(SHEET_QUESTIONS);
    s.getRange(1, 1, 1, 9).setValues([[
      'session_id','q_index','question','correct_answer',
      'student_answer','is_correct','wrong_count','topic','timestamp'
    ]]);
    s.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    s.setFrozenRows(1);
  }
}

// ── Writers ──────────────────────────────────────────────────
function saveSession(ss, d) {
  ss.getSheetByName(SHEET_SESSIONS).appendRow([
    d.session_id,
    d.date,
    d.source_desc    || '',
    d.total_questions || 0,
    d.correct        || 0,
    d.wrong_attempts_total || 0,
    d.subject_type   || ''
  ]);
  return { saved: true };
}

function saveQuestion(ss, d) {
  ss.getSheetByName(SHEET_QUESTIONS).appendRow([
    d.session_id,
    d.q_index,
    d.question,
    d.correct_answer,
    d.student_answer || '',
    d.is_correct ? 'TRUE' : 'FALSE',
    d.wrong_count    || 0,
    d.topic          || '',
    new Date().toISOString()
  ]);
  return { saved: true };
}

// ── Readers ──────────────────────────────────────────────────
function getAnalysisData(ss) {
  const sheet = ss.getSheetByName(SHEET_QUESTIONS);
  const raw   = sheet.getDataRange().getValues();
  if (raw.length <= 1) return { questions: [] };

  const headers = raw[0];
  const rows = raw.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  return { questions: rows };
}

function getHistory(ss) {
  const sessions  = ss.getSheetByName(SHEET_SESSIONS).getDataRange().getValues();
  const questions = ss.getSheetByName(SHEET_QUESTIONS).getDataRange().getValues();
  return { sessions, questions };
}
