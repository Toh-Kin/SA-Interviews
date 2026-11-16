/**
 * Interview Recorder - server (Google Apps Script), v0.3.1
 *
 * What it does
 *  - Hands the question list / school list / interviewer list to the phone app
 *  - Receives the audio recording (in small pieces) and saves it to Google Drive
 *  - Writes the interview details and per-question notes into this spreadsheet
 *
 * One-time setup: see README (run setup() once, then Deploy > New deployment > Web app).
 * Updating from an older version: paste this file, set PASSCODE again, run setup() once (it upgrades the
 * sheets in place), then Deploy > Manage deployments > Edit > New version.
 */

// ===== EDIT THESE TWO LINES =====
const PASSCODE = 'Rwanda2026!';   // shared secret typed (or sent by link) to the team's phones
const FOLDER_ID = '';            // Drive folder ID for recordings. Leave '' to auto-create "DS Teacher Interviews" in your My Drive.
// ================================

const FOLDER_NAME = 'DS Teacher Interviews';
const DEFAULT_TYPES = ['Mentor Teacher', 'Student Teacher'];
const SHEETS = {Q: 'Questions', S: 'Schools', W: 'Interviewers', I: 'Interviews', N: 'Notes'};
const HEAD_Q = ['No', 'Target', 'Section', 'Stage', 'Question_EN', 'Question_RW', 'Probe_EN', 'Probe_RW', 'Required', 'Guide_EN'];
// New columns are always added at the END so that older data stays aligned.
const HEAD_I = ['interview_id', 'date', 'start_time', 'school', 'interviewee_id', 'interviewer', 'duration', 'consent_confirmed',
                'questions_with_notes', 'audio_file', 'audio_url', 'overall_impressions', 'demo', 'app_version', 'device', 'uploaded_at',
                'interviewee_type'];
const HEAD_N = ['interview_id', 'date', 'school', 'interviewee_id', 'interviewer', 'q_no', 'stage', 'section', 'question',
                'first_asked_at', 'notes', 'marks_at', 'question_rw', 'interviewee_type'];

/* ---------- web app entry points ---------- */

function doGet() {
  return ContentService.createTextOutput('Interview Recorder server is running (v0.3.1).');
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.key !== PASSCODE) return out_({ok: false, error: 'Wrong passcode.'});
    switch (b.action) {
      case 'config':        return out_({ok: true, data: getConfig_()});
      case 'audioChunk':    return out_(audioChunk_(b));
      case 'audioFinalize': return out_(audioFinalize_(b));
      case 'saveInterview': return out_(saveInterview_(b.record));
      default:              return out_({ok: false, error: 'Unknown action.'});
    }
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- one-time setup / upgrade (run from the editor) ---------- */

function setup() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the spreadsheet: Extensions > Apps Script.');
  props.setProperty('SS_ID', ss.getId());

  // Questions: create with samples, or upgrade an older layout in place (data is kept)
  const q = sheet_(ss, SHEETS.Q, HEAD_Q);
  if (q.getLastRow() < 2) {
    q.getRange(2, 1, 10, HEAD_Q.length).setValues([
      [1, 'Both', 'Opening', '', 'Please tell me briefly about your role.', '(Kinyarwanda translation here)', 'How long have you been in this role?', '', 'Y', ''],
      [2, 'Mentor Teacher', 'A. Observation', 'Observation', 'Think about the lesson you observed. What did you notice about the learners?', '', 'What did you see or hear that made you say that?', '', 'Y', 'What we want to hear: what the learners did, in their own words, with a real example.'],
      [3, 'Mentor Teacher', 'A. Observation', 'Observation', 'What did the student-teacher do well in that lesson?', '', 'Can you give one specific moment?', '', 'Y', ''],
      [4, 'Mentor Teacher', 'B. Feedback', 'Feedback', 'What feedback did you give after the lesson?', '', 'When? Face to face or written? How long did it take?', '', 'Y', ''],
      [5, 'Mentor Teacher', 'C. Next action', 'Advice', 'What advice or next action did you suggest?', '', 'What exactly should they try in the next lesson?', '', 'Y', ''],
      [6, 'Mentor Teacher', 'D. Follow-up', 'Re-observation', 'Did you observe the student-teacher again after the advice?', '', 'If yes: what changed? If no: what got in the way?', '', 'Y', ''],
      [7, 'Student Teacher', 'A. Observation', 'Observation', 'After your lesson, what did your Mentor Teacher say about it?', '', 'What did they notice about the learners?', '', 'Y', ''],
      [8, 'Student Teacher', 'B. Feedback', 'Feedback', 'What feedback did you receive? How was it given?', '', 'When? Face to face or written?', '', 'Y', ''],
      [9, 'Student Teacher', 'C. Next action', 'Advice', 'What advice or next action did you get, and what did you do with it?', '', 'Did you try it in a later lesson?', '', 'Y', ''],
      [10, 'Student Teacher', 'D. Follow-up', 'Re-observation', 'Did your Mentor Teacher observe you again? Was it useful?', '', 'What changed, if anything?', '', '', '']
    ]);
  } else {
    upgradeQuestionsSheet_(q);
  }
  q.getRange('A1').setNote('Question number. Rows are shown to interviewers in this order.');
  q.getRange('B1').setNote('Who is asked: Mentor Teacher / Student Teacher / Both (blank = Both). The interviewer picks the type at the start.');
  q.getRange('C1').setNote('Heading shown above the question.');
  q.getRange('D1').setNote('Optional small tag: Observation / Feedback / Advice / Re-observation ... (useful when comparing across surveys).');
  q.getRange('E1').setNote('Question in English.');
  q.getRange('F1').setNote('Question in Kinyarwanda. Leave blank if there is none; the app then shows English only.');
  q.getRange('G1').setNote('Optional follow-up prompt (English).');
  q.getRange('H1').setNote('Optional follow-up prompt (Kinyarwanda).');
  q.getRange('I1').setNote('Y = required, blank = optional.');

  const s = sheet_(ss, SHEETS.S, ['School']);
  if (s.getLastRow() < 2) s.getRange(2, 1, 2, 1).setValues([['Sample School A (edit me)'], ['Sample School B (edit me)']]);
  const w = sheet_(ss, SHEETS.W, ['Interviewer']);
  if (w.getLastRow() < 2) w.getRange(2, 1, 3, 1).setValues([['Kin'], ['Local member 1 (edit me)'], ['Local member 2 (edit me)']]);
  ensureHeader_(sheet_(ss, SHEETS.I, HEAD_I), HEAD_I);
  ensureHeader_(sheet_(ss, SHEETS.N, HEAD_N), HEAD_N);
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) ss.deleteSheet(first);

  const f = mainFolder_();
  Logger.log('Setup done. Recordings folder: ' + f.getUrl());
  Logger.log('Now deploy: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).');
}

/* ---------- helpers ---------- */

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#e0f2fe');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Add any header names that are missing, at the end of row 1.
function ensureHeader_(sh, head) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  let col = cur.filter(String).length;
  head.forEach(h => {
    if (cur.indexOf(h) < 0) {
      col++;
      sh.getRange(1, col, 1, 1).setValues([[h]]).setFontWeight('bold').setBackground('#e0f2fe');
    }
  });
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const a = SpreadsheetApp.getActiveSpreadsheet();
  if (!a) throw new Error('Run setup() once first.');
  return a;
}

function mainFolder_() {
  if (FOLDER_ID) return DriveApp.getFolderById(FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('AUTO_FOLDER_ID');
  if (saved) { try { return DriveApp.getFolderById(saved); } catch (e) { /* recreate below */ } }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('AUTO_FOLDER_ID', f.getId());
  return f;
}

function tmpFolder_() {
  const main = mainFolder_();
  const it = main.getFoldersByName('_tmp');
  return it.hasNext() ? it.next() : main.createFolder('_tmp');
}

function col1_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim()).filter(String);
}

function p_(n, w) { let s = String(n); while (s.length < w) s = '0' + s; return s; }
function mmss_(ms) {
  const s = Math.max(0, Math.round(ms / 1000));
  const h = Math.floor(s / 3600), m = Math.floor(s % 3600 / 60), x = s % 60;
  return (h ? h + ':' + p_(m, 2) : p_(m, 2)) + ':' + p_(x, 2);
}

/* ---------- questions: read by header NAME, so old and new layouts both work ---------- */

const Q_ALIASES = {
  no: ['no', 'number', '#'],
  target: ['target', 'respondent', 'for', 'type', 'interviewee'],
  section: ['section'],
  stage: ['stage'],
  qEn: ['questionen', 'questionenglish', 'question', 'en', 'english'],
  qRw: ['questionrw', 'questionkinyarwanda', 'questionkin', 'rw', 'kinyarwanda'],
  pEn: ['probeen', 'probeenglish', 'probe'],
  pRw: ['probekinyarwanda', 'probe_rw', 'proberw', 'probekin'],
  required: ['required', 'mandatory'],
  gEn: ['guideen', 'guide', 'guidingquestion', 'whatwewanttohear']
};

function norm_(h) { return String(h).toLowerCase().replace(/[\s_()\-]/g, ''); }

function qColumns_(headerRow) {
  const idx = headerRow.map(norm_);
  const cols = {};
  Object.keys(Q_ALIASES).forEach(k => {
    cols[k] = -1;
    for (let a = 0; a < Q_ALIASES[k].length && cols[k] < 0; a++) cols[k] = idx.indexOf(norm_(Q_ALIASES[k][a]));
  });
  return cols;
}

function readQuestionRows_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const n = Math.max(0, sh.getLastRow() - 1);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const cols = qColumns_(head);
  const rows = n ? sh.getRange(2, 1, n, lastCol).getValues() : [];
  const get = (r, k) => (cols[k] >= 0 ? r[cols[k]] : '');
  return rows.map((r, i) => ({
    no: get(r, 'no') === '' ? i + 1 : get(r, 'no'),
    target: String(get(r, 'target')).trim(),
    section: String(get(r, 'section')).trim(),
    stage: String(get(r, 'stage')).trim(),
    qEn: String(get(r, 'qEn')).trim(),
    qRw: String(get(r, 'qRw')).trim(),
    pEn: String(get(r, 'pEn')).trim(),
    pRw: String(get(r, 'pRw')).trim(),
    gEn: String(get(r, 'gEn')).trim(),
    required: /^(y|yes|true|1|required)$/i.test(String(get(r, 'required')).trim())
  })).filter(q => q.qEn !== '' || q.qRw !== '');
}

// Old layout (No, Section, Stage, Question, Probe, Required) -> new layout, keeping every row.
function upgradeQuestionsSheet_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const same = HEAD_Q.every((h, i) => cur[i] === h) ;
  if (same) return;
  const rows = readQuestionRows_(sh);
  sh.clear();
  sh.getRange(1, 1, 1, HEAD_Q.length).setValues([HEAD_Q]).setFontWeight('bold').setBackground('#e0f2fe');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, HEAD_Q.length).setValues(rows.map(q =>
      [q.no, q.target || 'Both', q.section, q.stage, q.qEn, q.qRw, q.pEn, q.pRw, q.required ? 'Y' : '', q.gEn || '']));
  }
}

function getConfig_() {
  const ss = ss_();
  const questions = readQuestionRows_(ss.getSheetByName(SHEETS.Q)).map(q => ({
    no: q.no, target: (!q.target || /^(both|all)$/i.test(q.target)) ? 'Both' : q.target,
    section: q.section, stage: q.stage, qEn: q.qEn, qRw: q.qRw, pEn: q.pEn, pRw: q.pRw, gEn: q.gEn, required: q.required
  }));
  const types = [];
  questions.forEach(q => { if (q.target !== 'Both' && types.indexOf(q.target) < 0) types.push(q.target); });
  return {
    questions: questions,
    types: types.length ? types : DEFAULT_TYPES,
    schools: col1_(ss.getSheetByName(SHEETS.S)),
    interviewers: col1_(ss.getSheetByName(SHEETS.W))
  };
}

/* ---------- audio ---------- */

function audioChunk_(b) {
  const tmp = tmpFolder_();
  const name = b.id + '__' + p_(b.index, 4);
  const old = tmp.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  const bytes = Utilities.base64Decode(b.data);
  tmp.createFile(Utilities.newBlob(bytes, 'application/octet-stream', name));
  return {ok: true};
}

function audioFinalize_(b) {
  const main = mainFolder_();
  const marker = 'interview_id=' + b.id;
  const baseName = String(b.filename || (b.id + '.webm')).replace(/[\\/:*?"<>|]+/g, '_');

  // Already finalized earlier (a retry after a lost reply)? Return that file.
  const same = main.getFilesByName(baseName);
  while (same.hasNext()) {
    const f = same.next();
    if (f.getDescription() === marker && !f.isTrashed()) {
      trashPieces_(b.id);   // clean up any pieces re-sent by a retry
      return {ok: true, url: f.getUrl(), name: f.getName()};
    }
  }

  const tmp = tmpFolder_();
  let all = [];
  const pieces = [];
  for (let i = 0; i < b.total; i++) {
    const it = tmp.getFilesByName(b.id + '__' + p_(i, 4));
    if (!it.hasNext()) return {ok: false, error: 'Audio piece ' + (i + 1) + ' of ' + b.total + ' is missing on the server. Please send again.'};
    const f = it.next();
    pieces.push(f);
    all = all.concat(f.getBlob().getBytes());
  }
  // Pick a free file name (another interview may already use it)
  let name = baseName, k = 2;
  const dot = baseName.lastIndexOf('.');
  const stem = dot > 0 ? baseName.slice(0, dot) : baseName, ext = dot > 0 ? baseName.slice(dot) : '';
  while (main.getFilesByName(name).hasNext()) { name = stem + '_' + k + ext; k++; }
  const file = main.createFile(Utilities.newBlob(all, b.mime || 'audio/webm', name));
  file.setDescription(marker);
  pieces.forEach(f => f.setTrashed(true));
  return {ok: true, url: file.getUrl(), name: file.getName()};
}

function trashPieces_(id) {
  const tmp = tmpFolder_();
  const it = tmp.getFiles();
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf(id + '__') === 0) f.setTrashed(true); }
}

/* ---------- interview details + notes ---------- */

function saveInterview_(r) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    const sI = ss.getSheetByName(SHEETS.I), sN = ss.getSheetByName(SHEETS.N);
    ensureHeader_(sI, HEAD_I);
    ensureHeader_(sN, HEAD_N);
    const qs = r.questions || [];
    const withNotes = qs.filter(q => String((r.notes || {})[String(q.no)] || '').trim()).length;
    const row = [r.id, r.date, r.startTime, r.school, r.interviewee, r.interviewer, mmss_((r.durationSec || 0) * 1000),
                 r.consent ? 'Yes' : 'No', withNotes + '/' + qs.length, r.audioName || '', r.audioUrl || '',
                 r.general || '', r.demo ? 'DEMO' : '', r.appVersion || '', r.device || '', new Date().toISOString(),
                 r.type || ''];
    upsertRow_(sI, r.id, row);

    // replace this interview's note rows
    const last = sN.getLastRow();
    if (last > 1) {
      const ids = sN.getRange(2, 1, last - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(r.id)) sN.deleteRow(i + 2);
    }
    const rows = qs.map(q => {
      const v = (r.visits || []).filter(x => String(x.q) === String(q.no))[0];
      const marks = (r.flags || []).filter(x => String(x.q) === String(q.no)).map(x => mmss_(x.t)).join(', ');
      return [r.id, r.date, r.school, r.interviewee, r.interviewer, q.no, q.stage || '', q.section || '', q.qEn || '',
              v ? mmss_(v.t) : '', String((r.notes || {})[String(q.no)] || ''), marks, q.qRw || '', r.type || ''];
    });
    if (rows.length) {
      const start = sN.getLastRow() + 1;
      const rg = sN.getRange(start, 1, rows.length, HEAD_N.length);
      rg.setNumberFormat('@');          // store as plain text so notes starting with "=" are not treated as formulas
      rg.setValues(rows);
    }
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function upsertRow_(sh, id, row) {
  const last = sh.getLastRow();
  let target = last + 1;
  if (last > 1) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) { target = i + 2; break; }
  }
  const rg = sh.getRange(target, 1, 1, row.length);
  rg.setNumberFormat('@');
  rg.setValues([row]);
}
