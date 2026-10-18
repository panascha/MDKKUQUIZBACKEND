// Offline harness for question-id → Drive folder routing (images.gs upload + maintenance.gs migrateSubjectImages).
// Run: node tools/year-routing-harness.cjs (from repo root)
const fs = require('fs'), vm = require('vm'), path = require('path'), assert = require('assert');
const root = path.join(__dirname, '..');
const files = ['images.gs', 'maintenance.gs'];
const srcs = files.map(f => fs.readFileSync(path.join(root, f), 'utf8'));
files.forEach((f, i) => new vm.Script(srcs[i], { filename: f })); // syntax check

// Fake Drive: folder id = its path, so final path is readable straight off the folder
function makeDrive() {
  const byId = {};
  function folder(p) {
    if (byId[p]) return byId[p];
    const kids = {};
    const f = {
      path: p,
      getId: () => p,
      isTrashed: () => false,
      getFoldersByName: n => { let done = !kids[n]; return { hasNext: () => !done, next: () => { done = true; return kids[n]; } }; },
      createFolder: n => (kids[n] = folder(p ? p + '/' + n : n)),
      createFile: () => { drive.lastFileFolder = p; return { getId: () => 'FILE' }; },
    };
    return (byId[p] = f);
  }
  const rootFolder = folder('');
  const drive = { lastFileFolder: null, getFolderById: id => byId[id] || rootFolder };
  return drive;
}

function makeCtx(qRows, structRows) {
  const ctx = {
    console: { log() {}, warn() {}, error() {} },
    DRIVE_FOLDER_ID: '',
    DriveApp: makeDrive(),
    CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) },
    Utilities: { newBlob: () => ({}), base64Decode: () => [], sleep() {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty() {}, deleteProperty() {} }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'harness@kku.ac.th' }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => ({ getDataRange: () => ({ getValues: () => qRows }) }) }) },
    getStructureSheetDataCached: () => structRows,
  };
  vm.createContext(ctx);
  srcs.forEach((s, i) => vm.runInContext(s, ctx, { filename: files[i] }));
  return ctx;
}

// images.gs: upload with folder captured from createFile
function uploadPath(qid, subjectHint, yearHint, structRows) {
  const ctx = makeCtx([], structRows || [['Year', 'SubjectID']]);
  ctx.uploadQuestionImageToDrive('data:image/png;base64,AAAA', qid, 'Main', subjectHint, yearHint);
  return ctx.DriveApp.lastFileFolder;
}

// maintenance.gs: migrateSubjectImages with moveAndRenameFile captured
function migratePath(qid, subjectPrefix, yearFolder) {
  const ctx = makeCtx([['ID', 'q', 'img', 'choices'], [qid, '', 'https://drive.google.com/uc?id=X', '']], []);
  let dest = null;
  ctx.moveAndRenameFile = (url, name, target) => { dest = target.path; };
  ctx.migrateSubjectImages(subjectPrefix, yearFolder);
  return dest;
}

let pass = 0, fail = 0;
function check(name, got, want) {
  try { assert.strictEqual(got, want); pass++; console.log('PASS', name); }
  catch (e) { fail++; console.log('FAIL', name, '→ got', got, 'want', want); }
}

// T0.1 — the 4 plan cases, both code paths
const cases = [
  ['NL2_2025R1_001', 'NL2', 'NL', 'MD/YNL/NL2/2025/R1'],
  ['CVS_51MCQ1_01', 'CVS', '2', 'MD/Y2/CVS/51/MCQ1'],
  ['MED4_49MCQ_ROTA_001', 'MED4', '4', 'MD/Y4/MED4/49/MCQ'],
  ['MED4_xxMCQ_ROTX_001', 'MED4', '4', 'MD/Y4/MED4/xxMCQ'],
];
for (const [qid, subj, year, want] of cases) {
  check('upload  ' + qid, uploadPath(qid, subj, year), want);
  check('migrate ' + qid, migratePath(qid, subj, 'Y' + year), want);
}

// T0.2 — no hints: year resolved from Structure row ['NL','NL2']
check('upload  NL2_2025R1_001 via Structure year NL',
  uploadPath('NL2_2025R1_001', undefined, undefined, [['Year', 'SubjectID'], ['2', 'CVS'], ['NL', 'NL2']]),
  'MD/YNL/NL2/2025/R1');

console.log(pass + ' passed, ' + fail + ' failed');
process.exitCode = fail ? 1 : 0;
