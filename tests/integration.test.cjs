const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context, Workbook, objectRows } = require('./helpers.cjs');

const empty = () => ({ hasNext: () => false, next: () => null });

function env() {
  const books = {};
  const cache = {};
  const mocks = {
    console: { log() {}, error(m) { if (process.env.DEBUG_IT) process.stderr.write(String(m) + '\n'); }, warn() {} },
    WebAccess: { requireTeacher: () => 'dan@school.edu.au' },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    CacheService: { getUserCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) },
    MimeType: { GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet' },
    DriveApp: {
      getRootFolder: () => ({ getFoldersByName: empty }),
      getFileById: id => ({ isTrashed: () => false, getMimeType: () => 'application/vnd.google-apps.spreadsheet', getName: () => books[id].getName(), makeCopy: () => ({}) })
    },
    SpreadsheetApp: { openById: id => books[id], flush() {} },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'dan@school.edu.au' }) },
    Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2), parseDate: () => new Date('2026-11-20T12:59:00Z') },
    Classroom: { Courses: { get: () => ({ name: '9DAT1', section: '' }), CourseWork: { get: () => ({ title: 'Lamp project', maxPoints: 40, dueDate: { year: 2026, month: 11, day: 20 } }) } } },
    deriveStageFromCourseCode: c => (/^9/.test(c) ? 'Stage 5' : null),
    Config: { SHEET_SUBMISSIONS: 'Submissions', SHEET_CLASSROOM_CONFIG: 'ClassroomConfig' },
    apiSubmitGradingWorkbookSetup: (code, category, task, year) => {
      const id = 'BOOK' + Object.keys(books).length + 'x'.repeat(24);
      const wb = new Workbook(code + ' [' + year + '] - ' + task + ' - Grading');
      wb.id = id; wb.insertSheet(task); books[id] = wb;
      return { success: true, message: 'Created grading workbook.', workbook: { id, name: wb.getName() }, rubricFileId: '' };
    }
  };
  const g = context(mocks).load('AppConfig.gs', 'RubricEngine.gs', 'AssessmentFlags.gs', 'EvaluationRegister.gs', 'SyncService.gs', 'WebAppIntegration.gs', 'WebGrading.gs');
  return { g, books };
}

test('one workbook per class task, created and linked from a Classroom assignment (SC04)', () => {
  const { g, books } = env();
  const a = g.apiWebCreateFromClassroom('111', '222', '9DAT1', 'Lamp', '2026', 'Projects', '');
  const b = g.apiWebCreateFromClassroom('111', '333', '9DAT1', 'Clock', '2026', 'Projects', '');
  assert.equal(a.success, true, a.message);
  assert.notEqual(a.workbook.id, b.workbook.id);
  const link = objectRows(books[a.workbook.id].getSheetByName('ClassroomConfig'))[0];
  assert.equal(link.TaskName, 'Lamp'); assert.equal(link.CourseWorkID, '222'); assert.equal(link.DueDate, '2026-11-20T12:59:00.000Z');
  assert.equal(objectRows(books[b.workbook.id].getSheetByName('ClassroomConfig'))[0].CourseWorkID, '333');
  assert.ok(books[a.workbook.id].getSheetByName('CriterionAssessments'), 'grading tables prepared');
  assert.match(a.message, /sync automatically every hour after the due date/);
  assert.equal(g.apiWebCreateFromClassroom('abc', '222', '9DAT1', 'Lamp', '2026', 'Projects').success, false);
  assert.equal(g.apiWebCreateFromClassroom('111', '222', '9DAT1', 'Lamp', '2026', 'bad/name').success, false);
});
