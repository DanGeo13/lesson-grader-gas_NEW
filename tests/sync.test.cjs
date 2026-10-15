const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context } = require('./helpers.cjs');

function load(mocks = {}) {
  return context({
    Utilities: { parseDate: (s, tz, fmt) => { assert.equal(tz, 'Australia/Sydney'); const [d, t] = s.split(' '); return new Date(d + 'T' + t + ':00+11:00'); } },
    ...mocks
  }).load('AppConfig.gs', 'SyncService.gs');
}

test('due date without a time means 23:59 Sydney, not 23:59 UTC (F46)', () => {
  const g = load();
  assert.equal(g.SyncService.classroomDueIso({ year: 2026, month: 11, day: 20 }), '2026-11-20T12:59:00.000Z');
  assert.equal(g.SyncService.classroomDueIso({ year: 2026, month: 11, day: 20 }, { hours: 3, minutes: 30 }), '2026-11-20T03:30:00.000Z');
  assert.equal(g.SyncService.classroomDueIso(null), '');
});

test('automatic sync takes only active, task-linked, due-dated assignments whose due date has passed (SC08, SC09)', () => {
  const g = load();
  const now = Date.parse('2026-10-15T00:00:00Z');
  const rows = [
    { Active: true, TaskName: 'Past', CourseID: '1', CourseWorkID: '11', DueDate: '2026-10-10T12:59:00Z' },
    { Active: true, TaskName: 'Future', CourseID: '1', CourseWorkID: '12', DueDate: '2026-11-10T12:59:00Z' },
    { Active: true, TaskName: 'Undated', CourseID: '1', CourseWorkID: '13', DueDate: '' },
    { Active: false, TaskName: 'Inactive', CourseID: '1', CourseWorkID: '14', DueDate: '2026-10-01T00:00:00Z' },
    { Active: 'TRUE', TaskName: 'Synced before', CourseID: '1', CourseWorkID: '15', DueDate: '2026-10-01T00:00:00Z', AutoImported: true }
  ];
  const due = g.SyncService.dueLinks(rows, now).map(r => r.TaskName);
  assert.equal(due.join(), 'Past,Synced before', 'already auto-imported links keep syncing hourly');
});

test('hourly run imports due links in every library workbook, records results and isolates failures', () => {
  const imports = [];
  const patches = [];
  const books = { A: { getSheetByName: () => ({}) }, B: { getSheetByName: () => ({}) } };
  const configs = {
    A: [{ _row: 2, Active: true, TaskName: 'Bridge', CourseID: '1', CourseWorkID: '11', DueDate: '2026-01-01T00:00:00Z', AutoImported: false }],
    B: [{ _row: 2, Active: true, TaskName: 'Lamp', CourseID: '2', CourseWorkID: '21', DueDate: '2026-01-01T00:00:00Z', AutoImported: true }]
  };
  const g = load({
    WebAccess: { requireTriggerOwner: () => 'dan@school.edu.au' },
    LessonGraderWeb: { list: () => ({ success: true, data: [{ id: 'A', name: 'Bridge book' }, { id: 'B', name: 'Lamp book' }] }), open: id => Object.assign(books[id], { id }) },
    WebGrading: {
      ready: () => true,
      records: (ss, name) => (name === 'ClassroomConfig' ? configs[ss.id] : []),
      ensureSheet: () => {},
      patch: (ss, sheet, row, fields) => patches.push({ book: ss.id, fields }),
      importTask: (id, task, opts) => {
        imports.push({ id, task, actor: opts.actor });
        if (id === 'B') throw new Error('Classroom quota');
        return { success: true, message: 'Synced 3', stats: { newCount: 2, updated: 1 } };
      }
    }
  });
  const res = g.SyncService.runHourly(Date.parse('2026-10-15T00:00:00Z'));
  assert.equal(res.links, 2);
  assert.equal(res.imported, 3);
  assert.equal(res.errors.length, 1);
  assert.match(res.errors[0], /Lamp book \/ Lamp: Classroom quota/);
  assert.equal(imports[0].actor, 'dan@school.edu.au');
  assert.equal(patches.find(p => p.book === 'A').fields.AutoImported, true);
  assert.match(patches.find(p => p.book === 'B').fields.LastSyncResult, /^Failed/);
});

test('trigger install is idempotent and removal leaves manual sync available', () => {
  let triggers = [];
  const g = load({
    WebAccess: { requireTeacher: () => 'dan@school.edu.au' },
    ScriptApp: {
      getProjectTriggers: () => triggers,
      deleteTrigger: t => { triggers = triggers.filter(x => x !== t); },
      newTrigger: h => ({ timeBased: () => ({ everyHours: () => ({ create: () => triggers.push({ getHandlerFunction: () => h }) }), everyMinutes: () => ({ create: () => triggers.push({ getHandlerFunction: () => h }) }) }) })
    }
  });
  g.SyncService.installTriggers();
  g.SyncService.installTriggers();
  assert.equal(triggers.length, 2);
  assert.equal(g.SyncService.triggerStatus().data.hourlySync, true);
  assert.match(g.SyncService.removeTriggers().message, /Manual "Sync now" still works/);
  assert.equal(triggers.length, 0);
});
