const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context, Workbook } = require('./helpers.cjs');

test('health check reports configuration, AI, triggers and grading-rule gaps without writing or leaking secrets', () => {
  const props = { LESSON_GRADER_ALLOWED_EMAILS: 'dan@school.edu.au', GEMINI_API_KEY: 'SECRET_GEMINI_KEY_123456' };
  let writes = 0;
  const ss = new Workbook();
  ss.add('Bridge', [['Task']]);
  ss.add('Setup', [['GRADE SCALE', 'WEIGHT'], ['A', 1], ['B', 0.875], ['C', 0.7], ['D', 0.575], ['E', 0.25]]);
  const g = context({
    Session: { getActiveUser: () => ({ getEmail: () => 'dan@school.edu.au' }), getScriptTimeZone: () => 'Australia/Sydney' },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: () => { writes++; } }) },
    ScriptApp: { getProjectTriggers: () => [] },
    ClassroomService: { listTeacherCourses: () => ({ success: true, courses: [{}, {}] }) },
    LessonGraderWeb: { open: () => ss, taskNames: () => ['Bridge'] },
    WebGrading: { ready: () => false, records: () => [] }
  }).load('AppConfig.gs', 'WebAccess.gs', 'RubricEngine.gs', 'AssessmentFlags.gs', 'PromptRegistry.gs', 'AiGateway.gs', 'SyncService.gs', 'HealthCheck.gs');
  const res = g.apiWebHealthCheck('BOOK');
  const json = JSON.stringify(res);
  assert.ok(!json.includes('SECRET_GEMINI_KEY'));
  assert.equal(writes, 0);
  const find = name => res.data.checks.find(c => c.check === name);
  assert.equal(find('Teacher allowlist').status, 'pass');
  assert.equal(find('Authorised unlock list').status, 'warn');
  assert.equal(find('Hourly Classroom sync trigger').status, 'warn');
  assert.equal(find('Grading tables').status, 'fail');
  assert.equal(find('Reference points A/C/E (SC19)').status, 'warn');
  assert.match(find('Reference points A/C/E (SC19)').detail, /C 70% vs 50%/);
  assert.equal(find('Active rubric').status, 'warn');
  assert.equal(find('Classroom access').status, 'pass');
});
