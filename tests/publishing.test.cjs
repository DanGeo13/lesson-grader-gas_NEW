const { test } = require('node:test');
const assert = require('node:assert/strict');
const { objectRows } = require('./helpers.cjs');
const { setup, rowsOne, classroomFake, linkTask } = require('./env.cjs');

/* Two Classroom students graded and approved; a third left as a draft. */
function graded() {
  const { g, ss } = setup();
  const W = g.WebGrading;
  W.prepare(ss, false);
  W.activateRubric(ss.id, 'Rube Goldberg', 'Rube', '', rowsOne(), {}, true);
  linkTask(ss, 'Rube Goldberg', 'WORK-1');
  const att = id => ({ driveFile: { id, title: id, mimeType: 'application/pdf' } });
  classroomFake(g, { 'WORK-1': [
    { id: 'CS-A', userId: 'U-A', state: 'TURNED_IN', assignmentSubmission: { attachments: [att('A')] } },
    { id: 'CS-B', userId: 'U-B', state: 'TURNED_IN', assignmentSubmission: { attachments: [att('B')] } },
    { id: 'CS-C', userId: 'U-C', state: 'TURNED_IN', assignmentSubmission: { attachments: [att('C')] } }
  ] }, { 'U-A': { name: { fullName: 'Ava Chen' } }, 'U-B': { name: { fullName: 'Ben Ito' } }, 'U-C': { name: { fullName: 'Cal Ng' } } });
  W.importTask(ss.id, 'Rube Goldberg');
  const ids = {};
  objectRows(ss.getSheetByName('Submissions')).filter(s => s.StudentUserID.startsWith('U-')).forEach(s => { ids[s.StudentName] = s.SubmissionRecordID; });
  const approve = (name, ticks) => {
    const rev = W.detail(ss.id, 'Rube Goldberg', ids[name]).data.revision;
    const d = W.saveDraft(ss.id, 'Rube Goldberg', ids[name], ticks, { whatWentWell: 'Clear research.', markFeedbackCurrent: true }, rev, 'Teacher');
    return W.approve(ss.id, 'Rube Goldberg', ids[name], d.revision);
  };
  approve('Ava Chen', { C01: { checkboxes: { 'C01-A-1': true } }, C02: { checkboxes: { 'C02-A-1': true } } });   // 25/25
  approve('Ben Ito', { C01: { checkboxes: { 'C01-B-1': true } }, C02: { checkboxes: { 'C02-A-1': true } } });    // 23.75/25
  const rev = W.detail(ss.id, 'Rube Goldberg', ids['Cal Ng']).data.revision;
  W.saveDraft(ss.id, 'Rube Goldberg', ids['Cal Ng'], { C01: { checkboxes: { 'C01-A-1': true } }, C02: { checkboxes: {} } }, {}, rev, 'Teacher');
  const patches = [];
  g.Classroom.Courses.CourseWork.get = () => ({ title: 'Rube Goldberg machine', maxPoints: 20 });
  g.Classroom.Courses.CourseWork.StudentSubmissions.patch = (body, course, work, sub, opts) => { patches.push({ body, course, work, sub, opts }); return {}; };
  return { g, ss, ids, patches };
}

test('approval never contacts Classroom (SC39)', () => {
  const { patches, ss } = graded();
  assert.equal(patches.length, 0);
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades')).length, 2);
});

test('publishing is separate, previewed, typed-confirmed and affects only the selected student (SC40)', () => {
  const { g, ss, ids, patches } = graded();
  const P = g.Publishing;
  const preview = P.preview(ss.id, 'Rube Goldberg', [ids['Ben Ito'], ids['Cal Ng']]).data;
  assert.equal(patches.length, 0, 'preview sends nothing');
  assert.equal(preview.destination.assignmentTitle, 'Rube Goldberg machine');
  assert.equal(preview.eligible.length, 1);
  assert.equal(preview.eligible[0].name, 'Ben Ito');
  assert.equal(preview.eligible[0].grade, 19);              // 23.75 / 25 x 20 = 19.0
  assert.equal(preview.ineligible[0].reason, 'Not approved.');
  assert.match(preview.eligible[0].comment, /^Ben, here is your feedback/);
  assert.throws(() => P.publish(ss.id, 'Rube Goldberg', [ids['Ben Ito']], preview.previewToken, 'yes'), /Type PUBLISH/);
  const res = P.publish(ss.id, 'Rube Goldberg', [ids['Ben Ito'], ids['Cal Ng']], preview.previewToken, 'PUBLISH');
  assert.equal(res.success, true);
  assert.equal(patches.length, 1);
  assert.equal(patches[0].sub, 'CS-B');
  assert.equal(patches[0].body.assignedGrade, 19);
  assert.equal(patches[0].opts.updateMask, 'draftGrade,assignedGrade');
  const subs = objectRows(ss.getSheetByName('Submissions'));
  assert.match(subs.find(s => s.StudentName === 'Ben Ito').LastSyncResult, /^Published 19\/20 · draft/);
  assert.equal(subs.find(s => s.StudentName === 'Ava Chen').LastSyncResult, '', 'unselected approved student untouched');
  assert.equal(objectRows(ss.getSheetByName('PublishLog')).length, 1);
});

test('publishing refuses a stale preview and reports Classroom project limits honestly', () => {
  const { g, ss, ids } = graded();
  const P = g.Publishing;
  const preview = P.preview(ss.id, 'Rube Goldberg', [ids['Ava Chen']]).data;
  g.Classroom.Courses.CourseWork.get = () => ({ title: 'Rube Goldberg machine', maxPoints: 30 });   // assignment changed
  assert.throws(() => P.publish(ss.id, 'Rube Goldberg', [ids['Ava Chen']], preview.previewToken, 'PUBLISH'), /Preview again/);
  const fresh = P.preview(ss.id, 'Rube Goldberg', [ids['Ava Chen']]).data;
  g.Classroom.Courses.CourseWork.StudentSubmissions.patch = () => { throw new Error('@ProjectPermissionDenied The Developer Console project is not permitted to make this request.'); };
  const res = P.publish(ss.id, 'Rube Goldberg', [ids['Ava Chen']], fresh.previewToken, 'PUBLISH');
  assert.equal(res.success, false);
  assert.match(res.data.results[0].message, /only lets an app change grades on assignments that app created/);
  assert.match(objectRows(ss.getSheetByName('Submissions')).find(s => s.StudentName === 'Ava Chen').LastSyncResult, /^Publish failed/);
  assert.equal(objectRows(ss.getSheetByName('PublishLog'))[0].Status, 'Failed');
});

test('ungraded Classroom assignments cannot receive a numeric grade', () => {
  const { g, ss, ids } = graded();
  g.Classroom.Courses.CourseWork.get = () => ({ title: 'Ungraded', maxPoints: null });
  const preview = g.Publishing.preview(ss.id, 'Rube Goldberg', [ids['Ava Chen']]).data;
  assert.equal(preview.eligible.length, 0);
  assert.match(preview.ineligible[0].reason, /ungraded/);
});
