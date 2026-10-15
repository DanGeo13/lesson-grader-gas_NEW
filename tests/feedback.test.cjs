const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context } = require('./helpers.cjs');
const { rows, LEGACY_WEIGHTS, BANDS } = require('./fixtures/reference-rubric.cjs');
const g = context().load('RubricEngine.gs', 'FeedbackBuilder.gs');
const R = g.RubricEngine, F = g.FeedbackBuilder;
const profile = R.buildProfileFromRubricRows(rows, { taskName: 'Reference', name: 'Ref', weights: LEGACY_WEIGHTS, bands: BANDS });

function build(states, extra) { return F.build(profile, states, R.scoreProfile(profile, states), extra || {}); }

test('feedback uses only ticked observables, next-band targets and teacher observations (SC31, SC33)', () => {
  const states = { C01: { checkboxes: { 'C01-B-1': true }, teacherWrittenNote: 'explained sources well in class discussion' }, C02: { checkboxes: { 'C02-C-1': true, 'C02-C-2': true } } };
  const fb = build(states, { generalObservation: 'worked safely with the coping saw' });
  assert.match(fb.whatWentWell, /Research: Explains relevant sources\./);
  assert.match(fb.whatWentWell, /Research \(observed\): Explained sources well in class discussion\./);
  assert.match(fb.whatWentWell, /Observed in class: Worked safely with the coping saw\./);
  assert.match(fb.areasForImprovement, /Research: next, aim to show that you evaluates several sources/i);
  assert.match(fb.areasForImprovement, /Making: next, aim to show that you accurate measuring/i);
  // Nothing un-ticked is presented as achieved.
  assert.doesNotMatch(fb.whatWentWell, /Links research to design decisions|High-quality finish/);
  assert.equal(fb.goalsForNextAssessment.split('\n').length, 2);
});

test('feedback changes with the assessment, so it cannot contradict revised selections (SC32)', () => {
  const before = build({ C01: { checkboxes: { 'C01-A-1': true, 'C01-A-2': true } }, C02: { checkboxes: { 'C02-A-1': true } } });
  const after = build({ C01: { checkboxes: { 'C01-D-1': true } }, C02: { checkboxes: { 'C02-A-1': true } } });
  assert.match(before.whatWentWell, /Evaluates several sources/);
  assert.doesNotMatch(after.whatWentWell, /Evaluates several sources/);
  assert.match(after.whatWentWell, /Lists a source/);
});

test('rendered comment uses the first name once, in the opening, and follows output settings (SC34, SC35)', () => {
  const fb = { whatWentWell: 'Research: Explains relevant sources.', areasForImprovement: 'Making: next, aim to finish edges.', goalsForNextAssessment: '1. Making — Finish.' };
  const ctx = { studentName: 'Mia Taylor', taskName: 'Bridge', letter: 'B', points: 31.375, maxMarks: 40, percent: 78.4375 };
  const out = F.render('', fb, ctx, { useFirstName: true, gradeDisplay: 'letter+mark', includeGradeInComment: true });
  assert.match(out.text, /^Mia, here is your feedback for Bridge\./);
  assert.equal((out.text.match(/Mia/g) || []).length, 1);
  assert.doesNotMatch(out.text, /Taylor/);
  assert.match(out.text, /Result: B \(31\.4 \/ 40\)/);
  const plain = F.render('{{opening}} {{grade_display}}', fb, ctx, { useFirstName: false, gradeDisplay: 'percent' });
  assert.equal(plain.text, 'Here is your feedback for Bridge. 78.4%');
  const noGrade = F.render('', fb, ctx, { useFirstName: true, includeGradeInComment: false });
  assert.doesNotMatch(noGrade.text, /Result:/);
  const twice = F.render('{{opening}} Well done {{first_name}}.', fb, ctx, { useFirstName: true });
  assert.ok(twice.warnings.some(w => /appears 2 times/.test(w)));
});
