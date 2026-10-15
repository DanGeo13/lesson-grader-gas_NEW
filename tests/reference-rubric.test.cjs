/* Hand-worked reference cases. Every expected number below is derived in
   docs/REFERENCE_RUBRIC.md; if a test here changes, that document must too. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context } = require('./helpers.cjs');
const { rows, LEGACY_WEIGHTS, BANDS } = require('./fixtures/reference-rubric.cjs');
const g = context().load('RubricEngine.gs');
const R = g.RubricEngine;

function profile(weights) {
  return R.buildProfileFromRubricRows(rows, { taskName: 'Reference', name: 'Reference rubric', weights, bands: BANDS });
}
function ticks(...ids) { const m = {}; ids.forEach(id => { m[id] = true; }); return m; }
function score(p, c1, c2) { return R.scoreProfile(p, { C01: { checkboxes: c1 }, C02: { checkboxes: c2 } }); }
const close = (a, b) => assert.ok(Math.abs(a - b) < 1e-9, `${a} != ${b}`);

test('R1 unequal criterion weights and partial achievement (legacy weights)', () => {
  const p = profile(LEGACY_WEIGHTS);
  assert.equal(p.totalMaxMarks, 40);
  const r = score(p, ticks('C01-A-1'), ticks('C02-A-1', 'C02-A-2'));
  close(r.criteria.C01.points, 5);      // 1 of 2 A observables x 10 x 1.0 / 2
  assert.equal(r.criteria.C01.grade, 'D'); // 50% -> D (>= 50)
  close(r.criteria.C02.points, 20);     // 2 of 3 x 30 / 3
  assert.equal(r.criteria.C02.grade, 'C'); // 66.67%
  close(r.totalPoints, 25);
  close(r.percent, 62.5);
  assert.equal(r.letter, 'D');
});

test('R2 A/C/E reference points: A matches target; C and E conflict under legacy weights', () => {
  const p = profile(LEGACY_WEIGHTS);
  const a = score(p, ticks('C01-A-1', 'C01-A-2'), ticks('C02-A-1', 'C02-A-2', 'C02-A-3'));
  close(a.criteria.C01.percent, 100); close(a.criteria.C02.percent, 100);
  const c = score(p, ticks('C01-C-1', 'C01-C-2'), ticks('C02-C-1', 'C02-C-2', 'C02-C-3'));
  close(c.criteria.C01.percent, 70); close(c.criteria.C02.percent, 70);   // target 50
  const e = score(p, ticks('C01-E-1'), ticks('C02-E-1', 'C02-E-2'));
  close(e.criteria.C01.percent, 25); close(e.criteria.C02.percent, 25);   // target 0
  const report = R.referencePointReport(LEGACY_WEIGHTS);
  assert.equal(report.matchesTarget, false);
  assert.equal(report.differences.map(d => d.band).join(','), 'C,E');
});

test('R3 target rule set reproduces A=100%, C=50%, E=0% and refuses unconfirmed B/D', () => {
  assert.throws(() => R.targetRuleSetWeights('', ''), /not confirmed/);
  assert.throws(() => R.targetRuleSetWeights(0.4, 0.25), /between/);
  // 0.75 and 0.25 are ILLUSTRATIVE ONLY — not confirmed policy.
  const w = R.targetRuleSetWeights(0.75, 0.25);
  const p = profile(w);
  const a = score(p, ticks('C01-A-1', 'C01-A-2'), ticks('C02-A-1', 'C02-A-2', 'C02-A-3'));
  const c = score(p, ticks('C01-C-1', 'C01-C-2'), ticks('C02-C-1', 'C02-C-2', 'C02-C-3'));
  const e = score(p, ticks('C01-E-1'), ticks('C02-E-1', 'C02-E-2'));
  close(a.totalPoints, 40); close(c.totalPoints, 20); close(e.totalPoints, 0);
  // With the CURRENT overall bands, a full-C portfolio (50%) is a D: the
  // overall boundaries would also need confirming before adopting the target.
  assert.equal(c.letter, 'D');
  assert.equal(R.referencePointReport(w).matchesTarget, true);
});

test('R4 confirmed B/D behaviour under legacy weights (as implemented, policy unconfirmed)', () => {
  const p = profile(LEGACY_WEIGHTS);
  const b = score(p, ticks('C01-B-1', 'C01-B-2'), {});
  close(b.criteria.C01.points, 8.75);
  assert.equal(b.criteria.C01.grade, 'A');   // 87.5% >= 85: a full-B criterion reads as A
  const d = score(p, ticks('C01-D-1'), {});
  close(d.criteria.C01.points, 5.75);
  assert.equal(d.criteria.C01.grade, 'D');
});

test('R5 teacher overrides: refining AI ticks and the distinct override', () => {
  const p = profile(LEGACY_WEIGHTS);
  const ai = score(p, ticks('C01-B-1'), ticks('C02-A-1', 'C02-A-2', 'C02-A-3'));
  close(ai.totalPoints, 4.375 + 30);
  // Teacher removes one unsupported A tick and records the C evidence seen in class.
  const teacher = score(p, ticks('C01-B-1'), ticks('C02-A-1', 'C02-A-2', 'C02-C-1'));
  close(teacher.criteria.C02.points, 20 + 7);
  close(teacher.totalPoints, 4.375 + 27);
  // Distinct override awards full criterion marks regardless of other ticks.
  const md = score(p, ticks('C01-MD'), {});
  close(md.criteria.C01.points, 10);
  assert.equal(md.criteria.C01.distinctOverride, true);
});

test('R6 multi-band ticks are capped at the criterion maximum', () => {
  const p = profile(LEGACY_WEIGHTS);
  const r = score(p, ticks('C01-A-1', 'C01-A-2', 'C01-B-1', 'C01-B-2'), {});
  close(r.criteria.C01.rawPoints, 18.75);
  close(r.criteria.C01.points, 10);
});

test('R7 grade boundaries: below, at and above every boundary; float noise cannot drop a grade', () => {
  const cases = [[100, 'A'], [85.01, 'A'], [85, 'A'], [84.99, 'B'], [75, 'B'], [74.99, 'C'], [65, 'C'], [64.99, 'D'], [50, 'D'], [49.99, 'E'], [0, 'E']];
  for (const [pct, letter] of cases) assert.equal(R.overallLetter(pct, BANDS), letter, `${pct}%`);
  assert.equal(R.overallLetter(84.99999999999999, BANDS), 'A');
  assert.equal(R.overallLetter(64.99999999999999, BANDS), 'C');
  // 0.1 + 0.2 style accumulation: ten C ticks worth 0.7 each sum to 6.999999999999999.
  const noisy = R.buildProfileFromRubricRows(Array.from({ length: 10 }, (_, i) => ({ criterion: 'N', maxMarks: 10, band: 'C', description: 'obs ' + i })), { taskName: 'T', name: 'N', weights: LEGACY_WEIGHTS, bands: BANDS });
  const all = {}; noisy.criteria[0].bands.C.forEach(o => { all[o.id] = true; });
  const res = R.scoreProfile(noisy, { C01: { checkboxes: all } });
  assert.equal(res.percent, 70);
});

test('R8 rounding: Classroom grade is half-up to 1 d.p.; letters use unrounded values', () => {
  assert.equal(R.roundHalfUp(2.675, 2), 2.68);
  assert.equal(R.roundHalfUp(2.45, 1), 2.5);
  assert.equal(R.classroomGrade(25, 40, 100, 'scaled-points').grade, 62.5);
  assert.equal(R.classroomGrade(31.375, 40, 20, 'scaled-points').grade, 15.7); // 15.6875 -> 15.7
  assert.equal(R.classroomGrade(25, 40, 40, 'raw-points').grade, 25);
  assert.throws(() => R.classroomGrade(25, 40, 20, 'raw-points'), /scaled/);
  assert.throws(() => R.classroomGrade(25, 40, 0, 'scaled-points'), /ungraded/);
  assert.equal(R.classroomGrade(25, 40, 100, 'none').grade, null);
  // 84.996% displays as 85.0 but remains a B: display rounding never changes the letter.
  assert.equal(R.overallLetter(84.996, BANDS), 'B');
});

test('R9 reproducibility: identical selections give identical results regardless of feedback text', () => {
  const p = profile(LEGACY_WEIGHTS);
  const states = fb => ({ C01: { checkboxes: ticks('C01-A-1'), teacherWrittenNote: fb }, C02: { checkboxes: ticks('C02-C-1'), teacherWrittenNote: fb + '!' } });
  const one = R.scoreProfile(p, states('Excellent work'));
  const two = R.scoreProfile(p, states('Needs a lot more effort — completely different words'));
  assert.deepEqual(JSON.parse(JSON.stringify(one)), JSON.parse(JSON.stringify(two)));
});
