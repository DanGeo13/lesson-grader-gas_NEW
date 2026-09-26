const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context, Workbook, objectRows } = require('./helpers.cjs');
const g = context().load('RubricEngine.gs');
const E=g.RubricEngine;

function makeProfile(task='Rube Goldberg') {
  return E.buildProfileFromRubricRows([
    {criterion:'Design process',part:'Research',maxMarks:8,band:'A',description:'Research informs design choices'},
    {criterion:'Design process',part:'Research',maxMarks:8,band:'A',description:'Ideation is clearly annotated'},
    {criterion:'Design process',part:'Research',maxMarks:8,band:'B',description:'Relevant design research'},
    {criterion:'Practical outcome',part:'Making',maxMarks:12,band:'A',description:'Mechanism works reliably'}
  ],{taskName:task,name:'Rube Goldberg Rubric'});
}

test('imports arbitrary project criteria, parts and mark totals — never jewellery defaults',()=>{
  const p=makeProfile();
  assert.equal(E.validateProfile(p).ok,true);
  assert.equal(p.criteria.length,2);
  assert.equal(p.totalMaxMarks,20);
  assert.deepEqual(Array.from(p.parts),['Research','Making']);
  assert.equal(p.criteria[0].bands.A[0].id,'C01-A-1');
  assert.ok(!E.serializeForAi(p).includes('Jewellery'));
  assert.ok(E.serializeForAi(p).includes('Mechanism works reliably'));
});

test('deterministic scores match legacy weights for a single band; multi-band values cap at full marks',()=>{
  const p=makeProfile();
  const c=p.criteria[0];
  assert.equal(E.scoreCriterion(c,{ 'C01-A-1':true }).points,4);
  const all=E.scoreCriterion(c,{ 'C01-A-1':true,'C01-A-2':true,'C01-B-1':true });
  assert.equal(all.points,8);
  assert.ok(all.rawPoints>all.maxMarks);
  assert.equal(all.grade,'A');
  const md=E.scoreCriterion(c,{'C01-MD':true});
  assert.equal(md.points,8);
  assert.equal(md.score,1);
  const overall=E.scoreProfile(p,{C01:{checkboxes:{'C01-A-1':true}},C02:{checkboxes:{'C02-A-1':true}}});
  assert.equal(overall.totalPoints,16);
  assert.equal(overall.totalMaxMarks,20);
  assert.equal(overall.letter,'B');
});

test('unfamiliar bands require explicit teacher weight; missing marks block activation',()=>{
  const rows=[{criterion:'Water filter',maxMarks:15,band:'LEVEL 4',description:'Clean and clear output'}];
  const unknown=E.buildProfileFromRubricRows(rows,{taskName:'Water filter',name:'Water rubric'});
  assert.equal(E.validateProfile(unknown).ok,false);
  assert.match(E.validateProfile(unknown).errors.join(' '),/explicit 0–1 weight/);
  const reviewed=E.buildProfileFromRubricRows(rows,{taskName:'Water filter',name:'Water rubric',weights:{'LEVEL 4':0.9}});
  assert.equal(E.validateProfile(reviewed).ok,true);
  assert.equal(E.scoreProfile(reviewed,{C01:{checkboxes:{'C01-LEVEL 4-1':true}}}).totalPoints,13.5);
  const missing=E.buildProfileFromRubricRows([{...rows[0],maxMarks:0,band:'A'}],{taskName:'Water filter',name:'Water rubric'});
  assert.equal(E.validateProfile(missing).ok,false);
  assert.throws(()=>E.buildProfileFromRubricRows([{...rows[0],description:''}],{taskName:'Water',name:'Water'}),/needs a criterion, band and observable/);
});

test('migrates a legacy four-column RubricProfiles sheet by header, never overwrites prior JSON',()=>{
  let n=1;
  g.Utilities={getUuid:()=>`PROFILE-${n++}`};
  const ss=new Workbook();
  const original='{ "legacy": true }';
  ss.add('RubricProfiles',[
    ['ProfileID','ProfileName','JSONDefinition','CreatedAt'],
    ['LEGACY','Older draft',original,'2025-01-01']
  ]);
  E.ensureProfilesSheet(ss);
  assert.ok(ss.getSheetByName('RubricProfiles').rows[0].includes('TaskName'));
  const p=makeProfile();
  assert.equal(E.saveProfile(ss,p,'teacher@school.edu.au').success,true);
  assert.equal(E.getActiveProfile(ss,p.taskName).name,p.name);
  const rows=objectRows(ss.getSheetByName('RubricProfiles'));
  assert.equal(rows[0].JSONDefinition,original);
  assert.equal(rows[1].ProfileName,p.name);
  assert.equal(rows[1].TaskName,p.taskName);
  assert.equal(rows[1].Active,true);
  const p2=makeProfile();p2.name='Updated version';
  E.saveProfile(ss,p2,'teacher@school.edu.au');
  const all=objectRows(ss.getSheetByName('RubricProfiles'));
  assert.equal(all[1].Active,false);
  assert.equal(all[2].Active,true);
  assert.equal(E.getProfileById(ss,p.profileId).name,p.name);
  assert.equal(E.getActiveProfile(ss,p.taskName).name,'Updated version');
});

test('report escapes untrusted student text and uses task name from profile',()=>{
  const p=makeProfile();
  const html=E.renderReportHtml(p,{submission:{StudentName:'<script>alert(1)</script>',Class:'7TECHI',SubmissionVersion:1},criteriaMap:{},feedback:{whatWentWell:'<img onerror=alert(1)>'}},null,{taskName:p.taskName});
  assert.ok(html.includes('Rube Goldberg'));
  assert.ok(!html.includes('<script>'));
  assert.ok(!html.includes('<img onerror'));
  assert.ok(html.includes('&lt;script&gt;'));
});

test('same title in two project parts becomes two distinct criteria, not merged',()=>{
  const p=E.buildProfileFromRubricRows([
    {criterion:'Reflection',part:'Design folio',maxMarks:5,band:'A',description:'Explains research decisions'},
    {criterion:'Reflection',part:'Making',maxMarks:10,band:'B',description:'Evaluates prototype testing'}
  ],{taskName:'Independent project',name:'Project rubric'});
  assert.equal(p.criteria.length,2);
  assert.equal(p.criteria[0].part,'Design folio');
  assert.equal(p.criteria[1].part,'Making');
  assert.equal(p.totalMaxMarks,15);
});

test('validates editable A-E thresholds and rejects gaps, overlaps, or inverted bounds',()=>{
  const valid = [
    { letter: 'A', min: 85, max: 100 },
    { letter: 'B', min: 75, max: 85 },
    { letter: 'C', min: 65, max: 75 },
    { letter: 'D', min: 50, max: 65 },
    { letter: 'E', min: 0, max: 50 }
  ];
  assert.equal(E.validateThresholds(valid).ok, true);

  const overlap = [
    { letter: 'A', min: 80, max: 100 },
    { letter: 'B', min: 85, max: 90 },
    { letter: 'C', min: 65, max: 80 },
    { letter: 'D', min: 50, max: 65 },
    { letter: 'E', min: 0, max: 50 }
  ];
  assert.equal(E.validateThresholds(overlap).ok, false);
  assert.match(E.validateThresholds(overlap).errors.join(' '), /overlaps/);

  const gap = [
    { letter: 'A', min: 85, max: 100 },
    { letter: 'B', min: 70, max: 80 }, // gap between 80 and 85
    { letter: 'C', min: 60, max: 70 },
    { letter: 'D', min: 50, max: 60 },
    { letter: 'E', min: 0, max: 50 }
  ];
  assert.equal(E.validateThresholds(gap).ok, false);
  assert.match(E.validateThresholds(gap).errors.join(' '), /Gap detected/);

  const invalidMax = [
    { letter: 'A', min: 85, max: 95 }, // not 100
    { letter: 'B', min: 75, max: 85 },
    { letter: 'C', min: 65, max: 75 },
    { letter: 'D', min: 50, max: 65 },
    { letter: 'E', min: 0, max: 50 }
  ];
  assert.equal(E.validateThresholds(invalidMax).ok, false);
});

test('previews threshold recalculation effects on student grade distributions',()=>{
  const oldBands = [
    { letter: 'A', min: 85, max: 100 },
    { letter: 'B', min: 75, max: 85 },
    { letter: 'C', min: 65, max: 75 },
    { letter: 'D', min: 50, max: 65 },
    { letter: 'E', min: 0, max: 50 }
  ];
  const newBands = [
    { letter: 'A', min: 80, max: 100 }, // lowered to 80
    { letter: 'B', min: 70, max: 80 },
    { letter: 'C', min: 60, max: 70 },
    { letter: 'D', min: 50, max: 60 },
    { letter: 'E', min: 0, max: 50 }
  ];
  const sampleScores = [
    { studentName: 'Alex', percent: 82 }, // was B, becomes A
    { studentName: 'Jordan', percent: 72 }, // was C, becomes B
    { studentName: 'Taylor', percent: 90 }  // was A, stays A
  ];
  const preview = E.previewThresholdRecalculation(oldBands, newBands, sampleScores);
  assert.equal(preview.totalEvaluated, 3);
  assert.equal(preview.changedCount, 2);
  assert.equal(preview.changes[0].oldGrade, 'B');
  assert.equal(preview.changes[0].newGrade, 'A');
});

test('computes deterministic deduction rules: late days, fixed marks, percentage per day, weekend exclusion',()=>{
  // 1. Percentage per day: 10%/day on a 20 mark task, 2 days late -> -4 marks
  const res1 = E.calculateDeduction(18, 20, { type: 'percent_per_day', rate: 10 }, {
    turnedInTime: '2026-09-20T10:00:00Z',
    dueDate: '2026-09-18T10:00:00Z'
  });
  assert.equal(res1.lateDays, 2);
  assert.equal(res1.deductionPoints, 4);
  assert.equal(res1.finalScore, 14);
  assert.equal(res1.penaltyApplied, true);

  // 2. Fixed marks deduction: -5 marks
  const res2 = E.calculateDeduction(16, 20, { type: 'fixed_marks', rate: 5, reason: 'Late folio' }, {});
  assert.equal(res2.deductionPoints, 5);
  assert.equal(res2.finalScore, 11);

  // 3. Percentage of achieved score: 10% off 15 marks -> -1.5 marks
  const res3 = E.calculateDeduction(15, 20, { type: 'percent_of_achieved', rate: 10 }, {});
  assert.equal(res3.deductionPoints, 1.5);
  assert.equal(res3.finalScore, 13.5);

  // 4. On time submission with grace period
  const res4 = E.calculateDeduction(18, 20, { type: 'percent_per_day', rate: 10, graceHours: 4 }, {
    turnedInTime: '2026-09-18T12:00:00Z',
    dueDate: '2026-09-18T10:00:00Z'
  });
  assert.equal(res4.penaltyApplied, false);
  assert.equal(res4.finalScore, 18);

  // 5. Floor constraint (never below minFloor)
  const res5 = E.calculateDeduction(3, 20, { type: 'fixed_marks', rate: 10, minFloor: 0 }, {});
  assert.equal(res5.finalScore, 0);
});

test('renders feedback templates and warns about unknown tokens',()=>{
  const template = 'Great job {{student}} on {{task}}! Grade: {{grade}} ({{final_score}}/{{max_marks}}). {{what_went_well}} Areas to improve: {{areas_for_improvement}}. {{unknown_field}}';
  const data = {
    student: 'Sam',
    task: 'Jewellery Folio',
    grade: 'A',
    finalScore: 18,
    maxMarks: 20,
    whatWentWell: 'Excellent casting technique.',
    areasForImprovement: 'Refine the polish finish.'
  };
  const rendered = E.renderFeedbackTemplate(template, data);
  assert.ok(rendered.text.includes('Great job Sam on Jewellery Folio! Grade: A (18.0/20). Excellent casting technique.'));
  assert.ok(rendered.text.includes('Refine the polish finish.'));
  assert.equal(rendered.warnings.length, 1);
  assert.match(rendered.warnings[0], /unknown_field/);
});

test('validates curriculum outcome mappings and tallies outcome coverage',()=>{
  const criteria = [
    { criterionId: 'C01', maxMarks: 10, outcome: 'DT5-1' },
    { criterionId: 'C02', maxMarks: 10, outcome: 'DT5-2' },
    { criterionId: 'C03', maxMarks: 5, outcome: '' }
  ];
  const res = E.validateOutcomeMappings(criteria, ['DT5-1', 'DT5-2', 'DT5-3']);
  assert.equal(res.valid, false);
  assert.equal(res.mappedCriteria, 2);
  assert.equal(res.unmappedCriteria[0], 'C03');

  const p = makeProfile();
  p.criteria[0].outcome = 'DT5-1';
  p.criteria[1].outcome = 'DT5-2';
  const cov = E.calculateOutcomeCoverage(p, [
    { criterionId: 'C01', points: 8, teacherWrittenNote: 'Observed' },
    { criterionId: 'C02', points: 10, teacherWrittenNote: 'Tested' }
  ]);
  assert.equal(cov['DT5-1'].criteriaCount, 1);
  assert.equal(cov['DT5-1'].totalMarksEarned, 8);
  assert.equal(cov['DT5-2'].totalMarksEarned, 10);
});

test('generates formative rubric draft with discrete criteria, bands and outcomes',()=>{
  const draft = E.generateFormativeRubricDraft({
    title: 'Water Filtration Interim',
    description: 'Design and test a sand and charcoal filter column.',
    maxMarks: 15,
    courseCode: '8TECHI',
    outcomes: ['SC4-14LW', 'DT5-2', 'DT5-3']
  });
  assert.equal(draft.taskName, 'Water Filtration Interim');
  assert.equal(draft.reviewed, false);
  assert.equal(draft.assessmentType, 'Formative');
  assert.equal(draft.criteria.length, 3);
  assert.equal(draft.totalMaxMarks, 15);
  assert.equal(draft.criteria[0].outcome, 'SC4-14LW');
});

