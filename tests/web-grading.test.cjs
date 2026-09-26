const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context, Workbook, objectRows } = require('./helpers.cjs');

function setup() {
  const ss = new Workbook();
  ss.add('Rube Goldberg',[['Task']]);
  ss.add('Water Filter',[['Task']]);
  // Two web-app-era tables: only minimal headers. prepare() must add missing
  // headers without destroying these rows or changing their existing order.
  ss.add('Submissions',[
    ['SubmissionRecordID','StudentUserID','Status','Task','StudentName','SubmissionVersion','CurrentOfficial'],
    ['SUB-1','STU-1','New','Rube Goldberg','Mia Taylor',1,true]
  ]);
  ss.add('ClassroomConfig',[['ConfigID','CourseID','CourseName','CourseSection','CourseWorkID','AssignmentTitle','MaxPoints','SavedAt','SavedBy','Active','TaskName','DueDate','AutoImported']]);
  ss.add('RubricProfiles',[['ProfileID','ProfileName','JSONDefinition','CreatedAt']]);
  let id=1;
  const mocks={
    Utilities:{getUuid:()=>`id-${id++}`},
    SpreadsheetApp:{flush(){}},
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    WebAccess:{requireTeacher:()=> 'dan@school.edu.au'},
    LessonGraderWeb:{open:(wb)=> {assert.equal(wb,ss.id);return ss;},taskNames:()=>['Rube Goldberg','Water Filter'],isSharedRubric:()=>true},
    DriveService:{isEligibleForAi:(mime)=>mime === 'application/pdf'},
    DriveApp:{getFileById:()=>({makeCopy:()=>({getId:()=> 'backup-1'})})}
  };
  const g=context(mocks).load('RubricEngine.gs','WebGrading.gs');
  return {g,ss};
}
function rowsOne() {return [
  {criterion:'Research',part:'Process',maxMarks:10,band:'A',description:'Evidence of independent research'},
  {criterion:'Research',part:'Process',maxMarks:10,band:'B',description:'Some relevant research'},
  {criterion:'Making',part:'Practical',maxMarks:15,band:'A',description:'Working prototype'}
];}

test('web workflow: prepare, activate own rubric, draft, approve, lock, reassess, pin version',()=>{
  const {g,ss}=setup();const W=g.WebGrading;
  assert.equal(W.listTask(ss.id,'Rube Goldberg').data.prepared,false);
  assert.equal(W.prepare(ss,false).success,true);
  assert.equal(W.listTask(ss.id,'Rube Goldberg').data.prepared,true);
  const saved=W.activateRubric(ss.id,'Rube Goldberg','Rube Goldberg rubric','',rowsOne(),{},true);
  assert.equal(saved.success,true);
  assert.equal(saved.data.totalMaxMarks,25);
  const d=W.detail(ss.id,'Rube Goldberg','SUB-1').data;
  assert.equal(d.profile.criteria.length,2);
  assert.equal(d.assessmentId,'');
  assert.equal(d.revision,'');
  assert.throws(()=>W.approve(ss.id,'Rube Goldberg','SUB-1',''),/Save a complete draft/);
  const draft=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{
    C01:{checkboxes:{'C01-A-1':true},teacherWrittenNote:'Research observed.'},
    C02:{checkboxes:{'C02-A-1':true},teacherWrittenNote:'Prototype works.'}
  },{whatWentWell:'Good engineering',areasForImprovement:'Measure more',goalsForNextAssessment:'Test early'},'', 'Teacher');
  assert.equal(draft.success,true);
  assert.equal(draft.scores.totalPoints,25);
  assert.equal(objectRows(ss.getSheetByName('CriterionAssessments'))[0].FinalApprovedGrade,'');
  assert.throws(()=>W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{}, {},'outdated','Teacher'),/Reload/);
  const a=W.approve(ss.id,'Rube Goldberg','SUB-1',draft.revision);
  assert.equal(a.success,true);
  assert.equal(a.scores.totalMaxMarks,25);
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades'))[0].Grade,'A');
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].Status,'Approved');
  assert.throws(()=>W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{}, {},a.revision,'Teacher'),/final/);
  const oldId=d.profile.profileId;
  W.lockAssessment(ss.id,'Rube Goldberg','SUB-1');
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].CurrentOfficial,true);
  W.activateRubric(ss.id,'Rube Goldberg','Newer rubric','',[...rowsOne(),{criterion:'Presentation',maxMarks:5,band:'A',description:'Clear communication'}],{},true);
  assert.equal(W.detail(ss.id,'Rube Goldberg','SUB-1').data.profile.profileId,oldId);
  const r=W.reassess(ss.id,'Rube Goldberg','SUB-1');
  assert.equal(r.success,true);
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].CurrentOfficial,false);
  assert.equal(W.detail(ss.id,'Rube Goldberg',r.submissionRecordId).data.profile.criteria.length,3);
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades')).length,1);
  assert.throws(()=>W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{}, {},'','Teacher'),/final/);
  assert.equal(W.listTask(ss.id,'Rube Goldberg').data.stats.total,1);
});

test('cannot activate a malformed/unreviewed rubric and cannot grade unknown checkboxes',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  assert.throws(()=>W.activateRubric(ss.id,'Water Filter','Rubric','',rowsOne(),{},false),/Review/);
  assert.throws(()=>W.activateRubric(ss.id,'Water Filter','Rubric','',[{criterion:'Flow',maxMarks:0,band:'A',description:'Works'}],{},true),/mark allocation/);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  assert.throws(()=>W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{C01:{checkboxes:{'C01-FAKE':true}}}, {},'','Teacher'),/Unknown observable/);
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].Status,'New');
});

test('task-scoped import keeps two assignments with same student separate and versions changed files',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  const config=ss.getSheetByName('ClassroomConfig');
  config.appendRow(['cfg1','COURSE-1','7TECHI','','WORK-1','Rube Goldberg',100,'','',true,'Rube Goldberg','','']);
  config.appendRow(['cfg2','COURSE-1','7TECHI','','WORK-2','Water Filter',100,'','',true,'Water Filter','','']);
  const payload={
    'WORK-1':[ {id:'CLASS-SUB-1',userId:'STU-2',state:'TURNED_IN',assignmentSubmission:{attachments:[{driveFile:{id:'FILE-A',title:'folio.pdf',mimeType:'application/pdf',alternateLink:'https://drive.google.com/a'}}]}} ],
    'WORK-2':[ {id:'CLASS-SUB-2',userId:'STU-2',state:'TURNED_IN',assignmentSubmission:{attachments:[{driveFile:{id:'FILE-B',title:'prototype.pdf',mimeType:'application/pdf',alternateLink:'https://drive.google.com/b'}}]}} ]
  };
  g.Classroom={Courses:{CourseWork:{StudentSubmissions:{list:(_,work)=>({studentSubmissions:payload[work]})}}},UserProfiles:{get:()=>({name:{fullName:'Rory Smith'},emailAddress:'rory@school.edu.au'})}};
  assert.equal(W.importTask(ss.id,'Rube Goldberg').stats.newCount,1);
  assert.equal(W.importTask(ss.id,'Water Filter').stats.newCount,1);
  let all=objectRows(ss.getSheetByName('Submissions')).filter(s=>s.StudentUserID==='STU-2');
  assert.equal(all.length,2);
  assert.deepEqual(all.map(r=>r.Task).sort(),['Rube Goldberg','Water Filter']);
  assert.equal(W.importTask(ss.id,'Rube Goldberg').stats.unchanged,1);
  payload['WORK-1'][0].assignmentSubmission.attachments[0].driveFile.id='FILE-C';
  assert.equal(W.importTask(ss.id,'Rube Goldberg').stats.newVersions,1);
  all=objectRows(ss.getSheetByName('Submissions')).filter(s=>s.StudentUserID==='STU-2');
  const other=all.find(s=>s.Task==='Water Filter');
  assert.equal(other.AttachmentFileIDsJSON,'["FILE-B"]');
  assert.equal(all.filter(s=>s.Task==='Rube Goldberg' && s.CurrentOfficial===true).length,1);
  assert.equal(all.filter(s=>s.Task==='Rube Goldberg' && s.CurrentOfficial===false).length,1);
});

test('enable-grading backup failure stops all sheet migration',()=>{
  const {g,ss}=setup();const W=g.WebGrading;
  let sawBlank=false;
  g.DriveApp={getFileById:()=>({makeCopy:()=>{
    sawBlank=ss.getSheetByName('ApprovedGrades')===null;
    throw new Error('Drive backup blocked');
  }})};
  assert.throws(()=>W.prepare(ss,true),/Drive backup blocked/);
  assert.equal(sawBlank,true);
  assert.equal(ss.getSheetByName('ApprovedGrades'),null);
  g.DriveApp={getFileById:()=>({makeCopy:()=>({getId:()=> 'BACKUP-ID'})})};
  const enabled=W.prepare(ss,true);
  assert.equal(enabled.backupId,'BACKUP-ID');
  assert.equal(W.listTask(ss.id,'Rube Goldberg').data.prepared,true);
});

test('task-aware AI drafts only valid evidence; logs run without approving or changing Classroom',()=>{
  const {g,ss}=setup();const W=g.WebGrading;
  W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  // Append a file through the prepared sheet's named columns.
  const sh=ss.getSheetByName('SubmissionFiles');
  sh.appendRow(sh.rows[0].map(key=>({SubmissionRecordID:'SUB-1',DriveFileID:'FILE-1',FileName:'folio.pdf',MimeType:'application/pdf'})[key]??''));
  g.DriveService.prepareSingleFileForAi=()=>({status:'reviewed',fileId:'FILE-1',fileName:'folio.pdf',mimeType:'application/pdf',inlineData:{data:'dGVzdA=='}});
  let prompt='';
  g.GeminiService={callGeminiWithFallback:(contents)=>{
    prompt=JSON.stringify(contents);
    return {success:true,modelUsed:'gemini-2.5-flash',text:JSON.stringify({criteria:[
      {criterionId:'C01',tickedCheckboxes:['C01-A-1','C01-MD','C01-FAKE'],evidenceNotes:[{checkboxId:'C01-A-1',fileName:'folio.pdf',note:'Research observed'}]},
      {criterionId:'C02',tickedCheckboxes:['C02-A-1'],evidenceNotes:[]}
    ]})};
  }};
  const result=W.runAi(ss.id,'Rube Goldberg','SUB-1','');
  assert.equal(result.success,true);
  assert.ok(prompt.includes('Research'));
  assert.ok(!prompt.includes('Jewellery Design'));
  const d=W.detail(ss.id,'Rube Goldberg','SUB-1').data;
  assert.equal(d.submission.status,'InReview');
  assert.equal(d.criteriaMap.C01.checkboxes['C01-A-1'],true);
  assert.equal(d.criteriaMap.C01.checkboxes['C01-MD'],undefined);
  assert.equal(d.criteriaMap.C01.checkboxes['C01-FAKE'],undefined);
  assert.equal(d.criteriaMap.C01.aiProposedGrade,'A');
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades')).length,0);
  assert.equal(objectRows(ss.getSheetByName('AIAssessments')).some(row=>row.RunType==='EvidenceProposal' && row.ModelUsed==='gemini-2.5-flash'),true);
});

test('manual project submission needs no Classroom record and cannot be added twice',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Water Filter','Water rubric','',rowsOne(),{},true);
  const created=W.addManualSubmission(ss.id,'Water Filter','Casey Lee','casey@school.edu.au','8TECHI');
  assert.equal(created.success,true);
  const d=W.detail(ss.id,'Water Filter',created.submissionRecordId).data;
  assert.equal(d.submission.studentName,'Casey Lee');
  assert.equal(d.submission.className,'8TECHI');
  assert.equal(d.profile.name,'Water rubric');
  const row=objectRows(ss.getSheetByName('Submissions')).find(r=>r.SubmissionRecordID===created.submissionRecordId);
  assert.equal(row.SourceType,'Manual');
  assert.equal(row.ClassroomCourseWorkID,'');
  assert.throws(()=>W.addManualSubmission(ss.id,'Water Filter','Casey Lee','casey@school.edu.au','8TECHI'),/already exists/);
  assert.equal(objectRows(ss.getSheetByName('AssessmentHistory')).find(r=>r.SubmissionRecordID===created.submissionRecordId).Action,'ManualSubmissionCreated');
});

test('Classroom import will not silently duplicate a manually entered student for the same task',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.addManualSubmission(ss.id,'Water Filter','Casey Lee','casey@school.edu.au','8TECHI');
  ss.getSheetByName('ClassroomConfig').appendRow(['cfg','COURSE-1','8TECHI','','WORK-2','Water Filter',100,'','',true,'Water Filter','','']);
  g.Classroom={Courses:{CourseWork:{StudentSubmissions:{list:()=>({studentSubmissions:[{
    id:'CLASS-SUB-1',userId:'STU-4',assignmentSubmission:{attachments:[]}
  }]})}}},UserProfiles:{get:()=>({name:{fullName:'Casey Lee'},emailAddress:'casey@school.edu.au'})}};
  assert.throws(()=>W.importTask(ss.id,'Water Filter'),/manual submission already exists/);
  assert.equal(objectRows(ss.getSheetByName('Submissions')).filter(r=>r.Task==='Water Filter').length,1);
});

test('applies late deductions to draft and approved grades without compounding on recalculation',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const draft = W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{
    C01:{checkboxes:{'C01-A-1':true}},
    C02:{checkboxes:{'C02-A-1':true}}
  },{whatWentWell:'Good',areasForImprovement:'None',goalsForNextAssessment:'Keep it up'},'', 'Teacher', {
    deductionRule: { type: 'percent_per_day', rate: 10 },
    assessmentType: 'Summative'
  });
  assert.equal(draft.success, true);
  assert.equal(draft.scores.totalPoints, 25);
  // SUB-1 is not late by default, so deduction is 0
  assert.equal(draft.deduction.deductionPoints, 0);

  // Now test with late penalty of fixed -5 marks
  const approved = W.approve(ss.id,'Rube Goldberg','SUB-1',draft.revision, {
    deductionRule: { type: 'fixed_marks', rate: 5, reason: 'Late folio' },
    assessmentType: 'Summative'
  });
  assert.equal(approved.success, true);
  assert.equal(approved.deduction.rawScore, 25);
  assert.equal(approved.deduction.deductionPoints, 5);
  assert.equal(approved.deduction.finalScore, 20);

  const gradeRow = objectRows(ss.getSheetByName('ApprovedGrades'))[0];
  assert.equal(gradeRow.RawScore, 25);
  assert.equal(gradeRow.DeductionPoints, 5);
  assert.equal(gradeRow.FinalScore, 20);
  assert.equal(gradeRow.AssessmentType, 'Summative');
});

test('bulk grading queue isolates errors per student and creates drafts for review',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  // Add another student
  const sub2 = W.addManualSubmission(ss.id, 'Rube Goldberg', 'Sam Green', 'sam@school.edu.au', '7TECHI');

  // Attach a file for SUB-1
  const sh = ss.getSheetByName('SubmissionFiles');
  sh.appendRow(sh.rows[0].map(key=>({SubmissionRecordID:'SUB-1',DriveFileID:'FILE-1',FileName:'folio.pdf',MimeType:'application/pdf'})[key]??''));
  g.DriveService.prepareSingleFileForAi=(f)=>({status:'reviewed',fileId:f.DriveFileID,fileName:f.FileName,mimeType:'application/pdf',inlineData:{data:'dGVzdA=='}});
  g.GeminiService={callGeminiWithFallback:()=>({success:true,modelUsed:'gemini-2.5-flash',text:JSON.stringify({criteria:[
    {criterionId:'C01',tickedCheckboxes:['C01-A-1'],evidenceNotes:[]},
    {criterionId:'C02',tickedCheckboxes:['C02-A-1'],evidenceNotes:[]}
  ]})})};

  // SUB-1 will succeed, sub2 has no files so will be isolated as error
  const batch = W.bulkGrade(ss.id, 'Rube Goldberg', ['SUB-1', sub2.submissionRecordId]);
  assert.equal(batch.total, 2);
  assert.equal(batch.successful, 1);
  assert.equal(batch.failed, 1);
  assert.equal(batch.items[0].success, true);
  assert.equal(batch.items[1].success, false);
});

test('verifies cohort grading complete and sets explicit audit lock',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  
  // Before grading: incomplete
  const check1 = W.verifyGradingComplete(ss.id, 'Rube Goldberg');
  assert.equal(check1.complete, false);
  assert.equal(check1.pendingCount, 1);
  assert.throws(()=>W.setGradingComplete(ss.id, 'Rube Goldberg', true), /Cannot lock cohort/);

  // Draft and approve SUB-1
  const draft = W.saveDraft(ss.id, 'Rube Goldberg', 'SUB-1', {
    C01:{checkboxes:{'C01-A-1':true}},
    C02:{checkboxes:{'C02-A-1':true}}
  }, { whatWentWell: 'Good' }, '', 'Teacher');
  W.approve(ss.id, 'Rube Goldberg', 'SUB-1', draft.revision);

  // Now verify complete
  const check2 = W.verifyGradingComplete(ss.id, 'Rube Goldberg');
  assert.equal(check2.complete, true);
  assert.equal(check2.approvedCount, 1);

  // Preview before lock
  const preview = W.setGradingComplete(ss.id, 'Rube Goldberg', false);
  assert.equal(preview.preview, true);

  // Lock with confirmation
  const locked = W.setGradingComplete(ss.id, 'Rube Goldberg', true);
  assert.equal(locked.success, true);
  assert.equal(objectRows(ss.getSheetByName('AuditLog')).some(r=>r.Action==='GradingCompleted'), true);
});

test('suggests HIGH, MED, LOW work samples and saves copies to Drive with preview confirmation',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  
  // Create 3 students with files and approved marks
  const sh = ss.getSheetByName('SubmissionFiles');
  // Student 1: 25 marks (100%)
  sh.appendRow(sh.rows[0].map(k=>({SubmissionRecordID:'SUB-1',DriveFileID:'F1',FileName:'high.pdf',MimeType:'application/pdf'})[k]??''));
  const d1 = W.saveDraft(ss.id, 'Rube Goldberg', 'SUB-1', { C01:{checkboxes:{'C01-A-1':true}}, C02:{checkboxes:{'C02-A-1':true}} }, {}, '', 'Teacher');
  W.approve(ss.id, 'Rube Goldberg', 'SUB-1', d1.revision);

  // Student 2: 17.5 marks (70%)
  const s2 = W.addManualSubmission(ss.id, 'Rube Goldberg', 'Jordan Med', 'jordan@school.edu.au', '7TECHI');
  sh.appendRow(sh.rows[0].map(k=>({SubmissionRecordID:s2.submissionRecordId,DriveFileID:'F2',FileName:'med.pdf',MimeType:'application/pdf'})[k]??''));
  const rev2 = W.detail(ss.id, 'Rube Goldberg', s2.submissionRecordId).data.revision;
  const d2 = W.saveDraft(ss.id, 'Rube Goldberg', s2.submissionRecordId, { C01:{checkboxes:{'C01-B-1':true}}, C02:{checkboxes:{'C02-A-1':true}} }, {}, rev2, 'Teacher');
  W.approve(ss.id, 'Rube Goldberg', s2.submissionRecordId, d2.revision);

  // Student 3: 10 marks (40%)
  const s3 = W.addManualSubmission(ss.id, 'Rube Goldberg', 'Taylor Low', 'taylor@school.edu.au', '7TECHI');
  sh.appendRow(sh.rows[0].map(k=>({SubmissionRecordID:s3.submissionRecordId,DriveFileID:'F3',FileName:'low.pdf',MimeType:'application/pdf'})[k]??''));
  const rev3 = W.detail(ss.id, 'Rube Goldberg', s3.submissionRecordId).data.revision;
  const d3 = W.saveDraft(ss.id, 'Rube Goldberg', s3.submissionRecordId, { C01:{checkboxes:{'C01-B-1':true}}, C02:{checkboxes:{'C02-A-1':true}} }, {}, rev3, 'Teacher', {
    deductionRule: { type: 'fixed_marks', rate: 10 }
  });
  W.approve(ss.id, 'Rube Goldberg', s3.submissionRecordId, d3.revision, {
    deductionRule: { type: 'fixed_marks', rate: 10 }
  });

  // Suggest samples
  const sugg = W.suggestWorkSamples(ss.id, 'Rube Goldberg', 'score_percentile');
  assert.equal(sugg.success, true);
  assert.equal(sugg.proposals.high.studentName, 'Mia Taylor');
  assert.equal(sugg.proposals.med.studentName, 'Jordan Med');
  assert.equal(sugg.proposals.low.studentName, 'Taylor Low');

  // Preview before saving
  const prev = W.saveWorkSamples(ss.id, 'Rube Goldberg', sugg.proposals, { highFolderId: 'H_DIR', medFolderId: 'M_DIR', lowFolderId: 'L_DIR' }, false);
  assert.equal(prev.preview, true);

  // Mock Drive copy
  g.DriveApp.getFolderById = () => ({ id: 'DIR-1' });
  g.DriveApp.getFileById = (id) => ({ getId: () => id, makeCopy: (n, f) => ({ getId: () => 'COPY-' + id, getUrl: () => 'https://drive.google.com/copy/' + id }) });

  const saved = W.saveWorkSamples(ss.id, 'Rube Goldberg', sugg.proposals, { highFolderId: 'H_DIR', medFolderId: 'M_DIR', lowFolderId: 'L_DIR' }, true);
  assert.equal(saved.success, true);
  assert.equal(saved.results.length, 3);
  assert.equal(objectRows(ss.getSheetByName('AuditLog')).filter(r=>r.Action.startsWith('SampleSaved_')).length, 3);
});

test('previews and saves Registration & Evaluation prefill with safe export and audit log',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const d1 = W.saveDraft(ss.id, 'Rube Goldberg', 'SUB-1', { C01:{checkboxes:{'C01-A-1':true}}, C02:{checkboxes:{'C02-A-1':true}} }, {}, '', 'Teacher');
  W.approve(ss.id, 'Rube Goldberg', 'SUB-1', d1.revision);

  const preview = W.previewRegistrationPrefill(ss.id, 'Rube Goldberg', 'DOC-123', { dates: 'Term 1 2026' });
  assert.equal(preview.success, true);
  assert.equal(preview.proposed.unitName, 'Rube Goldberg');
  assert.ok(preview.proposed.gradeDistribution.includes('A: 1'));

  const saved = W.saveRegistrationPrefill(ss.id, 'Rube Goldberg', 'DOC-123', preview.proposed, true);
  assert.equal(saved.success, true);
  assert.ok(saved.exportMarkdown.includes('# Registration & Evaluation — Rube Goldberg'));
  assert.equal(objectRows(ss.getSheetByName('AuditLog')).some(r=>r.Action==='RegistrationPrefillSaved'), true);
});

test('flags incomplete submissions, drafts and sends Classroom messages with confirmation',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  const flags = W.flagIncompleteSubmissions(ss.id, 'Rube Goldberg');
  assert.equal(flags.success, true);
  assert.equal(flags.flags.notSubmitted.length, 1);

  const drafts = W.draftClassroomMessages(ss.id, 'Rube Goldberg', flags.flags.notSubmitted, 'Please submit by Friday.');
  assert.equal(drafts.drafts.length, 1);
  assert.ok(drafts.drafts[0].message.includes('Please submit by Friday.'));

  // Preview before sending
  const prev = W.sendClassroomMessages(ss.id, 'Rube Goldberg', drafts.drafts, false);
  assert.equal(prev.preview, true);

  // Send with confirmation
  const sent = W.sendClassroomMessages(ss.id, 'Rube Goldberg', drafts.drafts, true);
  assert.equal(sent.success, true);
  assert.equal(objectRows(ss.getSheetByName('ClassroomMessages')).length, 1);
});

test('retrieves class table matrix and syncs persistent ClassTracking sheet idempotently',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const d1 = W.saveDraft(ss.id, 'Rube Goldberg', 'SUB-1', { C01:{checkboxes:{'C01-A-1':true}}, C02:{checkboxes:{'C02-A-1':true}} }, {}, '', 'Teacher', {
    deductionRule: null, assessmentType: 'Summative'
  });
  W.approve(ss.id, 'Rube Goldberg', 'SUB-1', d1.revision);

  const table = W.getClassTableData(ss.id, 'Rube Goldberg');
  assert.equal(table.success, true);
  assert.equal(table.rows.length, 1);
  assert.equal(table.rows[0].studentName, 'Mia Taylor');
  assert.equal(table.rows[0].finalScore, 25);
  assert.equal(table.rows[0].grade, 'A');

  const sync = W.syncClassTable(ss.id, 'Rube Goldberg', { year: '2026', stage: 'Stage 5', courseCode: '7TECHI', className: '7TECHI', termPair: 'Terms 1 & 2' });
  assert.equal(sync.success, true);
  assert.equal(objectRows(ss.getSheetByName('ClassTracking')).length, 1);

  // Re-syncing is idempotent (updates row, does not duplicate)
  const sync2 = W.syncClassTable(ss.id, 'Rube Goldberg', { year: '2026', stage: 'Stage 5', courseCode: '7TECHI', className: '7TECHI', termPair: 'Terms 1 & 2' });
  assert.equal(sync2.success, true);
  assert.equal(objectRows(ss.getSheetByName('ClassTracking')).length, 1);
});

