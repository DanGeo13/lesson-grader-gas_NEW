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
  const props={};
  const driveFiles={};
  const mocks={
    SpreadsheetApp:{flush(){}},
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    WebAccess:{requireTeacher:()=> 'dan@school.edu.au',canUnlock:(e)=>mocks.unlockers.includes(e),providerStatus:()=>({})},
    unlockers:[],
    LessonGraderWeb:{open:(wb)=> {assert.equal(wb,ss.id);return ss;},taskNames:()=>['Rube Goldberg','Water Filter'],isSharedRubric:()=>true},
    DriveService:{isEligibleForAi:(mime)=>mime === 'application/pdf'},
    // Drive fake: registered files are readable; anything else behaves like a backup source.
    DriveApp:{getFileById:(fid)=>driveFiles[fid] || ({makeCopy:()=>({getId:()=> 'backup-1'})})},
    driveFiles,
    MimeType:{GOOGLE_DOCS:'application/vnd.google-apps.document'},
    PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>{props[k]=v;},deleteProperty:k=>{delete props[k];}})},
    props,
    Utilities:{getUuid:()=>`id-${id++}`,formatDate:(d)=>d.toISOString().slice(0,10),base64Encode:()=> 'dGVzdA=='}
  };
  const g=context(mocks).load('AppConfig.gs','RubricEngine.gs','AssessmentFlags.gs','Deidentify.gs','EvidenceService.gs','PromptRegistry.gs','AiGateway.gs','AiQueue.gs','WebGrading.gs','EvaluationRegister.gs');
  return {g,ss};
}
/* Registers a readable Drive file and its SubmissionFiles row. */
function addFile(g,ss,subId,fileId,name,mime,content){
  g.driveFiles[fileId]={getMimeType:()=>mime,getSize:()=>100,getBlob:()=>({getBytes:()=>[1,2,3],getContentType:()=>mime}),
    getAs:()=>({getBytes:()=>[1],getDataAsString:()=>content||''})};
  const sh=ss.getSheetByName('SubmissionFiles');
  sh.appendRow(sh.rows[0].map(key=>({SubmissionRecordID:subId,FileRecordID:'FR-'+fileId,DriveFileID:fileId,FileName:name,MimeType:mime,AIReviewStatus:'Pending'})[key]??''));
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
  },{whatWentWell:'Good engineering',areasForImprovement:'Measure more',goalsForNextAssessment:'Test early',markFeedbackCurrent:true},'', 'Teacher');
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
  assert.equal(W.detail(ss.id,'Rube Goldberg','SUB-1').data.submission.stateLabel,'Final (locked)');
  W.activateRubric(ss.id,'Rube Goldberg','Newer rubric','',[...rowsOne(),{criterion:'Presentation',maxMarks:5,band:'A',description:'Clear communication'}],{},true);
  assert.equal(W.detail(ss.id,'Rube Goldberg','SUB-1').data.profile.profileId,oldId);
  // Final lock: no reassessment and no unlock without authorisation (SC41).
  assert.throws(()=>W.reassess(ss.id,'Rube Goldberg','SUB-1'),/final-locked/);
  assert.throws(()=>W.unlock(ss.id,'Rube Goldberg','SUB-1','New evidence from class','UNLOCK'),/not authorised/);
  g.WebAccess.canUnlock=()=>true;
  assert.equal(W.detail(ss.id,'Rube Goldberg','SUB-1').data.canUnlock,true);
  assert.throws(()=>W.unlock(ss.id,'Rube Goldberg','SUB-1','short','UNLOCK'),/reason/);
  assert.throws(()=>W.unlock(ss.id,'Rube Goldberg','SUB-1','New evidence from class','yes'),/Type UNLOCK/);
  W.unlock(ss.id,'Rube Goldberg','SUB-1','New evidence from class','UNLOCK');
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].Status,'Approved');
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].UnlockReason,'New evidence from class');
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades')).length,1,'unlock never alters the approval snapshot');
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

test('AI drafts only evidence-cited ticks; uncited and unknown ticks are discarded and flagged (SC17, SC22)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;
  W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  addFile(g,ss,'SUB-1','FILE-1','folio.pdf','application/pdf');
  let prompt='';
  g.GeminiService={callGeminiWithFallback:(contents)=>{
    prompt=JSON.stringify(contents);
    return {success:true,modelUsed:'gemini-2.5-flash',text:'```json\n'+JSON.stringify({totalMarks:99,grade:'A+',criteria:[
      {criterionId:'C01',tickedCheckboxes:['C01-A-1','C01-B-1','C01-MD','C01-FAKE'],score:10,evidenceNotes:[{checkboxId:'C01-A-1',fileName:'folio.pdf',note:'Page 3 research table'}]},
      {criterionId:'C02',tickedCheckboxes:['C02-A-1'],evidenceNotes:[{checkboxId:'C02-A-1',fileName:'someone-else.pdf',note:'not supplied'}]}
    ]})+'\n```'};
  }};
  const result=W.runAi(ss.id,'Rube Goldberg','SUB-1','');
  assert.equal(result.success,true);
  assert.equal(result.applied,true);
  assert.ok(prompt.includes('Research'));
  assert.ok(prompt.includes('untrusted evidence'));
  const d=W.detail(ss.id,'Rube Goldberg','SUB-1').data;
  assert.equal(d.submission.status,'InReview');
  assert.equal(d.criteriaMap.C01.checkboxes['C01-A-1'],true);
  assert.equal(d.criteriaMap.C01.checkboxes['C01-B-1'],undefined,'tick without a cited file is discarded');
  assert.equal(d.criteriaMap.C01.checkboxes['C01-MD'],undefined);
  assert.equal(d.criteriaMap.C01.checkboxes['C01-FAKE'],undefined);
  assert.equal(d.criteriaMap.C01.aiProposal.unsupported.join(),'C01-B-1');
  assert.equal(d.criteriaMap.C01.aiEvidence[0].note,'Page 3 research table');
  assert.equal(d.criteriaMap.C02.aiProposal.insufficient,true,'citing a file that was not supplied is not evidence');
  // AI totals/grades in the reply are ignored: marks come from the engine (C01 A-1 = 10 x 1.0 / 1).
  assert.equal(d.scores.totalPoints,10);
  const codes=d.flags.flags.map(f=>f.code);
  assert.ok(codes.includes('AI_INSUFFICIENT_EVIDENCE'));
  assert.ok(codes.includes('AI_UNSUPPORTED_TICKS'));
  assert.equal(objectRows(ss.getSheetByName('SubmissionFiles'))[0].AIReviewStatus,'Reviewed');
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades')).length,0);
  assert.ok(objectRows(ss.getSheetByName('AIAssessments')).some(row=>row.RunType==='EvidenceProposal' && row.ExecutionStatus==='Success'));
  assert.throws(()=>W.approve(ss.id,'Rube Goldberg','SUB-1',d.revision),/Resolve before approving/);
});

test('AI suggestions never overwrite selections the teacher already saved (SC26)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  addFile(g,ss,'SUB-1','FILE-1','folio.pdf','application/pdf');
  const saved=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{C01:{checkboxes:{'C01-B-1':true}},C02:{checkboxes:{}}},{},'','Teacher');
  g.GeminiService={callGeminiWithFallback:()=>({success:true,modelUsed:'m',text:JSON.stringify({criteria:[
    {criterionId:'C01',tickedCheckboxes:['C01-A-1'],evidenceNotes:[{checkboxId:'C01-A-1',fileName:'folio.pdf',note:'p1'}]}]})})};
  const res=W.runAi(ss.id,'Rube Goldberg','SUB-1',saved.revision);
  assert.equal(res.applied,false);
  const d=W.detail(ss.id,'Rube Goldberg','SUB-1').data;
  assert.equal(d.criteriaMap.C01.checkboxes['C01-B-1'],true);
  assert.equal(d.criteriaMap.C01.checkboxes['C01-A-1'],undefined);
  assert.equal(d.criteriaMap.C01.aiProposal.checks['C01-A-1'],true,'kept as a suggestion');
});

test('AI outage returns a choice instead of an error, and manual marking still works (SC48, SC50)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  addFile(g,ss,'SUB-1','FILE-1','folio.pdf','application/pdf');
  g.GeminiService={callGeminiWithFallback:()=>({success:false,unavailable:true,message:'Model busy (code 503).'})};
  const res=W.runAi(ss.id,'Rube Goldberg','SUB-1','');
  assert.equal(res.success,false);
  assert.equal(res.unavailable,true);
  assert.equal(objectRows(ss.getSheetByName('CriterionAssessments')).length,0,'nothing drafted');
  // Manual path is unaffected by the outage.
  const draft=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{C01:{checkboxes:{'C01-A-1':true}},C02:{checkboxes:{'C02-A-1':true}}},{whatWentWell:'Clear research',markFeedbackCurrent:true},'','Teacher');
  assert.equal(W.approve(ss.id,'Rube Goldberg','SUB-1',draft.revision).success,true);
});

test('unreadable or unsupported files are disclosed and block approval until acknowledged (SC23)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  addFile(g,ss,'SUB-1','FILE-1','folio.pdf','application/pdf');
  addFile(g,ss,'SUB-1','FILE-2','model.stl','model/stl');
  const sh=ss.getSheetByName('SubmissionFiles');sh.appendRow(sh.rows[0].map(k=>({SubmissionRecordID:'SUB-1',FileRecordID:'FR-X',DriveFileID:'GONE',FileName:'locked.docx',MimeType:'application/pdf'})[k]??''));
  g.DriveApp.getFileById=(fid)=>{ if(fid==='GONE') throw new Error('denied'); return g.driveFiles[fid]; };
  g.GeminiService={callGeminiWithFallback:()=>({success:true,modelUsed:'m',text:JSON.stringify({criteria:[
    {criterionId:'C01',tickedCheckboxes:['C01-A-1'],evidenceNotes:[{checkboxId:'C01-A-1',fileName:'folio.pdf',note:'p1'}]},
    {criterionId:'C02',tickedCheckboxes:['C02-A-1'],evidenceNotes:[{checkboxId:'C02-A-1',fileName:'folio.pdf',note:'p2'}]}]})})};
  const res=W.runAi(ss.id,'Rube Goldberg','SUB-1','');
  assert.equal(res.files.find(f=>f.fileName==='model.stl').status,'Unsupported');
  assert.equal(res.files.find(f=>f.fileName==='locked.docx').status,'Unreadable');
  let d=W.detail(ss.id,'Rube Goldberg','SUB-1').data;
  assert.equal(d.attachments.find(a=>a.fileName==='locked.docx').status,'Unreadable');
  assert.ok(d.flags.flags.some(f=>f.code==='UNREADABLE_EVIDENCE' && f.severity==='blocking'));
  const s2=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{C01:{checkboxes:{'C01-A-1':true}},C02:{checkboxes:{'C02-A-1':true}}},{acknowledgedUnreadable:true},d.revision,'Teacher');
  assert.equal(s2.flags.flags.some(f=>f.code==='UNREADABLE_EVIDENCE'),false);
  assert.equal(W.approve(ss.id,'Rube Goldberg','SUB-1',s2.revision).success,true);
});

test('no creditable evidence can be confirmed explicitly instead of ticking false evidence (F12)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const d1=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{C01:{checkboxes:{'C01-A-1':true}},C02:{checkboxes:{}}},{},'','Teacher');
  assert.ok(d1.flags.flags.some(f=>f.code==='NO_SELECTION' && f.criterionId==='C02'));
  assert.throws(()=>W.approve(ss.id,'Rube Goldberg','SUB-1',d1.revision),/no creditable evidence/);
  const d2=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{C01:{checkboxes:{'C01-A-1':true}},C02:{checkboxes:{},review:{noEvidence:true}}},{},d1.revision,'Teacher');
  const a=W.approve(ss.id,'Rube Goldberg','SUB-1',d2.revision);
  assert.equal(a.scores.totalPoints,10);
});

test('feedback must be updated or confirmed after selections change (SC32, SC38)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const both={C01:{checkboxes:{'C01-A-1':true}},C02:{checkboxes:{'C02-A-1':true}}};
  const d1=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',both,{whatWentWell:'Excellent research',markFeedbackCurrent:true},'','Teacher');
  assert.equal(d1.flags.flags.some(f=>f.code==='FEEDBACK_NOT_REVIEWED'),false);
  const changed={C01:{checkboxes:{'C01-B-1':true}},C02:{checkboxes:{'C02-A-1':true}}};
  const d2=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',changed,{whatWentWell:'Excellent research'},d1.revision,'Teacher');
  assert.ok(d2.flags.flags.some(f=>f.code==='FEEDBACK_NOT_REVIEWED' && f.severity==='blocking'));
  assert.throws(()=>W.approve(ss.id,'Rube Goldberg','SUB-1',d2.revision),/Feedback was written before/);
  const d3=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',changed,{whatWentWell:'Sound research',markFeedbackCurrent:true},d2.revision,'Teacher');
  assert.equal(W.approve(ss.id,'Rube Goldberg','SUB-1',d3.revision).success,true);
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
  },{whatWentWell:'Good',areasForImprovement:'None',goalsForNextAssessment:'Keep it up',markFeedbackCurrent:true},'', 'Teacher', {
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

test('class batch drafts through the queue: per-student isolation, dedupe, readiness (SC28, SC49)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const sub2 = W.addManualSubmission(ss.id, 'Rube Goldberg', 'Sam Green', 'sam@school.edu.au', '7TECHI');
  addFile(g,ss,'SUB-1','FILE-1','folio.pdf','application/pdf');
  let calls=0;
  g.GeminiService={callGeminiWithFallback:()=>{calls++;return {success:true,modelUsed:'gemini-2.5-flash',text:JSON.stringify({criteria:[
    {criterionId:'C01',tickedCheckboxes:['C01-A-1'],evidenceNotes:[{checkboxId:'C01-A-1',fileName:'folio.pdf',note:'p1'}]},
    {criterionId:'C02',tickedCheckboxes:['C02-A-1'],evidenceNotes:[{checkboxId:'C02-A-1',fileName:'folio.pdf',note:'p2'}]}]})};}};
  const batch = W.bulkGrade(ss.id, 'Rube Goldberg', ['SUB-1', sub2.submissionRecordId]);
  assert.equal(batch.queued, 2);
  assert.equal(batch.counts.Done, 1);      // SUB-1 drafted
  assert.equal(batch.counts.Failed, 1);    // Sam has no files: isolated failure, not a crash
  assert.equal(batch.remaining, 0);
  assert.equal(calls, 1);
  const st=g.AiQueue.status(ss.id,'Rube Goldberg').data;
  assert.equal(st.counts.Done,1); assert.equal(st.counts.Failed,1);
  // Re-queue while a job is active is a no-op; approved work is never queued.
  assert.equal(g.AiQueue.enqueue(ss.id,'Rube Goldberg',['SUB-1']).queued,1);
  assert.equal(g.AiQueue.enqueue(ss.id,'Rube Goldberg',['SUB-1']).alreadyQueued,1);
});

test('queue waits through an outage, resumes, and skips work the teacher edited meanwhile (SC48, SC49)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const sub2 = W.addManualSubmission(ss.id, 'Rube Goldberg', 'Sam Green', 'sam@school.edu.au', '7TECHI');
  addFile(g,ss,'SUB-1','FILE-1','folio.pdf','application/pdf');
  addFile(g,ss,sub2.submissionRecordId,'FILE-2','sam.pdf','application/pdf');
  let up=false, calls=0;
  g.GeminiService={callGeminiWithFallback:(contents)=>{calls++; if(!up) return {success:false,unavailable:true,message:'503'};
    const name=JSON.stringify(contents).includes('sam.pdf')?'sam.pdf':'folio.pdf';
    return {success:true,modelUsed:'m',text:JSON.stringify({criteria:[{criterionId:'C01',tickedCheckboxes:['C01-A-1'],evidenceNotes:[{checkboxId:'C01-A-1',fileName:name,note:'p1'}]}]})};}};
  g.AiQueue.enqueue(ss.id,'Rube Goldberg',['SUB-1',sub2.submissionRecordId]);
  const down=g.AiQueue.process(ss.id,'Rube Goldberg',60000);
  assert.equal(down.counts.Waiting,1,'stops after the first outage instead of hammering the service');
  assert.equal(down.remaining,2);
  // Teacher marks Sam manually while the service is down.
  const rev=W.detail(ss.id,'Rube Goldberg',sub2.submissionRecordId).data.revision;
  W.saveDraft(ss.id,'Rube Goldberg',sub2.submissionRecordId,{C01:{checkboxes:{'C01-B-1':true}},C02:{checkboxes:{}}},{},rev,'Teacher');
  up=true;
  const resumed=g.AiQueue.process(ss.id,'Rube Goldberg',60000);
  assert.equal(resumed.counts.Done,1);
  assert.equal(resumed.counts.Skipped,1);
  assert.equal(W.detail(ss.id,'Rube Goldberg',sub2.submissionRecordId).data.criteriaMap.C01.checkboxes['C01-B-1'],true,'teacher work kept');
  const jobs=objectRows(ss.getSheetByName('AiQueue'));
  assert.equal(jobs.filter(j=>j.Status==='Done').length,1);
  assert.equal(jobs.filter(j=>j.Status==='Skipped').length,1);
  assert.equal(calls,2,'no duplicate processing: one failed attempt, one success');
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
  }, { whatWentWell: 'Good', markFeedbackCurrent: true }, '', 'Teacher');
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
  assert.equal(objectRows(ss.getSheetByName('AuditLog')).filter(r=>r.Action.startsWith('SampleSaved_') && r.Status==='Saved').length, 3);

  // Without destination folders nothing is copied, and the result says so (F04).
  const none = W.saveWorkSamples(ss.id, 'Rube Goldberg', sugg.proposals, {}, true);
  assert.equal(none.success, false);
  assert.match(none.message, /Copied 0 of 3/);
  g.DriveApp.getFileById = () => ({ makeCopy: () => { throw new Error('Permission denied'); } });
  const failed = W.saveWorkSamples(ss.id, 'Rube Goldberg', sugg.proposals, { highFolderId: 'H', medFolderId: 'M', lowFolderId: 'L' }, true);
  assert.equal(failed.success, false);
  assert.ok(failed.results.every(r => !r.copied && /Copy failed/.test(r.message)));
});

test('register entry pre-fills only workbook facts and commits to the workbook sheet (F03)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const d1 = W.saveDraft(ss.id, 'Rube Goldberg', 'SUB-1', { C01:{checkboxes:{'C01-A-1':true}}, C02:{checkboxes:{'C02-A-1':true}} }, {}, '', 'Teacher');
  W.approve(ss.id, 'Rube Goldberg', 'SUB-1', d1.revision);

  const preview = W.previewRegistrationPrefill(ss.id, 'Rube Goldberg', '');
  assert.equal(preview.success, true);
  assert.equal(preview.proposed.unitName, 'Rube Goldberg');
  assert.ok(preview.proposed.gradeDistribution.includes('A: 1'));
  assert.equal(preview.proposed.teacherVariations, '', 'no invented evaluation text');
  assert.equal(preview.proposed.courseCode, '7TECHI');
  assert.equal(preview.destination.type, 'sheet');

  assert.throws(()=>W.saveRegistrationPrefill(ss.id, 'Rube Goldberg', '', {}, true), /teacher evaluation/);
  const saved = W.saveRegistrationPrefill(ss.id, 'Rube Goldberg', '', { teacherEvaluation: 'Students iterated well.', studentFeedbackSummary: 'Time management' }, true);
  assert.equal(saved.success, true);
  const rows = objectRows(ss.getSheetByName('EvaluationRegister'));
  assert.equal(rows.length, 1);
  assert.equal(rows[0].TeacherEvaluation, 'Students iterated well.');
  assert.equal(rows[0].Status, 'Recorded in workbook');
  assert.equal(W.getAuditDashboard(ss.id).tasks.find(t=>t.taskName==='Rube Goldberg').registrationDocStatus, 'Complete');
});

test('register Doc write failure is reported, never claimed as success',()=>{
  const {g,ss}=setup();g.WebGrading.prepare(ss,false);
  g.DriveApp={getFileById:()=>({isTrashed:()=>false,getMimeType:()=>'application/vnd.google-apps.document',getName:()=>'Register 2026',getUrl:()=>'https://docs.google.com/document/d/x'})};
  g.DocumentApp={openById:()=>{throw new Error('No edit access');},ParagraphHeading:{HEADING2:'H2'}};
  const res=g.EvaluationRegister.commit(ss.id,'Rube Goldberg',{teacherEvaluation:'Reflection'},'DOC'+'1'.repeat(30),true);
  assert.equal(res.success,false);
  assert.match(res.message,/NOT updated/);
  assert.equal(objectRows(ss.getSheetByName('EvaluationRegister'))[0].Status,'Doc write failed — recorded in workbook only');
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

  // Confirming records drafts honestly: nothing is sent (F02).
  const sent = W.sendClassroomMessages(ss.id, 'Rube Goldberg', drafts.drafts, true);
  assert.equal(sent.success, true);
  assert.equal(sent.sent, 0);
  assert.match(sent.message, /Nothing was sent/);
  assert.doesNotMatch(sent.message, /Successfully sent/);
  const rows = objectRows(ss.getSheetByName('ClassroomMessages'));
  assert.equal(rows.length, 1);
  assert.equal(rows[0].Status, 'Drafted — not sent');
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


test('R10 reassessment keeps the earlier approval and reports what changed (SC42)',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  const v1=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{C01:{checkboxes:{'C01-B-1':true}},C02:{checkboxes:{'C02-A-1':true}}},{},'','Teacher');
  W.approve(ss.id,'Rube Goldberg','SUB-1',v1.revision);
  const snapshotBefore=JSON.stringify(objectRows(ss.getSheetByName('ApprovedGrades'))[0]);
  const r=W.reassess(ss.id,'Rube Goldberg','SUB-1');
  const v2id=r.submissionRecordId;
  let d=W.detail(ss.id,'Rube Goldberg',v2id).data;
  assert.equal(d.submission.stateLabel,'Not started · Reassessment v2');
  const v2=W.saveDraft(ss.id,'Rube Goldberg',v2id,{C01:{checkboxes:{'C01-A-1':true}},C02:{checkboxes:{'C02-A-1':true}}},{},d.revision,'Teacher');
  d=W.detail(ss.id,'Rube Goldberg',v2id).data;
  const c01=d.comparison.criteria.find(c=>c.criterionId==='C01');
  assert.equal(c01.added.join(),'C01-A-1');
  assert.equal(c01.removed.join(),'C01-B-1');
  assert.equal(d.comparison.before.totalPoints,23.75);   // 8.75 + 15
  assert.equal(d.comparison.after.totalPoints,25);
  assert.equal(d.comparison.approved.grade,'A');
  W.approve(ss.id,'Rube Goldberg',v2id,v2.revision);
  const grades=objectRows(ss.getSheetByName('ApprovedGrades'));
  assert.equal(grades.length,2);
  assert.equal(JSON.stringify(grades[0]),snapshotBefore,'v1 snapshot untouched');
  const v1row=objectRows(ss.getSheetByName('Submissions')).find(s=>s.SubmissionRecordID==='SUB-1');
  assert.equal(v1row.Status,'Approved'); assert.equal(v1row.CurrentOfficial,false);
});
