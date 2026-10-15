/* Shared in-memory Apps Script environment for workflow tests. */
const assert = require('node:assert/strict');
const { context, Workbook } = require('./helpers.cjs');

const MODULES = ['AppConfig.gs','RubricEngine.gs','AssessmentFlags.gs','Deidentify.gs','EvidenceService.gs','PromptRegistry.gs',
  'AiGateway.gs','AiQueue.gs','FeedbackBuilder.gs','Publishing.gs','Refinement.gs','WebGrading.gs','EvaluationRegister.gs'];

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
    deriveStageFromCourseCode:(c)=>/^(9|10)/.test(c)?'Stage 5':/^(7|8)/.test(c)?'Stage 4':null,
    DriveService:{isEligibleForAi:(mime)=>mime === 'application/pdf'},
    // Drive fake: registered files are readable; anything else behaves like a backup source.
    DriveApp:{getFileById:(fid)=>driveFiles[fid] || ({makeCopy:()=>({getId:()=> 'backup-1'})})},
    driveFiles,
    MimeType:{GOOGLE_DOCS:'application/vnd.google-apps.document'},
    PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>{props[k]=v;},deleteProperty:k=>{delete props[k];}})},
    props,
    Utilities:{getUuid:()=>`id-${id++}`,formatDate:(d)=>d.toISOString().slice(0,10),base64Encode:()=> 'dGVzdA=='}
  };
  const g=context(mocks).load(...MODULES);
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

function classroomFake(g,payload,profiles){
  g.Classroom={Courses:{CourseWork:{StudentSubmissions:{list:(_,work,params)=>{
    const all=payload[work]||[];
    // two pages to prove pagination is followed
    if(!params.pageToken) return {studentSubmissions:all.slice(0,1),nextPageToken:all.length>1?'p2':undefined};
    return {studentSubmissions:all.slice(1)};
  }}}},UserProfiles:{get:(id)=>profiles[id]}};
}
function linkTask(ss,task,work,due){
  ss.getSheetByName('ClassroomConfig').appendRow(['cfg-'+work,'COURSE-1','7TECHI','',work,task,100,'','',true,task,due||'',false]);
}


module.exports = { setup, addFile, rowsOne, classroomFake, linkTask };
