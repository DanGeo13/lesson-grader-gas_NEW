const {test}=require('node:test');
const assert=require('node:assert/strict');
const {context}=require('./helpers.cjs');

test('web access fails closed; Gemini key stored only in Script Properties, never returned',()=>{
  const values={};let email='teacher@school.edu.au';
  const props={getProperty:key=>values[key]||null,setProperty:(key,v)=>{values[key]=v;}};
  const g=context({
    Session:{getActiveUser:()=>({getEmail:()=>email})},
    PropertiesService:{getScriptProperties:()=>props},
    GeminiService:{testGeminiConnection:()=>({success:true,message:'Connected to Gemini successfully.'})}
  }).load('AppConfig.gs','WebAccess.gs');
  assert.equal(g.apiWebAuthStatus().data.authorised,false);
  assert.throws(()=>g.WebAccess.requireTeacher(),/Access denied/);
  values.LESSON_GRADER_ALLOWED_EMAILS='teacher@school.edu.au';
  assert.equal(g.apiWebAuthStatus().data.authorised,true);
  const secret='THIS_IS_A_TEST_KEY_NOT_REAL_123456';
  const result=g.apiWebSetGeminiKey(secret);
  assert.equal(result.success,true);
  assert.equal(values.GEMINI_API_KEY,secret);
  assert.ok(!JSON.stringify(result).includes(secret));
  assert.ok(!JSON.stringify(g.apiWebAuthStatus()).includes(secret));
  assert.equal(g.apiWebTestGeminiKey().success,true);
  email='outsider@school.edu.au';
  assert.throws(()=>g.apiWebSetGeminiKey(secret),/Access denied/);
  email='';assert.equal(g.apiWebAuthStatus().data.authorised,false);
});

test('Gemini transport uses x-goog-api-key header and no key in URL or error response',()=>{
  const key='THIS_IS_A_TEST_KEY_NOT_REAL_123456';
  const calls=[];
  const g=context({
    PropertiesService:{getScriptProperties:()=>({getProperty:()=>key})},
    UrlFetchApp:{fetch:(url,options)=>{
      calls.push({url,options});
      return {getResponseCode:()=>200,getContentText:()=>JSON.stringify({candidates:[{content:{parts:[{text:'{}'}]}}]})};
    }}
  }).load('Code.gs');
  const res=g.GeminiService.callGeminiWithFallback([{role:'user',parts:[{text:'Ping'}]}],false);
  assert.equal(res.success,true);
  assert.equal(calls.length,1);
  assert.equal(calls[0].options.headers['x-goog-api-key'],key);
  assert.ok(!calls[0].url.includes(key));
  assert.ok(!calls[0].url.includes('?key='));
  g.UrlFetchApp.fetch=()=>({getResponseCode:()=>403,getContentText:()=>`Forbidden: ${key}`});
  const failed=g.GeminiService.callGeminiWithFallback([{role:'user',parts:[{text:'Ping'}]}],false);
  assert.equal(failed.success,false);
  assert.ok(!failed.message.includes(key));
});

test('legacy and new public RPCs reject unauthorised callers before Google services run',()=>{
  let usedClassroom=false;
  const mocks={
    WebAccess:{requireTeacher:()=>{throw new Error('Access denied.');}},
    ClassroomService:{listTeacherCourses:()=>{usedClassroom=true;return {success:true};}},
    LessonGraderWeb:{list:()=>{usedClassroom=true;return {success:true};}}
  };
  const g=context(mocks).load('Code.gs','WebAppIntegration.gs','WebGrading.gs');
  assert.throws(()=>g.apiListTeacherCourses(),/Access denied/);
  assert.throws(()=>g.apiWebListCandidateWorkbooks(),/Access denied/);
  assert.throws(()=>g.apiWebEnableGrading('WORKBOOK'),/Access denied/);
  assert.equal(usedClassroom,false);
});

function settingsHarness(initial={}) {
  const values={LESSON_GRADER_ALLOWED_EMAILS:'teacher@school.edu.au',...initial};
  const props={getProperty:key=>values[key]||null,setProperty:(key,v)=>{values[key]=v;},deleteProperty:key=>{delete values[key];}};
  const g=context({
    Session:{getActiveUser:()=>({getEmail:()=>'teacher@school.edu.au'})},
    PropertiesService:{getScriptProperties:()=>props}
  }).load('AppConfig.gs','WebAccess.gs');
  return {g,values};
}

test('settings never store or return API keys, even from legacy JSON (F01)',()=>{
  const secret='THIS_IS_A_TEST_KEY_NOT_REAL_123456';
  // Legacy settings JSON that once captured a key through "Save All Settings".
  const {g,values}=settingsHarness({LESSON_GRADER_SETTINGS:JSON.stringify({geminiApiKey:secret,aiProviders:{geminiModel:'gemini-2.5-flash',apiKey:secret},feedbackTemplate:'Hi'})});
  const read=g.apiWebGetSettings();
  assert.ok(!JSON.stringify(read).includes(secret));
  assert.equal(read.data.settings.feedbackTemplate,'Hi');
  const rejected=g.apiWebSaveSettings({geminiApiKey:secret});
  assert.equal(rejected.success,false);
  assert.ok(!String(values.LESSON_GRADER_SETTINGS).includes('"geminiApiKey"') || !JSON.stringify(g.apiWebGetSettings()).includes(secret));
  const nested=g.apiWebSaveSettings({aiProviders:{draftProvider:'gemini',apiKey:secret}});
  assert.equal(nested.success,false);
  const ok=g.apiWebSaveSettings({outputFormat:{gradeDisplay:'letter',useFirstName:false}});
  assert.equal(ok.success,true);
  assert.ok(!values.LESSON_GRADER_SETTINGS.includes(secret), 'a successful save rewrites settings without the legacy key');
  assert.equal(g.apiWebGetSettings().data.settings.outputFormat.gradeDisplay,'letter');
});

test('settings reject unknown keys and grade-policy keys that do not affect grading (F31)',()=>{
  const {g}=settingsHarness();
  assert.match(g.apiWebSaveSettings({thresholds:[]}).message,/Setup sheet/);
  assert.match(g.apiWebSaveSettings({registeredApps:[]}).message,/Unknown setting/);
});

test('non-Gemini provider tests never report a mock success',()=>{
  const {g}=settingsHarness();
  const res=g.apiWebTestAiProvider('qwen');
  assert.equal(res.success,false);
  assert.doesNotMatch(res.message,/mock/i);
});

test('provider endpoints must be https and keys are never echoed',()=>{
  const {g,values}=settingsHarness();
  assert.equal(g.apiWebSetAiProviderConfig('qwen',{endpoint:'http://insecure.local/v1'}).success,false);
  const saved=g.apiWebSetAiProviderConfig('qwen',{endpoint:'https://ai.example.org/v1/chat/completions',key:'LOCAL_TEST_KEY_123'});
  assert.equal(saved.success,true);
  assert.equal(values.QWEN_ENDPOINT_URL,'https://ai.example.org/v1/chat/completions');
  assert.ok(!JSON.stringify(saved).includes('LOCAL_TEST_KEY_123'));
  assert.equal(g.apiWebGetSettings().data.providers.qwen.host,'ai.example.org');
});

test('unlock rights fail closed unless the teacher is on the unlock list',()=>{
  const {g,values}=settingsHarness();
  assert.equal(g.WebAccess.canUnlock('teacher@school.edu.au'),false);
  values.LESSON_GRADER_UNLOCK_EMAILS='teacher@school.edu.au';
  assert.equal(g.WebAccess.canUnlock('teacher@school.edu.au'),true);
  assert.equal(g.WebAccess.canUnlock('other@school.edu.au'),false);
});
