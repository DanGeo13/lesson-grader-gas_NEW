const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context, Workbook, objectRows } = require('./helpers.cjs');

function env(extraProps = {}) {
  const ss = new Workbook();
  ss.add('Bridge Build', [['Task']]);
  ss.add('Submissions', [
    ['SubmissionRecordID', 'StudentUserID', 'Status', 'Task', 'StudentName', 'StudentEmail', 'SubmissionVersion', 'CurrentOfficial'],
    ['SUB-1', '118822334455', 'New', 'Bridge Build', 'Mia Taylor', 'mia.taylor@school.edu.au', 1, true],
    ['SUB-2', '118822334466', 'New', 'Bridge Build', 'Noah Wilson', 'noah.wilson@school.edu.au', 1, true]
  ]);
  ss.add('ClassroomConfig', [['ConfigID', 'CourseID', 'CourseName', 'CourseSection', 'CourseWorkID', 'AssignmentTitle', 'MaxPoints', 'SavedAt', 'SavedBy', 'Active', 'TaskName', 'DueDate', 'AutoImported']]);
  let id = 1;
  const props = { ...extraProps };
  const files = {};
  const fetches = [];
  const docText = { DOC1: "# Mia Taylor's bridge folio\nContact mia.taylor@school.edu.au or 0412 345 678. Student 4412345.\nMia tested three truss designs; Noah helped." };
  const mocks = {
    SpreadsheetApp: { flush() {} },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    WebAccess: { requireTeacher: () => 'dan@school.edu.au', canUnlock: () => false, providerStatus: () => ({}) },
    LessonGraderWeb: { open: () => ss, taskNames: () => ['Bridge Build'], isSharedRubric: () => true },
    DriveService: { isEligibleForAi: () => true },
    DriveApp: { getFileById: fid => files[fid] },
    MimeType: { GOOGLE_DOCS: 'application/vnd.google-apps.document' },
    ScriptApp: { getOAuthToken: () => 'oauth-token' },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; } }) },
    Utilities: { getUuid: () => `id-${id++}`, base64Encode: () => 'aW1n', formatDate: d => String(d) },
    UrlFetchApp: {
      fetch: (url, opts) => {
        fetches.push({ url, opts });
        if (url.includes('/drive/v3/files/')) { const fid = url.split('/files/')[1].split('/')[0]; return { getResponseCode: () => 200, getContentText: () => docText[fid] }; }
        if (url === props.JEV_ENDPOINT_URL) {
          return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ model: 'jev-1', choices: [{ message: { content: JSON.stringify({ criteria: [
            { criterionId: 'C01', tickedCheckboxes: ['C01-A-1'], evidenceNotes: [{ checkboxId: 'C01-A-1', fileName: 'File 1', note: 'Tested three truss designs' }] }
          ] }) } }] }) };
        }
        return { getResponseCode: () => 404, getContentText: () => '' };
      }
    }
  };
  const g = context(mocks).load('AppConfig.gs', 'RubricEngine.gs', 'AssessmentFlags.gs', 'Deidentify.gs', 'EvidenceService.gs', 'PromptRegistry.gs', 'AiGateway.gs', 'AiQueue.gs', 'Refinement.gs', 'WebGrading.gs');
  g.WebGrading.prepare(ss, false);
  g.WebGrading.activateRubric(ss.id, 'Bridge Build', 'Bridge rubric', '', [
    { criterion: 'Testing', maxMarks: 10, band: 'A', description: 'Tests several designs' },
    { criterion: 'Testing', maxMarks: 10, band: 'C', description: 'Tests one design' }
  ], {}, true);
  function addDoc(subId, fid, name) {
    files[fid] = { getMimeType: () => 'application/vnd.google-apps.document', getSize: () => 10, getAs: () => ({ getDataAsString: () => docText[fid] }) };
    const sh = ss.getSheetByName('SubmissionFiles');
    sh.appendRow(sh.rows[0].map(k => ({ SubmissionRecordID: subId, FileRecordID: 'FR-' + fid, DriveFileID: fid, FileName: name, MimeType: 'application/vnd.google-apps.document', AIReviewStatus: 'Pending' })[k] ?? ''));
  }
  return { g, ss, props, fetches, docText, addDoc, files };
}

test('de-identification removes names, emails, phones, IDs and links; reports what it removed', () => {
  const { g } = env();
  const out = g.Deidentify.apply("Mia Taylor (mia.taylor@school.edu.au, 0412 345 678, ID 4412345) said Mia's design https://drive.google.com/x works. Taylor-made joints.",
    { names: ['Mia Taylor'], ids: ['4412345'], rosterNames: ['Noah Wilson'] });
  assert.ok(!/Mia|mia\.taylor|0412|4412345|drive\.google/.test(out.text), out.text);
  assert.ok(out.text.includes('[STUDENT]') && out.text.includes('[EMAIL]') && out.text.includes('[PHONE]') && out.text.includes('[LINK]'));
  assert.ok(out.counts.name >= 2);
  assert.throws(() => g.Deidentify.assertClean('Mia did well', { names: ['Mia Taylor'] }), /De-identification check failed/);
});

test('JEV receives only de-identified Markdown and the rubric — no names, emails, IDs or file names (SC43)', () => {
  const { g, ss, props, fetches, addDoc } = env({ JEV_ENDPOINT_URL: 'https://jev.example.org/v1/chat/completions', JEV_API_KEY: 'TEST_JEV_KEY_1234',
    LESSON_GRADER_SETTINGS: JSON.stringify({ aiProviders: { draftProvider: 'jev', visionProvider: 'none' } }) });
  addDoc('SUB-1', 'DOC1', 'Mia Taylor Upload 7Q');
  const res = g.WebGrading.runAi(ss.id, 'Bridge Build', 'SUB-1', '');
  assert.equal(res.success, true, res.message);
  const call = fetches.find(f => f.url === props.JEV_ENDPOINT_URL);
  const payload = call.opts.payload;
  for (const leak of ['Mia', 'Taylor', 'Noah', 'mia.taylor', '0412', '4412345', '118822334455', 'SUB-1', 'Upload 7Q']) {
    assert.ok(!payload.includes(leak), 'payload leaked ' + leak);
  }
  assert.ok(payload.includes('Tests several designs'), 'rubric content is included');
  assert.ok(payload.includes('File 1'));
  assert.equal(call.opts.headers.Authorization, 'Bearer TEST_JEV_KEY_1234');
  // Evidence note label maps back to the real file name for the teacher.
  const d = g.WebGrading.detail(ss.id, 'Bridge Build', 'SUB-1').data;
  assert.equal(d.criteriaMap.C01.aiEvidence[0].fileName, 'Mia Taylor Upload 7Q');
  const log = objectRows(ss.getSheetByName('AIAssessments')).find(r => r.RunType === 'EvidenceProposal');
  assert.match(log.PromptSummary, /de-identified/);
  assert.ok(!JSON.parse(log.ResponseJSON).payloadPreview.includes('Taylor'));
});

test('Markdown evidence history keeps every distinct version used for assessment (SC44)', () => {
  const { g, ss, addDoc, docText } = env({ LESSON_GRADER_SETTINGS: JSON.stringify({ aiProviders: { draftProvider: 'gemini' } }), GEMINI_API_KEY: 'k' });
  addDoc('SUB-1', 'DOC1', 'folio');
  g.GeminiService = { callGeminiWithFallback: () => ({ success: true, modelUsed: 'm', text: JSON.stringify({ criteria: [{ criterionId: 'C01', tickedCheckboxes: [], insufficientEvidence: true }] }) }) };
  const W = g.WebGrading;
  let rev = W.detail(ss.id, 'Bridge Build', 'SUB-1').data.revision;
  W.runAi(ss.id, 'Bridge Build', 'SUB-1', rev);
  rev = W.detail(ss.id, 'Bridge Build', 'SUB-1').data.revision;
  W.runAi(ss.id, 'Bridge Build', 'SUB-1', rev);            // same content: no duplicate version
  docText.DOC1 = docText.DOC1 + '\nAdded load-test results.';
  rev = W.detail(ss.id, 'Bridge Build', 'SUB-1').data.revision;
  W.runAi(ss.id, 'Bridge Build', 'SUB-1', rev);            // new content: new version
  const versions = objectRows(ss.getSheetByName('EvidenceVersions'));
  assert.equal(versions.length, 2);
  assert.ok(versions[1].Markdown.includes('load-test'));
  assert.notEqual(versions[0].ContentHash, versions[1].ContentHash);
  assert.equal(g.EvidenceService.history(ss, 'SUB-1').length, 2);
});

test('prompt versions must be tested on approved portfolios before activation; tests change nothing (SC46, SC47)', () => {
  const { g, ss, addDoc } = env({ GEMINI_API_KEY: 'k' });
  const W = g.WebGrading;
  addDoc('SUB-1', 'DOC1', 'folio');
  const d = W.saveDraft(ss.id, 'Bridge Build', 'SUB-1', { C01: { checkboxes: { 'C01-C-1': true } } }, {}, '', 'Teacher');
  W.approve(ss.id, 'Bridge Build', 'SUB-1', d.revision);
  const approvedBefore = JSON.stringify(objectRows(ss.getSheetByName('ApprovedGrades')));
  const ticksBefore = JSON.stringify(objectRows(ss.getSheetByName('CriterionAssessments')));
  assert.throws(() => g.PromptRegistry.create('too short'), /80/);
  const created = g.PromptRegistry.create('Assess the task {{task}} strictly against the rubric. Cite the file label and page for every tick. Reply in the JSON format: {"criteria":[...]}', 'stricter citations');
  const vid = created.data.id;
  assert.throws(() => g.PromptRegistry.activate(vid, 'ACTIVATE'), /Test this prompt/);
  g.GeminiService = { callGeminiWithFallback: () => ({ success: true, modelUsed: 'm', text: JSON.stringify({ criteria: [{ criterionId: 'C01', tickedCheckboxes: ['C01-A-1'], evidenceNotes: [{ checkboxId: 'C01-A-1', fileName: 'folio', note: 'three designs' }] }] }) }) };
  const t = g.PromptRegistry.test(vid, ss.id, 'Bridge Build', 3);
  assert.equal(t.success, true);
  assert.equal(t.data.completed, 1);
  assert.equal(t.data.potentialChanges.length, 1, 'reports that the approved C would become an A');
  assert.equal(JSON.stringify(objectRows(ss.getSheetByName('ApprovedGrades'))), approvedBefore, 'approved result unchanged');
  assert.equal(JSON.stringify(objectRows(ss.getSheetByName('CriterionAssessments'))), ticksBefore, 'no ticks changed');
  assert.equal(objectRows(ss.getSheetByName('PromptTests'))[0].PotentialChange, true);
  assert.throws(() => g.PromptRegistry.activate(vid, 'yes'), /ACTIVATE/);
  g.PromptRegistry.activate(vid, 'ACTIVATE');
  assert.equal(g.PromptRegistry.active().id, vid);
});

test('model JSON is parsed defensively and never trusted for marks', () => {
  const { g } = env();
  assert.equal(g.AiGateway.parseModelJson('not json'), null);
  assert.equal(g.AiGateway.parseModelJson('```json\n{"criteria":[]}\n```').criteria.length, 0);
  const profile = g.RubricEngine.buildProfileFromRubricRows([{ criterion: 'T', maxMarks: 10, band: 'A', description: 'x' }], { taskName: 'T', name: 'T' });
  assert.throws(() => g.AiGateway.validateProposal(profile, { score: 10 }, []), /expected format/);
  const v = g.AiGateway.validateProposal(profile, { criteria: [{ criterionId: 'C01', tickedCheckboxes: ['C01-A-1', '<script>'], marks: 99, evidenceNotes: [{ checkboxId: 'C01-A-1', fileName: 'F.PDF ', note: 'p1' }] }] }, ['f.pdf']);
  assert.equal(v.C01.checks['C01-A-1'], true);
  assert.equal(Object.keys(v.C01.checks).length, 1);
});

test('teacher corrections and pasted examples are stored de-identified (SC45)', () => {
  const { g, ss, addDoc } = env({ GEMINI_API_KEY: 'k' });
  const W = g.WebGrading;
  addDoc('SUB-1', 'DOC1', 'folio');
  g.GeminiService = { callGeminiWithFallback: () => ({ success: true, modelUsed: 'm', text: JSON.stringify({ criteria: [{ criterionId: 'C01', tickedCheckboxes: ['C01-A-1'], evidenceNotes: [{ checkboxId: 'C01-A-1', fileName: 'folio', note: 'p1' }] }] }) }) };
  W.runAi(ss.id, 'Bridge Build', 'SUB-1', '');
  let d = W.detail(ss.id, 'Bridge Build', 'SUB-1').data;
  const saved = W.saveDraft(ss.id, 'Bridge Build', 'SUB-1', { C01: { checkboxes: { 'C01-C-1': true }, teacherWrittenNote: 'Mia only tested one design in class.', review: { reviewed: true } } },
    { whatWentWell: 'Mia, your truss sketches were clear.', markFeedbackCurrent: true }, d.revision, 'Teacher');
  W.approve(ss.id, 'Bridge Build', 'SUB-1', saved.revision);
  const ex = objectRows(ss.getSheetByName('RefinementExamples'))[0];
  assert.equal(ex.Kind, 'Correction');
  assert.equal(ex.ChangedCriteria, 'C01');
  assert.ok(!/Mia/.test(ex.TeacherNotesDeidentified + ex.FeedbackDeidentified));
  assert.ok(ex.TeacherNotesDeidentified.includes('[STUDENT] only tested one design'));
  const preview = g.Refinement.addExample(ss.id, 'Bridge Build', 'Noah Wilson, your joints were neat and you measured twice before cutting.', false);
  assert.equal(preview.preview, true);
  assert.ok(!preview.data.text.includes('Noah'));
  g.Refinement.addExample(ss.id, 'Bridge Build', 'Noah Wilson, your joints were neat and you measured twice before cutting.', true);
  assert.equal(objectRows(ss.getSheetByName('RefinementExamples')).filter(r => r.Kind === 'FeedbackExample').length, 1);
});
