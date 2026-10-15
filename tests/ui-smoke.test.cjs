/* jsdom tests of the browser client against a mocked google.script.run.
   They prove client behaviour only — not a live Apps Script deployment. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { JSDOM, VirtualConsole } = require('jsdom');
const { renderWebApp } = require('../scripts/render-webapp.cjs');
const html = renderWebApp();

const WB = 'WORKBOOK' + '0'.repeat(30);
function profile() {
  return {
    profileId: 'P1', taskName: 'Lamp', name: 'Lamp rubric', totalMaxMarks: 20,
    gradeScaleBands: [{ letter: 'A', min: 85 }, { letter: 'B', min: 75 }, { letter: 'C', min: 65 }, { letter: 'D', min: 50 }, { letter: 'E', min: 0 }],
    criteria: [
      { criterionId: 'C01', title: 'Design', part: 'Folio', maxMarks: 10, outcome: '', mdOverride: { id: 'C01-MD', text: 'Outstanding' }, bandOrder: ['A', 'C'], bandWeights: { A: 1, C: 0.7 }, bands: { A: [{ id: 'C01-A-1', text: 'Justifies design choices' }], C: [{ id: 'C01-C-1', text: 'Describes design choices' }] } },
      { criterionId: 'C02', title: 'Making', part: 'Practical', maxMarks: 10, outcome: '', mdOverride: { id: 'C02-MD', text: 'Outstanding' }, bandOrder: ['A'], bandWeights: { A: 1 }, bands: { A: [{ id: 'C02-A-1', text: 'Secure joints' }, { id: 'C02-A-2', text: 'Clean finish' }] } }
    ]
  };
}
function detail(id, over = {}) {
  return Object.assign({
    submission: { submissionRecordId: id, studentName: id === 'S2' ? 'Ben Ito' : 'Ava Chen', taskName: 'Lamp', className: '9DAT1', status: 'New', stateLabel: 'Not started',
      version: 1, currentOfficial: true, reassessment: false, sourceType: 'Classroom', hasClassroomIds: true },
    workbook: { id: WB, name: '9DAT1 [2026] - Lamp - Grading' },
    profile: profile(), assessmentId: '', criteriaMap: {}, feedback: {}, attachments: [{ fileName: 'folio.pdf', driveFileId: 'FILEID12345', link: 'https://drive.google.com/x', status: 'Pending', mimeType: 'application/pdf' }],
    history: [], scores: null, flags: { flags: [], blocking: 0, advisory: 0, byCriterion: {} }, snapshot: null, comparison: null, evidenceHistory: [],
    revision: '', readOnly: false, locked: false, canUnlock: false, legacy: false, warning: ''
  }, over);
}

function boot({ responses = {}, narrow = false, storage = {}, speech = null } = {}) {
  const calls = [];
  const errors = [];
  const deferred = {};
  const base = {
    apiWebAuthStatus: () => ({ success: true, data: { authorised: true, email: 'dan@school.edu.au', canUnlock: false } }),
    apiWebListCandidateWorkbooks: () => ({ success: true, data: [{ id: WB, name: '9DAT1 [2026] - Lamp - Grading', courseCode: '9DAT1', year: '2026', taskName: 'Lamp', tasks: ['Lamp'], updatedAt: '2026-10-14T00:00:00Z' }] }),
    apiWebTaskDashboard: () => ({ success: true, data: { workbook: { id: WB, name: '9DAT1 [2026] - Lamp - Grading' }, taskName: 'Lamp', prepared: true,
      profile: { name: 'Lamp rubric', totalMaxMarks: 20, criteriaCount: 2, observableCount: 4 }, profileVersions: 1,
      classroom: { courseId: '1', courseName: '9DAT1', courseWorkId: '2', assignmentTitle: 'Lamp project', dueDate: '', maxPoints: 20, autoImported: false },
      submissions: [
        { submissionRecordId: 'S1', studentName: 'Ava Chen', studentEmail: 'ava@x', className: '9DAT1', version: 1, status: 'Approved', stateLabel: 'Teacher-approved', currentOfficial: true, updatedAt: '' },
        { submissionRecordId: 'S2', studentName: 'Ben Ito', studentEmail: 'ben@x', className: '9DAT1', version: 1, status: 'New', stateLabel: 'Not started', currentOfficial: true, updatedAt: '' }
      ], stats: { total: 2, approved: 1, awaiting: 1, drafts: 0, evidenceChanged: 0 } } }),
    apiWebAiQueueStatus: () => ({ success: true, data: { counts: { Queued: 0, Running: 0, Done: 0, Failed: 0 }, jobs: [] } }),
    apiWebGradeDetail: (wb, task, id) => ({ success: true, data: detail(id) })
  };
  const table = Object.assign({}, base, responses);
  function runner(success, failure) {
    return new Proxy({}, { get: (_t, prop) => {
      if (prop === 'withSuccessHandler') return h => runner(h, failure);
      if (prop === 'withFailureHandler') return h => runner(success, h);
      return (...args) => {
        calls.push([prop, ...args]);
        const respond = () => {
          try { if (!table[prop]) throw new Error('No mock for ' + prop); success(table[prop](...args)); }
          catch (err) { failure(err); }
        };
        if (deferred[prop]) deferred[prop].push(respond); else queueMicrotask(respond);
      };
    } });
  }
  const virtualConsole = new VirtualConsole();
  virtualConsole.on('jsdomError', e => { if (!/Not implemented/.test(e.message)) errors.push(e); if (process.env.DEBUG_UI) console.error(e); });
  const dom = new JSDOM(html, { runScripts: 'dangerously', virtualConsole, url: 'https://script.googleusercontent.com/app#/home', beforeParse(window) {
    window.google = { script: { run: runner() } };
    window.scrollTo = () => {};
    window.confirm = () => true;
    window.matchMedia = q => ({ matches: narrow && /max-width/.test(q), addListener() {}, removeListener() {} });
    Object.entries(storage).forEach(([k, v]) => window.localStorage.setItem(k, v));
    if (speech) window.webkitSpeechRecognition = speech;
  } });
  dom.window.addEventListener('unhandledrejection', e => { errors.push(e.reason); if (process.env.DEBUG_UI) console.error('UNHANDLED', e.reason); });
  return { dom, doc: dom.window.document, calls, table, deferred, errors };
}
async function until(check, what) {
  for (let i = 0; i < 100; i++) { if (check()) return; await new Promise(r => setTimeout(r, 5)); }
  throw new Error('Timed out waiting for ' + what);
}
/* Let follow-up requests (reloads after save/approve/publish) finish before tearing down. */
async function close(dom) { await new Promise(r => setTimeout(r, 60)); dom.window.close(); }
function go(dom, hash) { dom.window.location.hash = hash; }
function tick(doc, dom, id) { const cb = doc.querySelector('[data-obs="' + id + '"]'); cb.checked = !cb.checked; cb.dispatchEvent(new dom.window.Event('change', { bubbles: true })); }

test('landing lists workbooks with class, task and year, and offers to resume the last task (SC01, SC02)', async () => {
  const { dom, doc } = boot({ storage: { 'lessonGrader.last': JSON.stringify({ wb: WB, task: 'Lamp', at: '2026-10-14T00:00:00Z' }) } });
  await until(() => doc.querySelector('#wb-list .wb-card'), 'workbook cards');
  const card = doc.querySelector('#wb-list .wb-card');
  assert.equal(card.querySelector('h3').textContent, 'Lamp');
  assert.match(card.querySelector('.meta').textContent, /9DAT1 · 2026/);
  assert.match(doc.querySelector('#resume').textContent, /Resume Lamp/);
  assert.equal(doc.querySelector('#resume a').getAttribute('href'), '#/task/' + WB + '/Lamp');
  await close(dom);
});

test('teacher marks, saves and approves; save and approval never publish (SC06, SC37, SC39, SC51)', async () => {
  let saved = null;
  const { dom, doc, calls } = boot({ responses: {
    apiWebSaveDraft: (wb, task, sub, criteria, feedback, rev) => { saved = { wb, task, sub, criteria, feedback, rev }; return { success: true, revision: 'R1', savedAt: '2026-10-15T01:00:00Z', message: 'Draft saved.' }; },
    apiWebGradeDetail: (wb, task, id) => ({ success: true, data: saved ? detail(id, { assessmentId: 'A1', revision: 'R1', criteriaMap: saved.criteria, feedback: saved.feedback,
      scores: { criteria: { C01: { points: 10, grade: 'A' }, C02: { points: 5, grade: 'D' } }, totalPoints: 15, totalMaxMarks: 20, percent: 75, letter: 'B' },
      submission: Object.assign(detail(id).submission, { status: 'InReview', stateLabel: 'Draft' }) }) : detail(id) }),
    apiWebApprove: () => ({ success: true, message: 'Approved and saved to ApprovedGrades. Nothing was published to Classroom.' })
  } });
  await until(() => doc.querySelector('#wb-list a'), 'workbooks');
  go(dom, '#/task/' + WB + '/Lamp');
  await until(() => doc.querySelectorAll('#roster li').length === 2, 'roster');
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('[data-obs="C01-A-1"]'), 'criteria');
  assert.equal(doc.querySelector('#student-title').textContent, 'Ben Ito');
  tick(doc, dom, 'C01-A-1'); tick(doc, dom, 'C02-A-1');
  assert.equal(doc.querySelector('#save-status').textContent, 'Unsaved changes');
  assert.equal(doc.querySelector('#live-points').textContent, '15');
  assert.equal(doc.querySelector('#live-letter').textContent, 'B');
  doc.querySelector('#btn-save').click();
  assert.equal(doc.querySelector('#save-status').textContent, 'Saving…');
  await until(() => /Saved to the workbook/.test(doc.querySelector('#save-status').textContent), 'saved status');
  assert.equal(saved.wb, WB); assert.equal(saved.task, 'Lamp'); assert.equal(saved.sub, 'S2'); assert.equal(saved.rev, '');
  assert.equal(saved.criteria.C01.checkboxes['C01-A-1'], true);
  doc.querySelector('#btn-approve').click();
  await until(() => doc.querySelector('#dlg-approve').hasAttribute('open'), 'approval review');
  assert.match(doc.querySelector('#approve-body').textContent, /15 \/ 20 marks, 75%, grade B/);
  assert.match(doc.querySelector('#approve-body').textContent, /Justifies design choices/);
  doc.querySelector('#approve-confirm').click();
  await until(() => calls.some(c => c[0] === 'apiWebApprove'), 'approve call');
  assert.equal(JSON.stringify(calls.find(c => c[0] === 'apiWebApprove').slice(1)), JSON.stringify([WB, 'Lamp', 'S2', 'R1']));
  assert.equal(calls.filter(c => /Publish/.test(c[0])).length, 0);
  await close(dom);
});

test('a failed save is reported as NOT saved and the edits stay on screen (SC51)', async () => {
  const { dom, doc } = boot({ responses: { apiWebSaveDraft: () => { throw new Error('A newer assessment exists. Reload this student before saving.'); } } });
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('[data-obs="C01-A-1"]'), 'criteria');
  tick(doc, dom, 'C01-A-1');
  doc.querySelector('#btn-save').click();
  await until(() => doc.querySelector('#save-status').classList.contains('error'), 'error status');
  assert.match(doc.querySelector('#save-status').textContent, /^Not saved — A newer assessment exists.*still on screen/);
  assert.equal(doc.querySelector('[data-obs="C01-A-1"]').checked, true);
  await close(dom);
});

test('a slow response for a previous student never replaces the student now open (SC05)', async () => {
  const { dom, doc, deferred, calls } = boot();
  deferred.apiWebGradeDetail = [];
  go(dom, '#/studio/' + WB + '/Lamp/S1');
  await until(() => deferred.apiWebGradeDetail.length === 1, 'first request');
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => deferred.apiWebGradeDetail.length === 2, 'second request');
  deferred.apiWebGradeDetail[1]();             // Ben (current) answers first
  await until(() => doc.querySelector('#student-title').textContent === 'Ben Ito', 'Ben');
  deferred.apiWebGradeDetail[0]();             // Ava's stale answer arrives late
  await new Promise(r => setTimeout(r, 20));
  assert.equal(doc.querySelector('#student-title').textContent, 'Ben Ito');
  await close(dom);
});

test('on a phone, flagged criteria come first and open; others are collapsed but available (SC27)', async () => {
  const flagged = { flags: [{ code: 'NO_SELECTION', severity: 'blocking', criterionId: 'C02', message: 'Making: tick evidence.' }], blocking: 1, advisory: 0,
    byCriterion: { C02: [{ code: 'NO_SELECTION', severity: 'blocking', criterionId: 'C02', message: 'Making: tick evidence.' }] } };
  const { dom, doc } = boot({ narrow: true, responses: { apiWebGradeDetail: (w, t, id) => ({ success: true, data: detail(id, { flags: flagged }) }) } });
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelectorAll('#criteria details').length === 2, 'criteria');
  const cards = [...doc.querySelectorAll('#criteria details')];
  assert.equal(cards[0].getAttribute('data-cid'), 'C02');
  assert.equal(cards[0].open, true);
  assert.equal(cards[1].getAttribute('data-cid'), 'C01');
  assert.equal(cards[1].open, false);
  assert.ok(cards[1].querySelector('[data-obs="C01-A-1"]'), 'collapsed criterion is still in the page');
  await close(dom);
});

test('approval is blocked while a blocking flag is unresolved (SC38)', async () => {
  const flags = { flags: [{ code: 'FEEDBACK_NOT_REVIEWED', severity: 'blocking', criterionId: '', message: 'Feedback was written before the latest change.' }], blocking: 1, advisory: 0, byCriterion: {} };
  const { dom, doc } = boot({ responses: { apiWebGradeDetail: (w, t, id) => ({ success: true, data: detail(id, { assessmentId: 'A1', revision: 'R1', flags,
    scores: { criteria: {}, totalPoints: 0, totalMaxMarks: 20, percent: 0, letter: 'E' } }) }) } });
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('#student-title').textContent === 'Ben Ito', 'student loaded');
  doc.querySelector('#btn-approve').click();
  await until(() => doc.querySelector('#dlg-approve').hasAttribute('open'), 'dialog');
  assert.equal(doc.querySelector('#approve-confirm').disabled, true);
  assert.match(doc.querySelector('#approve-body').textContent, /Resolve before approving/);
  await close(dom);
});

test('when AI is unavailable the teacher can queue the student or continue manually (SC48)', async () => {
  const { dom, doc, calls } = boot({ responses: {
    apiWebRunAi: () => ({ success: false, unavailable: true, message: 'Gemini busy (503).' }),
    apiWebQueueAiDrafts: () => ({ success: true, queued: 1, message: '1 queued.' })
  } });
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('[data-obs="C01-A-1"]'), 'criteria');
  doc.querySelector('#btn-ai').click();
  await until(() => doc.querySelector('#dlg-confirm').hasAttribute('open'), 'choice dialog');
  assert.match(doc.querySelector('#confirm-body').textContent, /Gemini busy/);
  const queue = [...doc.querySelectorAll('#confirm-actions button')].find(b => b.textContent === 'Queue for later');
  queue.click();
  await until(() => calls.some(c => c[0] === 'apiWebQueueAiDrafts'), 'queue call');
  assert.equal(JSON.stringify(calls.find(c => c[0] === 'apiWebQueueAiDrafts').slice(1, 4)), JSON.stringify([WB, 'Lamp', ['S2']]));
  // Manual marking still available.
  assert.equal(doc.querySelector('[data-obs="C01-A-1"]').disabled, false);
  await close(dom);
});

test('dictation appends editable text, never saves, and explains permission problems (SC29, SC30)', async () => {
  const instances = [];
  class FakeRecognition { constructor() { instances.push(this); } start() { this.started = true; } stop() { this.onend && this.onend(); } }
  const { dom, doc, calls } = boot({ speech: FakeRecognition });
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('#student-title').textContent === 'Ben Ito', 'studio');
  doc.querySelector('#fb-well').value = 'Typed start.';
  doc.querySelector('[data-dictate="fb-well"]').click();
  const rec = instances[0];
  assert.equal(rec.lang, 'en-AU'); assert.equal(rec.continuous, true);
  assert.equal(doc.querySelector('[data-dictate="fb-well"]').getAttribute('aria-pressed'), 'true');
  const result = (text, isFinal) => { const r = [{ transcript: text }]; r.isFinal = isFinal; return r; };
  rec.onresult({ resultIndex: 0, results: [result('your joints are sekure', true)] });
  assert.equal(doc.querySelector('#fb-well').value, 'Typed start. your joints are sekure');
  assert.equal(doc.querySelector('#save-status').textContent, 'Unsaved changes');
  doc.querySelector('[data-dictate="fb-well"]').click();          // stop
  assert.match(doc.querySelector('[data-dictation-status="fb-well"]').textContent, /Check and edit it/);
  doc.querySelector('#fb-well').value = doc.querySelector('#fb-well').value.replace('sekure', 'secure');
  assert.equal(calls.filter(c => c[0] === 'apiWebSaveDraft' || c[0] === 'apiWebApprove').length, 0, 'dictation never saves or approves');
  doc.querySelector('[data-dictate="note-C01"]').click();
  instances[1].onerror({ error: 'not-allowed' }); instances[1].onend();
  assert.match(doc.querySelector('[data-dictation-status="note-C01"]').textContent, /permission was denied/);
  await close(dom);
});

test('dictation falls back to typing where speech recognition is unavailable', async () => {
  const { dom, doc } = boot();
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('#student-title').textContent === 'Ben Ito', 'studio');
  doc.querySelector('[data-dictate="fb-well"]').click();
  assert.match(doc.querySelector('[data-dictation-status="fb-well"]').textContent, /not available in this browser/);
  await close(dom);
});

test('publishing previews the selected students, needs PUBLISH typed, and sends only those students (SC40)', async () => {
  const { dom, doc, calls } = boot({ responses: {
    apiWebPublishPreview: (wb, task, ids) => ({ success: true, data: { destination: { assignmentTitle: 'Lamp project', courseName: '9DAT1', maxPoints: 20 }, mode: 'scaled-points', previewToken: 'tok1',
      eligible: [{ id: 'S1', name: 'Ava Chen', letter: 'A', points: 19, maxMarks: 20, grade: 19 }], ineligible: [] } }),
    apiWebPublish: () => ({ success: true, message: 'Published 1 of 1 student(s).', data: { results: [{ name: 'Ava Chen', published: true, message: 'Grade 19 / 20 saved as a draft grade.' }] } })
  } });
  go(dom, '#/task/' + WB + '/Lamp');
  await until(() => doc.querySelector('[data-select="S1"]'), 'roster');
  const box = doc.querySelector('[data-select="S1"]'); box.checked = true; box.dispatchEvent(new dom.window.Event('change', { bubbles: true }));
  doc.querySelector('#btn-publish').click();
  await until(() => doc.querySelector('#dlg-publish').hasAttribute('open'), 'publish dialog');
  assert.match(doc.querySelector('#publish-body').textContent, /"Lamp project" in 9DAT1/);
  doc.querySelector('#publish-typed').value = 'publish';
  doc.querySelector('#publish-confirm').click();
  assert.equal(calls.filter(c => c[0] === 'apiWebPublish').length, 0);
  doc.querySelector('#publish-typed').value = 'PUBLISH';
  doc.querySelector('#publish-confirm').click();
  await until(() => calls.some(c => c[0] === 'apiWebPublish'), 'publish call');
  const call = calls.find(c => c[0] === 'apiWebPublish');
  assert.equal(JSON.stringify(call.slice(1, 6)), JSON.stringify([WB, 'Lamp', ['S1'], 'tok1', 'PUBLISH']));
  assert.equal(call[6].returnToStudents, false);
  await close(dom);
});

test('final-locked work is read-only and Unlock appears only for authorised teachers (SC41)', async () => {
  const locked = id => detail(id, { readOnly: true, locked: true, canUnlock: false, assessmentId: 'A1',
    submission: Object.assign(detail(id).submission, { status: 'Locked', stateLabel: 'Final (locked)' }) });
  let auth = false;
  const { dom, doc } = boot({ responses: { apiWebGradeDetail: (w, t, id) => ({ success: true, data: Object.assign(locked(id), { canUnlock: auth }) }) } });
  go(dom, '#/studio/' + WB + '/Lamp/S1');
  await until(() => doc.querySelector('[data-obs="C01-A-1"]'), 'criteria');
  assert.equal(doc.querySelector('[data-obs="C01-A-1"]').disabled, true);
  assert.equal(doc.querySelector('#fb-well').disabled, true);
  assert.equal(doc.querySelector('#btn-save').hidden, true);
  assert.equal(doc.querySelector('#btn-unlock').hidden, true);
  assert.equal(doc.querySelector('#student-state').textContent, 'Final (locked)');
  auth = true;
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('#student-title').textContent === 'Ben Ito', 'reload');
  assert.equal(doc.querySelector('#btn-unlock').hidden, false);
  await close(dom);
});

test('settings save the Gemini key without echoing it, and never send it inside settings', async () => {
  const key = 'THIS_IS_A_TEST_KEY_NOT_REAL_123456';
  const { dom, doc, calls } = boot({ responses: {
    apiWebGetSettings: () => ({ success: true, data: { settings: { outputFormat: { gradeDisplay: 'letter', classroomGradeMode: 'scaled-points', useFirstName: true, includeGradeInComment: true }, feedbackTemplate: '', aiProviders: { draftProvider: 'gemini', visionProvider: 'none', geminiModel: 'gemini-2.5-flash' } },
      providers: { gemini: { configured: false }, qwen: { configured: false }, jev: { configured: false } } } }),
    apiWebTriggerStatus: () => ({ success: true, data: { hourlySync: false, aiQueue: false, note: '' } }),
    apiWebPromptList: () => ({ success: true, data: { active: 'builtin-v2', versions: [{ id: 'builtin-v2', status: 'builtin', text: 'x', tests: [] }] } }),
    apiWebSetGeminiKey: () => ({ success: true, message: 'Key saved.' }),
    apiWebSaveSettings: () => ({ success: true, data: {} })
  } });
  go(dom, '#/settings');
  await until(() => doc.querySelector('#out-grade').value === 'letter', 'settings loaded');
  doc.querySelector('#gemini-key').value = key;
  doc.querySelector('#gemini-save').click();
  await until(() => doc.querySelector('#gemini-key').value === '', 'key cleared');
  doc.querySelector('#ai-save').click();
  doc.querySelector('#out-save').click();
  await until(() => calls.filter(c => c[0] === 'apiWebSaveSettings').length === 2, 'settings saves');
  assert.ok(!JSON.stringify(calls.filter(c => c[0] === 'apiWebSaveSettings')).includes(key));
  assert.ok(!doc.body.textContent.includes(key));
  await close(dom);
});

test('leaving a student with unsaved changes asks first and stays put when cancelled', async () => {
  const { dom, doc } = boot();
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('[data-obs="C01-A-1"]'), 'criteria');
  tick(doc, dom, 'C01-A-1');
  dom.window.confirm = () => false;
  go(dom, '#/home');
  await new Promise(r => setTimeout(r, 30));
  assert.equal(doc.querySelector('#view-studio').hidden, false);
  assert.equal(doc.querySelector('[data-obs="C01-A-1"]').checked, true);
  await close(dom);
});

test('student and AI text is rendered as text, never as HTML', async () => {
  const hostile = '<img src=x onerror="window.__pwned=1">';
  const { dom, doc } = boot({ responses: { apiWebGradeDetail: (w, t, id) => ({ success: true, data: detail(id, {
    submission: Object.assign(detail(id).submission, { studentName: hostile }),
    criteriaMap: { C01: { checkboxes: {}, teacherWrittenNote: hostile, aiProposal: { checks: { 'C01-A-1': true }, unsupported: [] }, aiEvidence: [{ checkboxId: 'C01-A-1', fileName: hostile, note: hostile }] } } }) }) } });
  go(dom, '#/studio/' + WB + '/Lamp/S2');
  await until(() => doc.querySelector('#student-title').textContent === hostile, 'title');
  await new Promise(r => setTimeout(r, 20));
  assert.equal(dom.window.__pwned, undefined);
  assert.equal(doc.querySelectorAll('#view-studio img').length, 0);
  await close(dom);
});
