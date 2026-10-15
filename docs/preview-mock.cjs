/* Browser-only design preview. No Google data, no persistence beyond this page,
   no real API calls. Served ONLY by docs/preview.cjs, never deployed to Apps Script.
   Mirrors the shape of the real apiWeb* responses closely enough to exercise the UI. */
(function () {
  'use strict';
  window.__LESSON_GRADER_PREVIEW__ = true;
  var WB = 'PREVIEW_WORKBOOK_LAMP_9DAT1_2026';
  var WB2 = 'PREVIEW_WORKBOOK_CLOCK_8TECH_2026';
  var bands = [{ letter: 'A', min: 85, max: 100 }, { letter: 'B', min: 75, max: 84.99 }, { letter: 'C', min: 65, max: 74.99 }, { letter: 'D', min: 50, max: 64.99 }, { letter: 'E', min: 0, max: 49.99 }];
  var W = { A: 1, B: 0.875, C: 0.7, D: 0.575, E: 0.25 };
  function crit(id, title, part, marks, obs) {
    var b = {}; var order = [];
    Object.keys(obs).forEach(function (band) { order.push(band); b[band] = obs[band].map(function (t, i) { return { id: id + '-' + band + '-' + (i + 1), text: t }; }); });
    var w = {}; order.forEach(function (k) { w[k] = W[k]; });
    return { criterionId: id, title: title, part: part, section: '', maxMarks: marks, outcome: '', mdOverride: { id: id + '-MD', text: 'Distinct evidence: outstanding work beyond the listed observables.' }, bandOrder: order, bandWeights: w, bands: b };
  }
  var profile = { schemaVersion: '1.0', profileId: 'PREVIEW_PROFILE', taskName: 'Desk lamp', name: 'Desk lamp — design & make', source: 'sharedRubric', gradeScaleBands: bands, totalMaxMarks: 40, parts: ['Folio', 'Practical'],
    criteria: [
      crit('C01', 'Research and ideas', 'Folio', 10, { A: ['Evaluates several existing lamps and justifies choices', 'Links research to design decisions'], B: ['Explains relevant existing designs'], C: ['Describes some existing lamps'], E: ['Attempts research'] }),
      crit('C02', 'Planning and safety', 'Folio', 10, { A: ['Realistic plan with safety checkpoints', 'Records changes to the plan with reasons'], C: ['Lists the main steps'], E: ['Partial plan'] }),
      crit('C03', 'Making', 'Practical', 20, { A: ['Accurate marking out and cutting', 'Secure, clean joints', 'High-quality finish'], C: ['Mostly accurate cutting', 'Functional joints'], E: ['Attempts construction'] })
    ] };
  var names = ['Ava Chen', 'Ben Ito', 'Chloe Nguyen', 'Dylan Moore', 'Ella Rossi', 'Finn Walker'];
  var subs = names.map(function (n, i) {
    return { submissionRecordId: 'PREVIEW_SUB_' + i, studentName: n, studentEmail: n.toLowerCase().replace(' ', '.') + '@school.example', className: '9DAT1', version: 1,
      status: ['Approved', 'InReview', 'New', 'InReview', 'Locked', 'New'][i], currentOfficial: true, reassessment: false, updatedAt: '2026-10-1' + i + 'T03:00:00Z',
      late: i === 3, evidenceChangedAt: i === 1 ? '2026-10-14T05:00:00Z' : '', published: i === 0 ? 'Published 34.5/40 · draft' : '', hasClassroomIds: true };
  });
  function label(s) { return { Approved: 'Teacher-approved', Locked: 'Final (locked)', InReview: 'Draft', New: 'Not started' }[s.status]; }
  subs.forEach(function (s) { s.stateLabel = label(s); });
  var states = {
    PREVIEW_SUB_1: { C01: { checkboxes: { 'C01-B-1': true }, teacherWrittenNote: 'Good comparison table of three lamps.', aiProposal: { checks: { 'C01-B-1': true, 'C01-A-2': false }, insufficient: false, unsupported: ['C01-A-1'] },
      aiEvidence: [{ checkboxId: 'C01-B-1', fileName: 'Lamp folio.pdf', note: 'Page 3: comparison table' }] },
      C02: { checkboxes: {}, aiProposal: { checks: {}, insufficient: true, unsupported: [] } },
      C03: { checkboxes: { 'C03-A-1': true, 'C03-C-2': true }, aiProposal: { checks: { 'C03-A-1': true }, insufficient: false, unsupported: [] }, aiEvidence: [{ checkboxId: 'C03-A-1', fileName: 'Making photos.jpg', note: 'Clean mitre on base' }] } }
  };
  var feedback = { PREVIEW_SUB_1: { whatWentWell: 'Research and ideas: Explains relevant existing designs.', areasForImprovement: '', goalsForNextAssessment: '', basisHash: '' } };
  function score(st) {
    var out = { criteria: {}, totalPoints: 0, totalMaxMarks: 40 };
    profile.criteria.forEach(function (c) {
      var ch = (st[c.criterionId] || {}).checkboxes || {}, t = 0;
      if (ch[c.mdOverride.id]) t = c.maxMarks; else c.bandOrder.forEach(function (b) { var it = c.bands[b]; t += it.filter(function (o) { return ch[o.id]; }).length * c.maxMarks * c.bandWeights[b] / it.length; });
      var raw = t; t = Math.min(t, c.maxMarks);
      var pct = t / c.maxMarks * 100;
      out.criteria[c.criterionId] = { points: t, rawPoints: raw, maxMarks: c.maxMarks, percent: pct, grade: letter(pct) };
      out.totalPoints += t;
    });
    out.percent = out.totalPoints / 40 * 100; out.letter = letter(out.percent); return out;
  }
  function letter(p) { for (var i = 0; i < bands.length; i++) if (p >= bands[i].min) return bands[i].letter; return 'E'; }
  function flagsFor(id, st, fb) {
    var f = [];
    profile.criteria.forEach(function (c) {
      var s = st[c.criterionId] || {};
      var any = Object.keys(s.checkboxes || {}).some(function (k) { return s.checkboxes[k]; });
      if (!any && !(s.review && s.review.noEvidence)) f.push({ code: 'NO_SELECTION', severity: 'blocking', criterionId: c.criterionId, message: c.title + ': tick the evidenced observables, or confirm there is no creditable evidence (0 marks).' });
      if (s.aiProposal && s.aiProposal.insufficient && !(s.review && (s.review.reviewed || s.review.noEvidence))) f.push({ code: 'AI_INSUFFICIENT_EVIDENCE', severity: 'blocking', criterionId: c.criterionId, message: c.title + ': the AI found no supported evidence. Check the work yourself, then mark this criterion reviewed.' });
      if (s.aiProposal && s.aiProposal.unsupported && s.aiProposal.unsupported.length) f.push({ code: 'AI_UNSUPPORTED_TICKS', severity: 'advisory', criterionId: c.criterionId, message: c.title + ': 1 AI suggestion discarded because no evidence was cited.' });
    });
    if (fb && fb.whatWentWell && !fb.basisHash) f.push({ code: 'FEEDBACK_NOT_REVIEWED', severity: 'blocking', criterionId: '', message: 'Feedback was written before the latest change to selections or notes. Update it, or confirm it still matches.' });
    var by = {}; f.forEach(function (x) { if (x.criterionId) (by[x.criterionId] = by[x.criterionId] || []).push(x); });
    return { flags: f, blocking: f.filter(function (x) { return x.severity === 'blocking'; }).length, advisory: f.filter(function (x) { return x.severity === 'advisory'; }).length, byCriterion: by };
  }
  function detail(id) {
    var s = subs.filter(function (x) { return x.submissionRecordId === id; })[0] || subs[2];
    var st = states[id] || {};
    var fb = feedback[id] || {};
    var ro = s.status === 'Approved' || s.status === 'Locked';
    return { submission: Object.assign({}, s, { taskName: 'Desk lamp', sourceType: 'Classroom', approvedAt: '2026-10-12T02:00:00Z', lockedAt: '2026-10-13T02:00:00Z', hasPendingEvidence: false }),
      workbook: { id: WB, name: '9DAT1 [2026] - Desk lamp - Grading' }, profile: profile, assessmentId: st.C01 ? 'PREVIEW_ASSESS' : '', criteriaMap: st, feedback: fb,
      attachments: [{ fileName: 'Lamp folio.pdf', driveFileId: 'PREVIEWFILE001', link: 'https://drive.google.com/preview-only', status: st.C01 ? 'Reviewed' : 'Pending', mimeType: 'application/pdf' },
        { fileName: 'Making photos.jpg', driveFileId: 'PREVIEWFILE002', link: 'https://drive.google.com/preview-only', status: st.C01 ? 'Reviewed' : 'Pending', mimeType: 'image/jpeg' },
        { fileName: 'lamp-model.stl', driveFileId: '', link: '', status: 'Unsupported', mimeType: 'model/stl', limitations: 'File type is not supported for automatic reading.' }],
      history: [{ action: 'ClassroomImported', when: '2026-10-10T01:00:00Z', by: 'preview@school.example', notes: 'Classroom evidence imported; no grade changed.' }],
      scores: score(st), flags: flagsFor(id, st, fb), snapshot: null, comparison: null, evidenceHistory: [], revision: 'R-' + id, readOnly: ro, locked: s.status === 'Locked', canUnlock: false, legacy: false, warning: '' };
  }
  function ok(data, message) { return { success: true, data: data, message: message || '' }; }
  var api = {
    apiWebAuthStatus: function () { return ok({ authorised: true, email: 'teacher@school.example', canUnlock: false }); },
    apiWebListCandidateWorkbooks: function () { return ok([
      { id: WB, name: '9DAT1 [2026] - Desk lamp - Grading', courseCode: '9DAT1', year: '2026', taskName: 'Desk lamp', tasks: ['Desk lamp'], updatedAt: '2026-10-14T04:00:00Z' },
      { id: WB2, name: '8TECH [2026] - Clock - Grading', courseCode: '8TECH', year: '2026', taskName: 'Clock', tasks: ['Clock'], updatedAt: '2026-10-02T04:00:00Z' }]); },
    apiWebTaskDashboard: function () { return ok({ workbook: { id: WB, name: '9DAT1 [2026] - Desk lamp - Grading' }, taskName: 'Desk lamp', prepared: true,
      profile: { name: profile.name, totalMaxMarks: 40, criteriaCount: 3, observableCount: 16 }, profileVersions: 1,
      classroom: { courseId: '1', courseName: '9DAT1', courseWorkId: '2', assignmentTitle: 'Desk lamp project', dueDate: '2026-10-09T12:59:00Z', maxPoints: 40, autoImported: true },
      submissions: subs, stats: { total: 6, approved: 2, awaiting: 4, drafts: 2, evidenceChanged: 1 } }); },
    apiWebAiQueueStatus: function () { return ok({ counts: { Queued: 0, Running: 0, Done: 2, Failed: 0, Skipped: 0, Cancelled: 0 }, jobs: [] }); },
    apiWebGradeDetail: function (wb, task, id) { return ok(detail(id)); },
    apiWebSaveDraft: function (wb, task, id, criteria, fb) {
      var prior = states[id] || {};
      Object.keys(criteria).forEach(function (k) { criteria[k].aiProposal = (prior[k] || {}).aiProposal; criteria[k].aiEvidence = (prior[k] || {}).aiEvidence; });
      states[id] = criteria; fb.basisHash = fb.markFeedbackCurrent ? 'hpreview' : ''; feedback[id] = fb;
      subs.forEach(function (s) { if (s.submissionRecordId === id && s.status === 'New') { s.status = 'InReview'; s.stateLabel = 'Draft'; } });
      return { success: true, revision: 'R2', savedAt: new Date().toISOString(), message: 'Draft saved to the workbook (preview only).' };
    },
    apiWebBuildFeedback: function () { return ok({ whatWentWell: 'Research and ideas: Explains relevant existing designs.\nMaking: Accurate marking out and cutting.', areasForImprovement: 'Research and ideas: next, aim to show that you evaluate several existing lamps and justify choices.', goalsForNextAssessment: '1. Planning and safety — Realistic plan with safety checkpoints.', sources: { ticked: 2, targets: 1, notes: 0 } }); },
    apiWebOutputPreview: function () { return ok({ basedOn: 'current draft (not approved)', comment: { text: 'Ben, here is your feedback for Desk lamp.\n\nResult: C (27.5 / 40)\n\nResearch and ideas: Explains relevant existing designs.', warnings: [] }, mapping: { destination: { maxPoints: 40 }, grade: { grade: 27.5, note: '27.5 / 40 scaled to 40' } } }); },
    apiWebApprove: function (wb, task, id) { subs.forEach(function (s) { if (s.submissionRecordId === id) { s.status = 'Approved'; s.stateLabel = 'Teacher-approved'; } }); return { success: true, message: 'Approved (preview only). Nothing was published.' }; },
    apiWebRunAi: function () { return { success: false, unavailable: true, message: 'Preview mode has no AI service.' }; },
    apiWebQueueAiDrafts: function (wb, task, ids) { return { success: true, queued: ids.length, message: ids.length + ' queued (preview only).' }; },
    apiWebProcessAiQueue: function () { return { success: true, counts: { Done: 0, Skipped: 0, Failed: 0, Waiting: 1 }, remaining: 1, message: 'AI service unavailable in preview.' }; },
    apiWebImportTask: function () { return { success: true, message: 'Synced 6 Classroom submissions: 0 new, 0 updated, 0 held, 6 unchanged (preview).', stats: {} }; },
    apiWebPublishPreview: function (wb, task, ids) { return ok({ destination: { assignmentTitle: 'Desk lamp project', courseName: '9DAT1', maxPoints: 40 }, mode: 'scaled-points', previewToken: 'preview',
      eligible: subs.filter(function (s) { return ids.indexOf(s.submissionRecordId) !== -1 && (s.status === 'Approved' || s.status === 'Locked'); }).map(function (s) { return { id: s.submissionRecordId, name: s.studentName, letter: 'B', points: 31.5, maxMarks: 40, grade: 31.5, alreadyPublished: s.published }; }),
      ineligible: subs.filter(function (s) { return ids.indexOf(s.submissionRecordId) !== -1 && s.status !== 'Approved' && s.status !== 'Locked'; }).map(function (s) { return { id: s.submissionRecordId, name: s.studentName, reason: 'Not approved.' }; }) }); },
    apiWebPublish: function () { return { success: false, message: 'Preview mode cannot publish.', data: { results: [] } }; },
    apiWebRubricSource: function () { var rows = []; profile.criteria.forEach(function (c) { c.bandOrder.forEach(function (b) { c.bands[b].forEach(function (o) { rows.push({ criterion: c.title, part: c.part, maxMarks: c.maxMarks, band: b, description: o.text, outcome: '', weight: c.bandWeights[b] }); }); }); });
      return ok({ profile: profile, summary: { name: profile.name, criteriaCount: 3, observableCount: rows.length, totalMaxMarks: 40 }, activeRows: rows, sheetRows: rows, drift: false,
        versions: [{ profileId: 'P', name: profile.name, active: true }], gradeScale: { bands: bands, weights: W },
        referencePoints: { matchesTarget: false, message: 'Configured weights differ from the confirmed reference points (C 70% vs 50%, E 25% vs 0%). Existing results are not recalculated.' } }); },
    apiWebListSharedRubrics: function () { return ok([{ id: 'PREVIEW_RUBRIC_LAMP', name: 'Desk lamp — design & make' }]); },
    apiWebListTeacherCourses: function () { return { success: true, courses: [{ id: '101', name: '9DAT1 Design & Technology', section: '' }, { id: '102', name: '8TECH Technology', section: '' }] }; },
    apiWebListCourseWork: function () { return { success: true, coursework: [{ id: '201', title: 'Desk lamp project', dueDate: '2026-10-09', maxPoints: 40, state: 'PUBLISHED' }, { id: '202', title: 'Workshop safety quiz', dueDate: '', maxPoints: 0, state: 'PUBLISHED' }] }; },
    apiWebRegisterPreview: function () { return ok({ context: { taskName: 'Desk lamp', courseCode: '9DAT1', year: '2026', className: '9DAT1', assignmentTitle: 'Desk lamp project', dates: 'Due 9 Oct 2026; Marked 12 Oct 2026 – 13 Oct 2026', outcomes: '', gradeDistribution: 'A: 0, B: 1, C: 1, D: 0, E: 0 (2 approved)' }, destination: { type: 'sheet', name: 'EvaluationRegister sheet in this workbook' } }); },
    apiWebGetSettings: function () { return ok({ settings: { outputFormat: { gradeDisplay: 'letter+mark', classroomGradeMode: 'scaled-points', useFirstName: true, includeGradeInComment: true }, feedbackTemplate: '{{opening}}\n\nResult: {{grade_display}}\n\n{{what_went_well}}', aiProviders: { draftProvider: 'gemini', visionProvider: 'none', geminiModel: 'gemini-2.5-flash', qwenModel: 'qwen3-vl', jevModel: '' } },
      providers: { gemini: { configured: false }, qwen: { configured: false, host: '' }, jev: { configured: false, host: '' } } }); },
    apiWebTriggerStatus: function () { return ok({ hourlySync: false, aiQueue: false, note: 'Preview only.' }); },
    apiWebPromptList: function () { return ok({ active: 'builtin-v2', versions: [{ id: 'builtin-v2', status: 'builtin', text: 'Built-in drafting prompt (preview).', tests: [] }] }); },
    apiWebGetGradeScale: function () { return ok({ bands: bands, weights: W }); },
    apiWebHealthCheck: function () { return ok({ counts: { pass: 2, warn: 3, fail: 0, info: 1 }, checks: [{ area: 'Configuration', check: 'Teacher allowlist', status: 'pass', detail: 'Preview' }, { area: 'Sync', check: 'Hourly Classroom sync trigger', status: 'warn', detail: 'Not installed (preview).' }] }); },
    apiWebGetClassTableData: function () { return ok({ rows: subs.map(function (s) { return { studentName: s.studentName, status: s.status, finalScore: s.status === 'Approved' ? 31.5 : null, percent: s.status === 'Approved' ? 78.75 : null, grade: s.status === 'Approved' ? 'B' : 'Not approved' }; }) }); }
  };
  function runner(success, failure) {
    return new Proxy({}, { get: function (_t, prop) {
      if (prop === 'withSuccessHandler') return function (h) { return runner(h, failure); };
      if (prop === 'withFailureHandler') return function (h) { return runner(success, h); };
      return function () {
        var args = Array.prototype.slice.call(arguments);
        setTimeout(function () {
          try { if (!api[prop]) throw new Error('Not available in the design preview: ' + String(prop)); success(api[prop].apply(null, args)); }
          catch (err) { failure && failure(err); }
        }, 120);
      };
    } });
  }
  window.google = { script: { run: runner() } };
})();
