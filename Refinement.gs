/* ============================================================================
 * Refinement.gs — de-identified examples for improving prompts. (SC45)
 *
 * - On approval, if an AI proposal existed, the difference between the AI
 *   ticks and the teacher's final ticks is recorded with de-identified notes
 *   and feedback ("Correction").
 * - Teachers can paste exemplar feedback ("FeedbackExample"); it is
 *   de-identified against the whole class roster before storage.
 * Examples are stored in the workbook's RefinementExamples sheet. Nothing here
 * is sent to any AI service automatically.
 * ============================================================================ */
var Refinement = (function () {
  var SHEET = 'RefinementExamples';
  var HEADERS = ['ExampleID', 'Kind', 'TaskName', 'RubricProfileID', 'AIChecksJSON', 'TeacherChecksJSON', 'ChangedCriteria',
    'TeacherNotesDeidentified', 'FeedbackDeidentified', 'RemovedIdentifiers', 'CreatedAt', 'CreatedBy'];

  function text(v) { return v == null ? '' : String(v); }

  function identity(ss, sub) {
    return { names: [text(sub.StudentName)], rosterNames: WebGrading.records(ss, 'Submissions').map(function (r) { return text(r.StudentName); }),
      ids: [text(sub.StudentUserID), text(sub.StudentEmail), text(sub.SubmissionRecordID)].filter(Boolean) };
  }

  /** Called from WebGrading.approve; never blocks approval. */
  function recordApproval(ss, taskName, sub, profile, states, feedback, user) {
    var ai = {}, teacher = {}, changed = [];
    profile.criteria.forEach(function (c) {
      var st = states[c.criterionId] || {};
      if (!st.aiProposal) return;
      ai[c.criterionId] = AssessmentFlags.ticked(st.aiProposal.checks);
      teacher[c.criterionId] = AssessmentFlags.ticked(st.checkboxes);
      if (AssessmentFlags.canonical(ai[c.criterionId]) !== AssessmentFlags.canonical(teacher[c.criterionId])) changed.push(c.criterionId);
    });
    if (!Object.keys(ai).length) return null;
    var id = identity(ss, sub);
    var removed = 0;
    function clean(t) { var r = Deidentify.apply(t, id); Object.keys(r.counts).forEach(function (k) { removed += r.counts[k]; }); return r.text; }
    var notes = profile.criteria.map(function (c) { var n = text((states[c.criterionId] || {}).teacherWrittenNote).trim(); return n ? c.criterionId + ': ' + clean(n) : ''; }).filter(Boolean).join('\n');
    var fb = ['whatWentWell', 'areasForImprovement', 'goalsForNextAssessment'].map(function (k) { return text(feedback[k]).trim() ? clean(feedback[k]) : ''; }).filter(Boolean).join('\n\n');
    WebGrading.ensureSheet(ss, SHEET, HEADERS);
    WebGrading.append(ss, SHEET, { ExampleID: Utilities.getUuid(), Kind: 'Correction', TaskName: taskName, RubricProfileID: profile.profileId,
      AIChecksJSON: JSON.stringify(ai), TeacherChecksJSON: JSON.stringify(teacher), ChangedCriteria: changed.join(','),
      TeacherNotesDeidentified: notes.slice(0, 20000), FeedbackDeidentified: fb.slice(0, 20000), RemovedIdentifiers: removed,
      CreatedAt: new Date().toISOString(), CreatedBy: user });
    return { changed: changed };
  }

  /** Stores a teacher-supplied feedback example after de-identifying it against the roster. */
  function addExample(workbookId, taskName, exampleText, confirmed) {
    var user = WebAccess.requireTeacher();
    var target = WebGrading.openTask(workbookId, taskName);
    var t = text(exampleText).trim();
    if (t.length < 20 || t.length > 8000) throw new Error('Paste an example of 20–8,000 characters.');
    var roster = WebGrading.records(target.ss, 'Submissions').map(function (r) { return text(r.StudentName); });
    var cleaned = Deidentify.apply(t, { rosterNames: roster });
    if (confirmed !== true) {
      return { success: true, preview: true, data: { text: cleaned.text, counts: cleaned.counts },
        message: 'Check the de-identified text. Names not on this class roster cannot be detected automatically.' };
    }
    WebGrading.ensureSheet(target.ss, SHEET, HEADERS);
    WebGrading.append(target.ss, SHEET, { ExampleID: Utilities.getUuid(), Kind: 'FeedbackExample', TaskName: target.task,
      FeedbackDeidentified: cleaned.text, RemovedIdentifiers: Object.keys(cleaned.counts).reduce(function (s, k) { return s + cleaned.counts[k]; }, 0),
      CreatedAt: new Date().toISOString(), CreatedBy: user });
    return { success: true, message: 'De-identified example saved to RefinementExamples.' };
  }

  function list(workbookId, taskName) {
    WebAccess.requireTeacher();
    var target = WebGrading.openTask(workbookId, taskName);
    var rows = target.ss.getSheetByName(SHEET) ? WebGrading.records(target.ss, SHEET).filter(function (r) { return text(r.TaskName) === target.task; }) : [];
    return { success: true, data: rows.slice(-50).reverse().map(function (r) {
      return { kind: text(r.Kind), changedCriteria: text(r.ChangedCriteria), feedback: text(r.FeedbackDeidentified).slice(0, 600), createdAt: text(r.CreatedAt) };
    }) };
  }

  return { HEADERS: HEADERS, recordApproval: recordApproval, addExample: addExample, list: list };
})();

function apiWebAddRefinementExample(workbookId, taskName, text, confirmed) { return Refinement.addExample(workbookId, taskName, text, confirmed); }
function apiWebListRefinementExamples(workbookId, taskName) { return Refinement.list(workbookId, taskName); }
