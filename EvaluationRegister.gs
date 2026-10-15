/* ============================================================================
 * EvaluationRegister.gs — evaluation / register entries for a class task.
 *
 * Pre-fills ONLY facts that exist in the workbook (task, class, course code,
 * dates, outcomes, approved grade distribution). Teacher evaluation fields
 * start blank; the teacher types or dictates them, previews, then commits.
 * Destinations: a Google Doc the teacher can edit (appended section) and/or
 * the workbook's EvaluationRegister sheet. Every result reports exactly what
 * was written. (REVIEW.md F03, SC53, SC54)
 * ============================================================================ */
var EvaluationRegister = (function () {
  var SHEET = 'EvaluationRegister';
  var HEADERS = ['EntryID', 'TaskName', 'CourseCode', 'ClassName', 'Year', 'Dates', 'Outcomes', 'GradeDistribution',
    'TeacherEvaluation', 'TeachingVariations', 'Adjustments', 'StudentFeedbackSummary', 'Destination',
    'DestinationID', 'DestinationName', 'CommittedAt', 'CommittedBy', 'Status'];
  var TEXT_FIELDS = { teacherEvaluation: 6000, teachingVariations: 4000, adjustments: 4000, studentFeedbackSummary: 4000 };
  var DOC_ID_RE = /^[A-Za-z0-9_-]{20,200}$/;

  function text(v) { return v == null ? '' : String(v); }

  /** Parses "CODE [YEAR] - Task - Grading". Pure. */
  function parseWorkbookName(name) {
    var m = /^\s*([^\[]+?)\s*\[(\d{4})\]\s*-\s*(.+?)\s*-\s*Grading\s*$/.exec(text(name));
    return m ? { courseCode: m[1], year: m[2], taskName: m[3] } : { courseCode: '', year: '', taskName: '' };
  }

  function sydneyDate(value) {
    if (!value) return '';
    var d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, AppConfig.TIME_ZONE, 'd MMM yyyy');
  }

  /** Builds the factual context for one task from workbook rows. Pure given rows. */
  function buildContext(input) {
    var grades = input.grades || [];
    var current = {};
    (input.submissions || []).forEach(function (s) {
      if (text(s.CurrentOfficial) === '' || s.CurrentOfficial === true || text(s.CurrentOfficial).toLowerCase() === 'true') current[text(s.SubmissionRecordID)] = true;
    });
    var dist = { A: 0, B: 0, C: 0, D: 0, E: 0 };
    var counted = 0;
    grades.forEach(function (g) {
      if (!current[text(g.SubmissionRecordID)]) return;
      var letter = text(g.Grade).toUpperCase();
      if (dist[letter] !== undefined) { dist[letter]++; counted++; }
    });
    var outcomes = [];
    ((input.profile && input.profile.criteria) || []).forEach(function (c) {
      var code = text(c.outcome).trim();
      if (code && outcomes.indexOf(code) === -1) outcomes.push(code);
    });
    var approvedTimes = grades.map(function (g) { return text(g.ApprovedAt); }).filter(Boolean).sort();
    var link = input.link || {};
    var parsed = parseWorkbookName(input.workbookName);
    var dueText = input.formatDate ? input.formatDate(link.DueDate) : text(link.DueDate);
    var dates = [];
    if (dueText) dates.push('Due ' + dueText);
    if (approvedTimes.length) dates.push('Marked ' + (input.formatDate ? input.formatDate(approvedTimes[0]) : approvedTimes[0]) +
      (approvedTimes.length > 1 ? ' – ' + (input.formatDate ? input.formatDate(approvedTimes[approvedTimes.length - 1]) : approvedTimes[approvedTimes.length - 1]) : ''));
    return {
      taskName: input.taskName,
      courseCode: parsed.courseCode,
      year: parsed.year,
      className: text(link.CourseName) || text((input.submissions || [])[0] && input.submissions[0].Class),
      assignmentTitle: text(link.AssignmentTitle),
      dates: dates.join('; '),
      outcomes: outcomes.join(', '),
      gradeDistribution: counted ? ['A', 'B', 'C', 'D', 'E'].map(function (l) { return l + ': ' + dist[l]; }).join(', ') + ' (' + counted + ' approved)' : 'No approved results yet',
      approvedCount: counted
    };
  }

  function rows(ss, name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    var h = data[0].map(function (x) { return text(x).trim(); });
    return data.slice(1).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = r[i] instanceof Date ? r[i].toISOString() : r[i]; }); return o; });
  }

  function context(workbookId, taskName) {
    var ss = LessonGraderWeb.open(workbookId);
    var task = text(taskName).trim();
    if (LessonGraderWeb.taskNames(ss).indexOf(task) === -1) throw new Error('Choose a task tab in this workbook.');
    var links = rows(ss, 'ClassroomConfig').filter(function (r) { return text(r.TaskName) === task; });
    return buildContext({
      taskName: task, workbookName: ss.getName(), link: links[links.length - 1],
      submissions: rows(ss, 'Submissions').filter(function (r) { return text(r.Task).trim() === task; }),
      grades: rows(ss, 'ApprovedGrades').filter(function (r) { return text(r.TaskName).trim() === task; }),
      profile: RubricEngine.getActiveProfile(ss, task), formatDate: sydneyDate
    });
  }

  function describeDestination(docId) {
    var id = text(docId).trim();
    if (!id) return { type: 'sheet', id: '', name: 'EvaluationRegister sheet in this workbook' };
    if (!DOC_ID_RE.test(id)) throw new Error('That does not look like a Google Doc ID.');
    var file = DriveApp.getFileById(id);
    if (file.isTrashed() || file.getMimeType() !== MimeType.GOOGLE_DOCS) throw new Error('The register destination must be a Google Doc you can edit.');
    return { type: 'doc', id: id, name: file.getName(), url: file.getUrl() };
  }

  function preview(workbookId, taskName, docId) {
    WebAccess.requireTeacher();
    return { success: true, data: { context: context(workbookId, taskName), destination: describeDestination(docId) } };
  }

  function cleanEntry(entry) {
    var out = {};
    Object.keys(TEXT_FIELDS).forEach(function (k) {
      out[k] = text(entry && entry[k]).trim();
      if (out[k].length > TEXT_FIELDS[k]) throw new Error('Register field "' + k + '" is too long.');
    });
    if (!out.teacherEvaluation) throw new Error('Write or dictate the teacher evaluation before committing.');
    return out;
  }

  function ensureSheet(ss) {
    var sheet = ss.getSheetByName(SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  function appendDocSection(docId, ctx, entry, user) {
    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    body.appendParagraph(ctx.taskName + (ctx.className ? ' — ' + ctx.className : '')).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    [['Course', ctx.courseCode + (ctx.year ? ' (' + ctx.year + ')' : '')], ['Dates', ctx.dates], ['Outcomes', ctx.outcomes],
     ['Grade distribution', ctx.gradeDistribution], ['Teacher evaluation', entry.teacherEvaluation],
     ['Teaching variations', entry.teachingVariations], ['Adjustments', entry.adjustments],
     ['Student feedback summary', entry.studentFeedbackSummary],
     ['Committed', sydneyDate(new Date()) + ' by ' + user]].forEach(function (pair) {
      if (!pair[1]) return;
      var p = body.appendParagraph(pair[0] + ': ' + pair[1]);
      p.editAsText().setBold(0, pair[0].length, true);
    });
    doc.saveAndClose();
  }

  /** Commits a teacher-reviewed entry. Requires confirmed === true. */
  function commit(workbookId, taskName, entry, docId, confirmed) {
    var user = WebAccess.requireTeacher();
    if (confirmed !== true) throw new Error('Preview the entry and confirm the destination before committing.');
    var ss = LessonGraderWeb.open(workbookId);
    var ctx = context(workbookId, taskName);
    var clean = cleanEntry(entry);
    var dest = describeDestination(docId);
    var status = 'Recorded in workbook';
    var docError = '';
    if (dest.type === 'doc') {
      try { appendDocSection(dest.id, ctx, clean, user); status = 'Written to Google Doc and workbook'; }
      catch (e) { docError = e && e.message ? e.message : 'Document could not be updated.'; status = 'Doc write failed — recorded in workbook only'; }
    }
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sheet = ensureSheet(ss);
      var values = { EntryID: Utilities.getUuid(), TaskName: ctx.taskName, CourseCode: ctx.courseCode, ClassName: ctx.className,
        Year: ctx.year, Dates: ctx.dates, Outcomes: ctx.outcomes, GradeDistribution: ctx.gradeDistribution,
        TeacherEvaluation: clean.teacherEvaluation, TeachingVariations: clean.teachingVariations, Adjustments: clean.adjustments,
        StudentFeedbackSummary: clean.studentFeedbackSummary, Destination: dest.type, DestinationID: dest.id,
        DestinationName: dest.name, CommittedAt: new Date().toISOString(), CommittedBy: user, Status: status };
      var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (x) { return text(x).trim(); });
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, h.length).setValues([h.map(function (k) { return values[k] !== undefined ? values[k] : ''; })]);
    } finally { lock.releaseLock(); }
    if (docError) return { success: false, message: 'The Google Doc was NOT updated (' + docError + '). The entry was recorded in the workbook EvaluationRegister sheet.', data: { destination: dest, status: status } };
    return { success: true, message: status + ': ' + dest.name + '.', data: { destination: dest, status: status } };
  }

  return { HEADERS: HEADERS, parseWorkbookName: parseWorkbookName, buildContext: buildContext,
    context: context, preview: preview, commit: commit, describeDestination: describeDestination };
})();

function apiWebRegisterPreview(workbookId, taskName, docId) { return EvaluationRegister.preview(workbookId, taskName, docId); }
function apiWebRegisterCommit(workbookId, taskName, entry, docId, confirmed) { return EvaluationRegister.commit(workbookId, taskName, entry, docId, confirmed); }
