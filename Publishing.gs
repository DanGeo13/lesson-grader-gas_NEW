/* ============================================================================
 * Publishing.gs — explicit, selective publication of APPROVED results to
 * Google Classroom. Never called by approval, sync or AI. (SC39, SC40, SC35)
 *
 * Flow: preview(selected) → teacher checks students, grades and destination →
 * publish(selected, previewToken, 'PUBLISH'). The token binds the confirmation
 * to exactly what was previewed; any change in between is refused.
 *
 * What Classroom accepts (REVIEW.md §4 Q7, A4):
 *   - a NUMBER in draftGrade/assignedGrade (no letters). The letter and the
 *     comment are produced as text for the teacher to paste: the Classroom API
 *     cannot write private comments.
 *   - grades can only be patched on coursework created by this Apps Script's
 *     Google Cloud project; teacher-created assignments return
 *     ProjectPermissionDenied. This is reported per student, never hidden.
 *   - Grades stay as drafts invisible to students unless "return to student"
 *     is explicitly chosen.
 * ============================================================================ */
var Publishing = (function () {
  var LOG = 'PublishLog';
  var LOG_HEADERS = ['PublishID', 'TaskName', 'SubmissionRecordID', 'StudentName', 'CourseID', 'CourseWorkID', 'ClassroomSubmissionID',
    'Letter', 'Grade', 'MaxPoints', 'Mode', 'Returned', 'Status', 'Message', 'PublishedAt', 'PublishedBy'];

  function text(v) { return v == null ? '' : String(v); }
  function yes(v) { return v === true || text(v).toLowerCase() === 'true'; }

  function destination(ss, task) {
    var links = WebGrading.records(ss, 'ClassroomConfig').filter(function (r) { return text(r.TaskName) === task && yes(r.Active) && r.CourseID && r.CourseWorkID; });
    if (!links.length) throw new Error('This task is not linked to a Classroom assignment.');
    var link = links[links.length - 1];
    var dest = { courseId: text(link.CourseID), courseName: text(link.CourseName), courseWorkId: text(link.CourseWorkID),
      assignmentTitle: text(link.AssignmentTitle), maxPoints: Number(link.MaxPoints) || 0, live: false };
    try {
      var cw = Classroom.Courses.CourseWork.get(dest.courseId, dest.courseWorkId);
      dest.assignmentTitle = text(cw.title) || dest.assignmentTitle;
      dest.maxPoints = cw.maxPoints == null ? 0 : Number(cw.maxPoints);
      dest.live = true;
    } catch (e) { dest.warning = 'Could not re-read the assignment from Classroom; showing the saved link details.'; }
    return dest;
  }

  function build(workbookId, taskName, submissionIds) {
    var target = WebGrading.openTask(workbookId, taskName);
    var ss = target.ss;
    var ids = (Array.isArray(submissionIds) ? submissionIds : []).map(text).filter(Boolean);
    if (!ids.length) throw new Error('Select at least one approved student to publish.');
    if (ids.length > 200) throw new Error('Publish at most 200 students at a time.');
    var dest = destination(ss, target.task);
    var settings = AppConfig.readSettings();
    var mode = settings.outputFormat.classroomGradeMode;
    var subs = WebGrading.records(ss, 'Submissions');
    var grades = WebGrading.records(ss, 'ApprovedGrades');
    var eligible = [], ineligible = [];
    ids.forEach(function (id) {
      var s = subs.filter(function (r) { return text(r.SubmissionRecordID) === id && text(r.Task).trim() === target.task; })[0];
      if (!s) { ineligible.push({ id: id, name: '', reason: 'Not in this task.' }); return; }
      var name = text(s.StudentName);
      if (['Approved', 'Locked'].indexOf(text(s.Status)) === -1) { ineligible.push({ id: id, name: name, reason: 'Not approved.' }); return; }
      if (!WebGrading.isCurrent(s)) { ineligible.push({ id: id, name: name, reason: 'A newer version exists.' }); return; }
      if (!text(s.ClassroomSubmissionID) || text(s.ClassroomCourseID) !== dest.courseId || text(s.ClassroomCourseWorkID) !== dest.courseWorkId) {
        ineligible.push({ id: id, name: name, reason: 'No matching Classroom submission (manual or different assignment).' }); return;
      }
      var g = grades.filter(function (r) { return text(r.SubmissionRecordID) === id; }).pop();
      if (!g) { ineligible.push({ id: id, name: name, reason: 'No approval snapshot.' }); return; }
      var points = Number(g.FinalScore !== '' ? g.FinalScore : g.Points);
      var mapped;
      try { mapped = RubricEngine.classroomGrade(points, Number(g.MaxMarks), dest.maxPoints, mode); }
      catch (e) { ineligible.push({ id: id, name: name, reason: e.message }); return; }
      var fb = {};
      try { fb = JSON.parse(text(g.FeedbackJSON) || '{}'); } catch (e2) { fb = {}; }
      var comment = FeedbackBuilder.render(settings.feedbackTemplate, fb,
        { studentName: name, taskName: target.task, letter: text(g.Grade), points: points, maxMarks: g.MaxMarks, percent: Number(g.Percent) },
        settings.outputFormat);
      eligible.push({ id: id, name: name, letter: text(g.Grade), points: points, maxMarks: Number(g.MaxMarks), grade: mapped.grade, note: mapped.note,
        classroomSubmissionId: text(s.ClassroomSubmissionID), alreadyPublished: text(s.LastSyncResult), comment: comment.text, commentWarnings: comment.warnings, row: s._row });
    });
    var token = AssessmentFlags.hash({ w: workbookId, t: target.task, c: dest.courseWorkId, m: dest.maxPoints, mode: mode,
      e: eligible.map(function (e) { return [e.id, e.grade]; }) });
    return { ss: ss, task: target.task, dest: dest, mode: mode, eligible: eligible, ineligible: ineligible, token: token };
  }

  function strip(p) {
    return { destination: p.dest, mode: p.mode, previewToken: p.token, ineligible: p.ineligible,
      eligible: p.eligible.map(function (e) { var c = Object.assign({}, e); delete c.row; return c; }) };
  }

  function preview(workbookId, taskName, submissionIds) {
    WebAccess.requireTeacher();
    var p = build(workbookId, taskName, submissionIds);
    return { success: true, data: strip(p),
      message: p.eligible.length + ' student(s) ready to publish to "' + p.dest.assignmentTitle + '" (' + p.dest.courseName + '). Nothing has been sent yet.' };
  }

  function explain(e) {
    var m = text(e && e.message);
    if (/ProjectPermissionDenied|@ProjectPermissionDenied|Developer Console project/i.test(m)) {
      return 'Classroom only lets an app change grades on assignments that app created. Enter this grade in Classroom manually.';
    }
    if (/permission|forbidden|403/i.test(m)) return 'Classroom refused access for your account.';
    return m.slice(0, 200) || 'Classroom rejected the update.';
  }

  /**
   * Publishes exactly the previewed students. Per-student isolation; honest results.
   * @param {string} previewToken token returned by preview
   * @param {string} confirmation must be the word PUBLISH
   * @param {{returnToStudents?:boolean}} options
   */
  function publish(workbookId, taskName, submissionIds, previewToken, confirmation, options) {
    var user = WebAccess.requireTeacher();
    if (text(confirmation).trim() !== 'PUBLISH') throw new Error('Type PUBLISH to confirm.');
    var p = build(workbookId, taskName, submissionIds);
    if (text(previewToken) !== p.token) throw new Error('Something changed since the preview (grades, selection or assignment). Preview again before publishing.');
    if (!p.eligible.length) throw new Error('None of the selected students can be published.');
    var giveBack = !!(options && options.returnToStudents === true);
    var ss = p.ss;
    WebGrading.ensureSheet(ss, LOG, LOG_HEADERS);
    var results = [];
    p.eligible.forEach(function (e) {
      var r = { id: e.id, name: e.name, grade: e.grade, published: false, returned: false, message: '' };
      try {
        if (e.grade !== null) {
          Classroom.Courses.CourseWork.StudentSubmissions.patch({ draftGrade: e.grade, assignedGrade: e.grade },
            p.dest.courseId, p.dest.courseWorkId, e.classroomSubmissionId, { updateMask: 'draftGrade,assignedGrade' });
        }
        r.published = true;
        if (giveBack) {
          try {
            var api = Classroom.Courses.CourseWork.StudentSubmissions;
            var giveBackFn = api['return'] || api['return_']; // reserved word: name differs between runtimes
            giveBackFn.call(api, {}, p.dest.courseId, p.dest.courseWorkId, e.classroomSubmissionId);
            r.returned = true;
          }
          catch (re) { r.message = 'Grade saved but not returned: ' + explain(re); }
        }
        r.message = r.message || (e.grade === null ? 'No grade sent (output setting).' : 'Grade ' + e.grade + ' / ' + p.dest.maxPoints + (r.returned ? ', returned to student.' : ' saved as a draft grade (not yet returned).'));
      } catch (err) {
        r.message = explain(err);
      }
      var stamp = new Date().toISOString();
      WebGrading.append(ss, LOG, { PublishID: Utilities.getUuid(), TaskName: p.task, SubmissionRecordID: e.id, StudentName: e.name,
        CourseID: p.dest.courseId, CourseWorkID: p.dest.courseWorkId, ClassroomSubmissionID: e.classroomSubmissionId,
        Letter: e.letter, Grade: e.grade === null ? '' : e.grade, MaxPoints: p.dest.maxPoints, Mode: p.mode, Returned: r.returned,
        Status: r.published ? 'Published' : 'Failed', Message: r.message, PublishedAt: stamp, PublishedBy: user });
      WebGrading.patch(ss, 'Submissions', e.row, r.published ?
        { ClassroomAssignedGrade: e.grade === null ? '' : e.grade, LastClassroomSyncAt: stamp, LastSyncResult: 'Published ' + (e.grade === null ? '(no grade)' : e.grade + '/' + p.dest.maxPoints) + (r.returned ? ' · returned' : ' · draft') } :
        { LastClassroomSyncAt: stamp, LastSyncResult: 'Publish failed: ' + r.message.slice(0, 120) });
      WebGrading.event(ss, e.id, '', r.published ? 'PublishedToClassroom' : 'PublishFailed', {}, { grade: e.grade, maxPoints: p.dest.maxPoints, returned: r.returned }, user, r.message);
      results.push(r);
    });
    var ok = results.filter(function (r) { return r.published; }).length;
    return { success: ok === results.length, partial: ok > 0 && ok < results.length, data: { results: results },
      message: 'Published ' + ok + ' of ' + results.length + ' student(s) to "' + p.dest.assignmentTitle + '".' + (ok < results.length ? ' See the list for failures.' : '') };
  }

  /** Student-facing comment and Classroom mapping preview for one approved or draft assessment. */
  function outputPreview(workbookId, taskName, submissionId) {
    WebAccess.requireTeacher();
    var d = WebGrading.detail(workbookId, taskName, submissionId).data;
    if (!d.profile || !d.scores) throw new Error('No rubric is available for this assessment.');
    var settings = AppConfig.readSettings();
    var src = d.snapshot || { points: d.scores.totalPoints, maxMarks: d.scores.totalMaxMarks, percent: d.scores.percent, grade: d.scores.letter };
    var comment = FeedbackBuilder.render(settings.feedbackTemplate, d.feedback, { studentName: d.submission.studentName, taskName: d.submission.taskName,
      letter: src.grade, points: src.points, maxMarks: src.maxMarks, percent: src.percent }, settings.outputFormat);
    var mapping = null;
    try {
      var target = WebGrading.openTask(workbookId, taskName);
      var dest = destination(target.ss, target.task);
      mapping = { destination: dest, mode: settings.outputFormat.classroomGradeMode,
        grade: RubricEngine.classroomGrade(src.points, src.maxMarks, dest.maxPoints, settings.outputFormat.classroomGradeMode) };
    } catch (e) { mapping = { error: e.message }; }
    return { success: true, data: { comment: comment, mapping: mapping, basedOn: d.snapshot ? 'approved snapshot' : 'current draft (not approved)' } };
  }

  return { LOG_HEADERS: LOG_HEADERS, preview: preview, publish: publish, outputPreview: outputPreview, explain: explain };
})();

function apiWebPublishPreview(workbookId, taskName, submissionIds) { return Publishing.preview(workbookId, taskName, submissionIds); }
function apiWebPublish(workbookId, taskName, submissionIds, previewToken, confirmation, options) {
  return Publishing.publish(workbookId, taskName, submissionIds, previewToken, confirmation, options);
}
function apiWebOutputPreview(workbookId, taskName, submissionId) { return Publishing.outputPreview(workbookId, taskName, submissionId); }
