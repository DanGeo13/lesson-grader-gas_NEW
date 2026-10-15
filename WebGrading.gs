/* Workbook-scoped web grading: assessment state, versioning and the teacher
 * review lifecycle. No task-specific criteria, sheet names or AI prompt.
 *
 * States (Submissions.Status): New → InReview (draft) → Approved → Locked.
 *   - Approve never publishes to Classroom (Publishing.gs does that, separately).
 *   - Locked is final: no reassessment until an authorised unlock.
 *   - Reassessment creates a NEW version row; the earlier approval is kept.
 * Every public endpoint requires an allowlisted teacher account, and every
 * read/write takes an explicit, library-validated workbook ID. */
var WebGrading = (function () {
  var SUB = ['SubmissionRecordID','ClassroomCourseID','ClassroomCourseWorkID','ClassroomSubmissionID','StudentUserID','StudentName','StudentEmail','Class','Task','SubmissionVersion','SourceType','ClassroomState','TurnedInTime','UpdateTime','Late','AttachmentSummary','AttachmentFileIDsJSON','AttachmentMetadataJSON','DriveFolderID','Status','ParentSubmissionRecordID','CurrentOfficial','LockedAt','LockedBy','ApprovedAt','ApprovedBy','ClassroomAssignedGrade','LastClassroomSyncAt','LastSyncResult','Notes','AssessmentType','RawScore','DeductionPoints','DeductionReason','FinalScore'];
  var FILES = ['SubmissionRecordID','FileRecordID','SourceType','DriveFileID','FileName','MimeType','AlternateLink','ThumbnailUrl','FileSize','EligibleForAI','AIReviewStatus','AIExtractedText','Limitations','CreatedAt'];
  var ASSESS = ['AssessmentID','SubmissionRecordID','CriterionID','AIProposedGrade','DerivedGrade','FinalApprovedGrade','DistinctOverrideSelected','CheckboxesJSON','AIEvidenceJSON','TeacherAdjusted','TeacherAudioTranscript','TeacherWrittenNote','Confidence','MinimumEvidenceIncomplete','Status','CreatedAt','UpdatedAt','ApprovedAt','ApprovedBy','RubricProfileID'];
  var HISTORY = ['HistoryID','SubmissionRecordID','AssessmentID','Action','CriterionID','PreviousStateJSON','NewStateJSON','ActorEmail','Timestamp','Notes'];
  var AI = ['AssessmentID','SubmissionRecordID','RunType','RunTimestamp','ModelUsed','PromptVersion','InputFilesJSON','PromptSummary','ResponseJSON','ExecutionStatus','ErrorMessage','LatencySeconds','TriggeredBy','TeacherOutcome'];
  var GRADES = ['GradeID','TaskName','SubmissionRecordID','StudentUserID','StudentName','AssessmentID','RubricProfileID','Points','MaxMarks','Percent','Grade','FeedbackJSON','ApprovedAt','ApprovedBy','AssessmentType','RawScore','DeductionPoints','DeductionReason','FinalScore','OutcomesJSON'];
  var AUDIT = ['AuditID','TaskName','Action','TargetID','TargetURL','ActorEmail','Timestamp','Status','DetailsJSON'];
  var TRACKING = ['TrackingID','Year','Stage','CourseCode','ClassID','TermPair','StudentUserID','StudentName','TaskName','AssessmentType','RawScore','FinalScore','MaxMarks','Percent','Grade','OutcomesJSON','Status','UpdatedAt'];
  var MESSAGES = ['MessageID','CourseID','CourseWorkID','StudentUserID','StudentName','StudentEmail','Reason','MessageText','Status','SentAt','SentBy'];
  /* Core tables: "Enable grading" (with a Drive backup) creates these. */
  var HEADERS = { Submissions: SUB, SubmissionFiles: FILES, CriterionAssessments: ASSESS, AssessmentHistory: HISTORY, AIAssessments: AI, ApprovedGrades: GRADES, AuditLog: AUDIT, ClassTracking: TRACKING, ClassroomMessages: MESSAGES };
  /* Additive columns introduced by this review. Added on first write, never
     reordering or rewriting existing data (REVIEW.md §7 migration notes). */
  var EXTRA = {
    Submissions: ['EvidenceChangedAt', 'PendingEvidenceJSON', 'UnlockedAt', 'UnlockedBy', 'UnlockReason'],
    CriterionAssessments: ['TeacherReviewJSON', 'AIProposedChecksJSON'],
    SubmissionFiles: ['EvidenceStatusUpdatedAt']
  };
  var FINAL = ['Approved', 'Locked', 'Superseded'];
  /* Human-editable mirror of each task's ACTIVE rubric (one row per observable).
     Teachers may edit it in Sheets; the app shows the difference and activates
     the edited rows as a new version only when the teacher confirms. (SC14) */
  var RUBRIC_ROWS = ['TaskName', 'Criterion', 'Part', 'Section', 'MaxMarks', 'Outcome', 'Band', 'Description', 'BandWeight', 'ProfileID', 'SyncedAt'];

  function fail(message) { return { success: false, message: message }; }
  function uid() { return Utilities.getUuid(); }
  function now() { return new Date().toISOString(); }
  function text(v) { return v == null ? '' : String(v); }
  function value(v) { return v instanceof Date ? v.toISOString() : v; }
  function yes(v) { return v === true || String(v).toLowerCase() === 'true'; }
  function parseJson(v, fallback) { try { var p = JSON.parse(text(v)); return p == null ? fallback : p; } catch (e) { return fallback; } }
  function errString(e) { return e && e.message ? e.message : 'Operation failed.'; }
  function isCurrent(row) { return row.CurrentOfficial === '' || row.CurrentOfficial == null || yes(row.CurrentOfficial); }
  function safeLink(url) { return /^https:\/\/(drive|docs)\.google\.com\//.test(text(url)) ? text(url) : ''; }
  function toNumber(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = parseFloat(String(v == null ? '' : v).replace(/[^\d.\-]/g, ''));
    return isFinite(n) ? n : 0;
  }

  function columnMap(sheet) {
    if (!sheet || sheet.getLastRow() < 1) throw new Error('Workbook is not prepared for web grading.');
    var names = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var map = Object.create(null);
    for (var c = 0; c < names.length; c++) map[text(names[c]).trim()] = c;
    return { names: names, map: map };
  }

  function records(ss, name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    var names = data[0];
    var rows = [];
    for (var r = 1; r < data.length; r++) {
      var item = { _row: r + 1 };
      for (var c = 0; c < names.length; c++) item[text(names[c]).trim()] = value(data[r][c]);
      rows.push(item);
    }
    return rows;
  }

  function append(ss, sheetName, fields) {
    var sheet = ss.getSheetByName(sheetName);
    var h = columnMap(sheet);
    var row = h.names.map(function (name) {
      return Object.prototype.hasOwnProperty.call(fields, name) ? fields[name] : '';
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    return sheet.getLastRow();
  }

  function patch(ss, sheetName, rowNumber, fields) {
    var sheet = ss.getSheetByName(sheetName);
    var h = columnMap(sheet);
    Object.keys(fields).forEach(function (name) {
      if (h.map[name] === undefined) throw new Error('Workbook needs column ' + name + '. Use Enable grading first.');
      sheet.getRange(rowNumber, h.map[name] + 1).setValue(fields[name]);
    });
  }

  function ensureSheet(ss, name, required) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (!sheet.getLastRow()) {
      sheet.getRange(1, 1, 1, required.length).setValues([required]);
      sheet.getRange(1, 1, 1, required.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      return sheet;
    }
    var h = columnMap(sheet);
    var missing = required.filter(function (name) { return h.map[name] === undefined; });
    if (!missing.length) return sheet;
    var at = h.names.length + 1;
    var last = at + missing.length - 1;
    if (sheet.getMaxColumns() < last) sheet.insertColumnsAfter(sheet.getMaxColumns(), last - sheet.getMaxColumns());
    sheet.getRange(1, at, 1, missing.length).setValues([missing]);
    return sheet;
  }

  /* Additive migration for columns added by this review. Safe to call often. */
  function upgrade(ss) {
    Object.keys(EXTRA).forEach(function (name) {
      if (ss.getSheetByName(name)) ensureSheet(ss, name, HEADERS[name].concat(EXTRA[name]));
    });
  }

  function prepare(ss, takeBackup) {
    /* Existing workbooks: take a Drive copy BEFORE adding any headers. Never
       continue if backup failed; user explicitly chooses this step in the UI. */
    var backupId = '';
    if (takeBackup) {
      var copy = DriveApp.getFileById(ss.getId()).makeCopy('[Backup ' + now().replace(/[:.]/g, '-') + '] ' + ss.getName());
      backupId = copy.getId();
    }
    Object.keys(HEADERS).forEach(function (key) { ensureSheet(ss, key, HEADERS[key].concat(EXTRA[key] || [])); });
    RubricEngine.ensureProfilesSheet(ss);
    SpreadsheetApp.flush();
    return { success: true, message: takeBackup ? 'Backup created; web grading enabled. Existing assessment rows were not changed.' : 'Web grading tables ready.', backupId: backupId };
  }

  function openTask(workbookId, taskName) {
    var ss = LessonGraderWeb.open(workbookId); // validates Drive library membership
    var task = text(taskName).trim();
    if (!task || LessonGraderWeb.taskNames(ss).indexOf(task) === -1) throw new Error('Choose a task tab in this workbook.');
    return { ss: ss, task: task };
  }

  function ready(ss) {
    return Object.keys(HEADERS).every(function (name) {
      var sh = ss.getSheetByName(name);
      if (!sh) return false;
      var h = columnMap(sh);
      return HEADERS[name].every(function (col) { return h.map[col] !== undefined; });
    }) && !!ss.getSheetByName('RubricProfiles');
  }

  function selectSubmission(ss, task, submissionId) {
    var rows = records(ss, 'Submissions');
    for (var i = 0; i < rows.length; i++) {
      if (text(rows[i].SubmissionRecordID) === text(submissionId) && text(rows[i].Task).trim() === task) return rows[i];
    }
    throw new Error('Submission not found for this task.');
  }

  function lastRevision(ss, submissionId) {
    var rows = records(ss, 'AssessmentHistory');
    for (var i = rows.length - 1; i >= 0; i--) {
      if (text(rows[i].SubmissionRecordID) === text(submissionId)) return text(rows[i].HistoryID);
    }
    return '';
  }

  function event(ss, subId, assessmentId, action, before, after, user, note) {
    var id = uid();
    append(ss, 'AssessmentHistory', {
      HistoryID: id, SubmissionRecordID: subId, AssessmentID: assessmentId,
      Action: action, CriterionID: '', PreviousStateJSON: JSON.stringify(before || {}),
      NewStateJSON: JSON.stringify(after || {}), ActorEmail: user, Timestamp: now(), Notes: note || ''
    });
    return id;
  }

  function editable(sub) {
    if (text(sub.Status) === 'Locked') throw new Error('This assessment is final-locked. An authorised teacher must unlock it before any change.');
    if (FINAL.indexOf(text(sub.Status)) !== -1) throw new Error('This assessment is final. Create a new reassessment version; approved work cannot be edited.');
    if (!isCurrent(sub)) throw new Error('This is a historical version. Open the current submission to grade.');
  }

  function withLock(fn) {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { return fn(); } finally { lock.releaseLock(); }
  }

  /* ---- reading an assessment ------------------------------------------- */

  function latestAssessmentRows(all) {
    var assessmentId = all.length ? text(all[all.length - 1].AssessmentID) : '';
    return { assessmentId: assessmentId, rows: all.filter(function (r) { return text(r.AssessmentID) === assessmentId; }) };
  }

  function statesFromRows(rows) {
    var states = {};
    rows.forEach(function (r) {
      states[text(r.CriterionID)] = {
        checkboxes: parseJson(r.CheckboxesJSON, {}),
        teacherWrittenNote: text(r.TeacherWrittenNote), teacherAudioTranscript: text(r.TeacherAudioTranscript),
        aiProposedGrade: text(r.AIProposedGrade), aiEvidence: parseJson(r.AIEvidenceJSON, []),
        aiProposal: parseJson(r.AIProposedChecksJSON, null),
        review: parseJson(r.TeacherReviewJSON, {}),
        status: text(r.Status)
      };
    });
    return states;
  }

  function latestFeedback(ss, submissionId, assessmentId) {
    var rows = records(ss, 'AIAssessments').filter(function (r) {
      return text(r.SubmissionRecordID) === text(submissionId) && text(r.AssessmentID) === assessmentId && text(r.RunType) === 'TeacherDraft';
    });
    return rows.length ? parseJson(rows[rows.length - 1].ResponseJSON, {}) : {};
  }

  function attachmentsFor(ss, submissionId) {
    return records(ss, 'SubmissionFiles').filter(function (r) {
      return text(r.SubmissionRecordID) === text(submissionId) && text(r.AIReviewStatus) !== 'Removed';
    }).map(function (r) {
      return { fileRecordId: text(r.FileRecordID), fileName: text(r.FileName) || 'Untitled file', driveFileId: text(r.DriveFileID),
        link: safeLink(r.AlternateLink), mimeType: text(r.MimeType), eligible: yes(r.EligibleForAI),
        status: text(r.AIReviewStatus) || 'Pending', limitations: text(r.Limitations) };
    });
  }

  function pinnedProfile(ss, task, current) {
    var pinnedIds = current.map(function (r) { return text(r.RubricProfileID); }).filter(Boolean);
    var legacy = current.length > 0 && (pinnedIds.length !== current.length || pinnedIds.some(function (id) { return id !== pinnedIds[0]; }));
    var profile = !legacy && pinnedIds.length ? RubricEngine.getProfileById(ss, pinnedIds[0]) : RubricEngine.getActiveProfile(ss, task);
    if (pinnedIds.length && !profile) legacy = true;
    return { profile: legacy ? null : profile, legacy: legacy };
  }

  function snapshotFor(ss, submissionId) {
    var rows = records(ss, 'ApprovedGrades').filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId); });
    if (!rows.length) return null;
    var g = rows[rows.length - 1];
    return { points: toNumber(g.FinalScore !== '' ? g.FinalScore : g.Points), maxMarks: toNumber(g.MaxMarks), percent: toNumber(g.Percent),
      grade: text(g.Grade), approvedAt: text(g.ApprovedAt), approvedBy: text(g.ApprovedBy), rubricProfileId: text(g.RubricProfileID) };
  }

  /* Changes between an approved earlier version and this version (SC42). */
  function compareWithParent(ss, sub, profile, states) {
    var parentId = text(sub.ParentSubmissionRecordID);
    if (!parentId) return null;
    var parentRows = latestAssessmentRows(records(ss, 'CriterionAssessments').filter(function (r) { return text(r.SubmissionRecordID) === parentId; })).rows;
    var parentStates = statesFromRows(parentRows);
    var parentProfileId = parentRows.length ? text(parentRows[0].RubricProfileID) : '';
    var parentProfile = parentProfileId ? RubricEngine.getProfileById(ss, parentProfileId) : null;
    var before = parentProfile ? RubricEngine.scoreProfile(parentProfile, parentStates) : null;
    var after = profile ? RubricEngine.scoreProfile(profile, states) : null;
    var criteria = (profile ? profile.criteria : []).map(function (c) {
      var was = (parentStates[c.criterionId] || {}).checkboxes || {};
      var now_ = (states[c.criterionId] || {}).checkboxes || {};
      var wasT = AssessmentFlags.ticked(was), nowT = AssessmentFlags.ticked(now_);
      return {
        criterionId: c.criterionId, title: c.title,
        before: before && before.criteria[c.criterionId] ? { points: before.criteria[c.criterionId].points, grade: before.criteria[c.criterionId].grade } : null,
        after: after ? { points: after.criteria[c.criterionId].points, grade: after.criteria[c.criterionId].grade } : null,
        added: nowT.filter(function (id) { return wasT.indexOf(id) === -1; }),
        removed: wasT.filter(function (id) { return nowT.indexOf(id) === -1; })
      };
    });
    return { parentId: parentId, sameRubric: parentProfileId === (profile && profile.profileId), approved: snapshotFor(ss, parentId),
      before: before ? { totalPoints: before.totalPoints, percent: before.percent, letter: before.letter } : null,
      after: after ? { totalPoints: after.totalPoints, percent: after.percent, letter: after.letter } : null, criteria: criteria };
  }

  function assess(ss, task, sub) {
    var submissionId = text(sub.SubmissionRecordID);
    var all = records(ss, 'CriterionAssessments').filter(function (r) { return text(r.SubmissionRecordID) === submissionId; });
    var latest = latestAssessmentRows(all);
    var pin = pinnedProfile(ss, task, latest.rows);
    var states = statesFromRows(latest.rows);
    var feedback = latestFeedback(ss, submissionId, latest.assessmentId);
    var attachments = attachmentsFor(ss, submissionId);
    var scores = pin.profile ? RubricEngine.scoreProfile(pin.profile, states) : null;
    var flags = pin.profile ? AssessmentFlags.compute({ profile: pin.profile, states: states, scores: scores, attachments: attachments, feedback: feedback,
      evidenceChangedAt: text(sub.EvidenceChangedAt), final: FINAL.indexOf(text(sub.Status)) !== -1 }) :
      { flags: [], blocking: 0, advisory: 0, byCriterion: {} };
    return { assessmentId: latest.assessmentId, rows: latest.rows, profile: pin.profile, legacy: pin.legacy, states: states,
      feedback: feedback, attachments: attachments, scores: scores, flags: flags };
  }

  function listTask(workbookId, taskName) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var profile = RubricEngine.getActiveProfile(ss, target.task);
    var all = records(ss, 'Submissions').filter(function (row) { return text(row.Task).trim() === target.task && !!row.SubmissionRecordID; });
    var assessRows = records(ss, 'CriterionAssessments');
    var submissions = all.map(function (r) {
      var id = text(r.SubmissionRecordID);
      var hasDraft = assessRows.some(function (a) { return text(a.SubmissionRecordID) === id; });
      var reassessment = !!text(r.ParentSubmissionRecordID);
      return {
        submissionRecordId: id, studentName: text(r.StudentName), studentUserId: text(r.StudentUserID),
        studentEmail: text(r.StudentEmail), className: text(r.Class), version: Number(r.SubmissionVersion) || 1,
        status: text(r.Status) || 'New', stateLabel: AppConfig.stateLabel(r.Status, reassessment, r.SubmissionVersion),
        currentOfficial: isCurrent(r), reassessment: reassessment, hasDraft: hasDraft,
        updatedAt: text(r.UpdateTime), late: yes(r.Late), sourceType: text(r.SourceType),
        evidenceChangedAt: text(r.EvidenceChangedAt), published: text(r.LastSyncResult),
        publishedAt: text(r.LastClassroomSyncAt), hasClassroomIds: !!(text(r.ClassroomCourseID) && text(r.ClassroomSubmissionID)),
        finalScore: r.FinalScore === '' || r.FinalScore == null ? null : toNumber(r.FinalScore)
      };
    });
    submissions.sort(function (a, b) { return a.studentName.localeCompare(b.studentName) || (b.version - a.version); });
    var current = submissions.filter(function (s) { return s.currentOfficial; });
    var graded = current.filter(function (s) { return s.status === 'Approved' || s.status === 'Locked'; });
    var profileEntries = RubricEngine.readProfiles(ss).filter(function (p) { return text(p.row.TaskName) === target.task; });
    var links = records(ss, 'ClassroomConfig').filter(function (r) { return text(r.TaskName) === target.task && yes(r.Active); });
    var link = links[links.length - 1];
    return { success: true, data: {
      workbook: { id: workbookId, name: ss.getName() }, taskName: target.task,
      prepared: ready(ss), profile: RubricEngine.profileSummary(profile),
      profileVersions: profileEntries.length, submissions: submissions,
      classroom: link ? { courseId: text(link.CourseID), courseName: text(link.CourseName), courseWorkId: text(link.CourseWorkID),
        assignmentTitle: text(link.AssignmentTitle), dueDate: text(link.DueDate), maxPoints: toNumber(link.MaxPoints),
        autoImported: yes(link.AutoImported) } : null,
      stats: { total: current.length, approved: graded.length, awaiting: current.length - graded.length,
        drafts: current.filter(function (s) { return s.status === 'InReview'; }).length,
        evidenceChanged: current.filter(function (s) { return !!s.evidenceChangedAt; }).length }
    } };
  }

  function detail(workbookId, taskName, submissionId, actor) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    if (!ready(ss)) throw new Error('Enable web grading for this workbook first.');
    var sub = selectSubmission(ss, target.task, submissionId);
    var a = assess(ss, target.task, sub);
    var history = records(ss, 'AssessmentHistory').filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId); }).map(function (r) {
      return { action: text(r.Action), when: text(r.Timestamp), by: text(r.ActorEmail), notes: text(r.Notes) };
    });
    var snapshot = snapshotFor(ss, submissionId);
    var partialApproval = !!snapshot && ['Approved','Locked'].indexOf(text(sub.Status)) === -1;
    var locked = text(sub.Status) === 'Locked';
    var readOnly = a.legacy || partialApproval || FINAL.indexOf(text(sub.Status)) !== -1 || !isCurrent(sub);
    var user = actor || WebAccess.requireTeacher();
    var reassessment = !!text(sub.ParentSubmissionRecordID);
    return { success: true, data: {
      submission: { submissionRecordId: text(sub.SubmissionRecordID), studentName: text(sub.StudentName),
        studentEmail: text(sub.StudentEmail), studentUserId: text(sub.StudentUserID), taskName: target.task, className: text(sub.Class),
        status: text(sub.Status) || 'New', stateLabel: AppConfig.stateLabel(sub.Status, reassessment, sub.SubmissionVersion),
        version: Number(sub.SubmissionVersion) || 1, approvedAt: text(sub.ApprovedAt), lockedAt: text(sub.LockedAt),
        currentOfficial: isCurrent(sub), reassessment: reassessment,
        late: yes(sub.Late), parentId: text(sub.ParentSubmissionRecordID), sourceType: text(sub.SourceType),
        evidenceChangedAt: text(sub.EvidenceChangedAt), hasPendingEvidence: !!text(sub.PendingEvidenceJSON),
        classroomState: text(sub.ClassroomState), published: text(sub.LastSyncResult), publishedAt: text(sub.LastClassroomSyncAt),
        hasClassroomIds: !!(text(sub.ClassroomCourseID) && text(sub.ClassroomSubmissionID)) },
      workbook: { id: workbookId, name: ss.getName() },
      profile: a.profile, assessmentId: a.assessmentId, criteriaMap: a.states,
      feedback: a.feedback, attachments: a.attachments, history: history,
      scores: a.scores, flags: a.flags, snapshot: snapshot,
      evidenceHistory: typeof EvidenceService !== 'undefined' ? EvidenceService.history(ss, submissionId) : [],
      comparison: compareWithParent(ss, sub, a.profile, a.states),
      revision: lastRevision(ss, submissionId), readOnly: readOnly, locked: locked,
      canUnlock: locked && WebAccess.canUnlock(user),
      legacy: a.legacy,
      warning: partialApproval ? 'An approval snapshot exists but the submission status was not finalised. This version is read-only until an administrator reconciles the workbook from its backup and history.' :
        a.legacy ? 'This submission contains assessment rows without a single pinned rubric version. It is read-only here; reconcile it before web grading.' :
        !a.profile ? 'No approved rubric is active for this task. Review and activate a rubric before grading.' : ''
    } };
  }

  /* ---- validating teacher input ---------------------------------------- */

  function validateChecks(profile, incoming) {
    if (!incoming || typeof incoming !== 'object' || Array.isArray(incoming)) throw new Error('Assessment criteria are required.');
    var known = RubricEngine.knownObservableIds(profile);
    var clean = {};
    for (var i = 0; i < profile.criteria.length; i++) {
      var c = profile.criteria[i];
      var state = incoming[c.criterionId] || {};
      var checks = state.checkboxes || {};
      if (typeof checks !== 'object' || Array.isArray(checks)) throw new Error('Invalid checkbox state.');
      var map = {};
      Object.keys(checks).forEach(function (id) {
        if (known[id] !== c.criterionId) throw new Error('Unknown observable ' + id + ' for ' + c.criterionId + '. Refresh the rubric.');
        if (checks[id] === true) map[id] = true;
      });
      var note = text(state.teacherWrittenNote);
      if (note.length > 6000) throw new Error('Criterion note is too long (max 6,000 characters).');
      var review = state.review || {};
      clean[c.criterionId] = { checkboxes: map, teacherWrittenNote: note,
        review: { reviewed: !!review.reviewed, noEvidence: !!review.noEvidence && !Object.keys(map).length } };
    }
    Object.keys(incoming).forEach(function (id) {
      if (!profile.criteria.some(function (c) { return c.criterionId === id; })) throw new Error('Unknown criterion ' + id + '. Refresh the rubric.');
    });
    return clean;
  }

  var FEEDBACK_FIELDS = { whatWentWell: 6000, areasForImprovement: 6000, goalsForNextAssessment: 6000, teacherEvaluation: 6000, generalObservation: 4000 };

  /* Student-facing fields: whatWentWell, areasForImprovement, goalsForNextAssessment.
     Teacher-only: teacherEvaluation (personal evaluation), generalObservation. */
  function validateFeedback(input) {
    input = input || {};
    var out = {};
    Object.keys(FEEDBACK_FIELDS).forEach(function (k) {
      out[k] = text(input[k]);
      if (out[k].length > FEEDBACK_FIELDS[k]) throw new Error('Feedback field is too long (max ' + FEEDBACK_FIELDS[k] + ' characters).');
    });
    out.acknowledgedUnreadable = !!input.acknowledgedUnreadable;
    out.basisHash = /^h[0-9a-z]{1,16}$/.test(text(input.basisHash)) ? text(input.basisHash) : '';
    return out;
  }

  function studentFeedback(fb) {
    return { whatWentWell: text(fb.whatWentWell), areasForImprovement: text(fb.areasForImprovement), goalsForNextAssessment: text(fb.goalsForNextAssessment) };
  }

  function normaliseOptions(options) {
    if (typeof options === 'string') return { assessmentType: options };
    return options && typeof options === 'object' ? options : {};
  }

  /* ---- writing a draft -------------------------------------------------- */

  /**
   * Saves the teacher's (or an AI proposal's) draft. Never approves, never publishes.
   * @param {Object} incoming {criterionId: {checkboxes, teacherWrittenNote, review:{reviewed,noEvidence}}}
   * @param {Object} feedback student feedback + teacher-only fields; markFeedbackCurrent stamps the basis
   * @param {string} expectedRevision optimistic concurrency token (last HistoryID)
   * @param {string} source 'Teacher' | 'AI'
   * @param {Object=} options server-side only: {ai:{criterionId:{checks,evidence,grade,insufficient,unsupported}}, actor, assessmentType, deductionRule}
   */
  function saveDraft(workbookId, taskName, submissionId, incoming, feedback, expectedRevision, source, options) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    if (!ready(ss)) throw new Error('Enable web grading first.');
    var opt = normaliseOptions(options);
    var actor = opt.actor || WebAccess.requireTeacher();
    return withLock(function () {
      upgrade(ss);
      var sub = selectSubmission(ss, target.task, submissionId);
      editable(sub);
      if (text(expectedRevision) !== lastRevision(ss, submissionId)) throw new Error('A newer assessment exists. Reload this student before saving.');
      var existing = records(ss, 'CriterionAssessments').filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId); });
      if (existing.some(function (r) { return ['Approved','Locked'].indexOf(text(r.Status)) !== -1; }) || !!snapshotFor(ss, submissionId)) {
        throw new Error('An approved assessment already exists. Start a new reassessment version.');
      }
      var assessmentId = existing.length ? text(existing[existing.length - 1].AssessmentID) : uid();
      var pinnedIds = existing.map(function (r) { return text(r.RubricProfileID); });
      if (pinnedIds.some(function (id) { return !id || id !== pinnedIds[0]; })) throw new Error('Legacy or conflicting assessment rows require reconciliation before web grading.');
      var profile = pinnedIds.length ? RubricEngine.getProfileById(ss, pinnedIds[0]) : RubricEngine.getActiveProfile(ss, target.task);
      if (!profile) throw new Error('Activate a reviewed rubric for this task first.');
      var clean = validateChecks(profile, incoming);
      var previous = latestFeedback(ss, submissionId, assessmentId);
      var fb = validateFeedback(feedback);
      fb.basisHash = feedback && feedback.markFeedbackCurrent === true ?
        AssessmentFlags.feedbackBasis(profile, clean, fb.generalObservation) : text(previous.basisHash);
      var scored = RubricEngine.scoreProfile(profile, clean);
      var deduction = RubricEngine.calculateDeduction(scored.totalPoints, scored.totalMaxMarks, opt.deductionRule, sub);
      var timestamp = now();

      profile.criteria.forEach(function (c) {
        var state = clean[c.criterionId];
        var score = scored.criteria[c.criterionId];
        var old = existing.filter(function (r) { return text(r.CriterionID) === c.criterionId; })[0];
        var ai = opt.ai && opt.ai[c.criterionId];
        var aiProposal = ai ? { checks: ai.checks || {}, insufficient: !!ai.insufficient, unsupported: ai.unsupported || [], provider: text(ai.provider), at: timestamp } :
          (old ? parseJson(old.AIProposedChecksJSON, null) : null);
        var fields = {
          AssessmentID: assessmentId, SubmissionRecordID: submissionId, CriterionID: c.criterionId,
          AIProposedGrade: ai ? text(ai.grade).slice(0, 16) : (old ? text(old.AIProposedGrade) : ''),
          DerivedGrade: score.grade,
          DistinctOverrideSelected: !!state.checkboxes[c.mdOverride.id],
          CheckboxesJSON: JSON.stringify(state.checkboxes),
          AIEvidenceJSON: ai ? JSON.stringify(ai.evidence || []) : (old ? text(old.AIEvidenceJSON) || '[]' : '[]'),
          TeacherAdjusted: !!(aiProposal && AssessmentFlags.canonical(AssessmentFlags.ticked(aiProposal.checks)) !== AssessmentFlags.canonical(AssessmentFlags.ticked(state.checkboxes))),
          TeacherAudioTranscript: old ? text(old.TeacherAudioTranscript) : '',
          TeacherWrittenNote: state.teacherWrittenNote,
          MinimumEvidenceIncomplete: score.incompleteMinimumEvidence,
          Status: 'InReview', UpdatedAt: timestamp, RubricProfileID: profile.profileId,
          TeacherReviewJSON: JSON.stringify(state.review),
          AIProposedChecksJSON: aiProposal ? JSON.stringify(aiProposal) : ''
        };
        if (old) patch(ss, 'CriterionAssessments', old._row, fields);
        else { fields.CreatedAt = timestamp; append(ss, 'CriterionAssessments', fields); }
      });
      var subFields = {
        Status: 'InReview',
        AssessmentType: opt.assessmentType || text(sub.AssessmentType) || 'Summative',
        RawScore: scored.totalPoints,
        DeductionPoints: deduction.deductionPoints,
        DeductionReason: deduction.reason,
        FinalScore: deduction.finalScore
      };
      if (source !== 'AI' && text(sub.EvidenceChangedAt)) subFields.EvidenceChangedAt = '';
      patch(ss, 'Submissions', sub._row, subFields);
      append(ss, 'AIAssessments', {
        AssessmentID: assessmentId, SubmissionRecordID: submissionId, RunType: 'TeacherDraft',
        RunTimestamp: timestamp, ModelUsed: '', PromptVersion: 'web-v2', InputFilesJSON: '[]',
        PromptSummary: 'Teacher-reviewed feedback', ResponseJSON: JSON.stringify(fb),
        ExecutionStatus: 'Success', TriggeredBy: actor, TeacherOutcome: 'Pending'
      });
      var revision = event(ss, submissionId, assessmentId, source === 'AI' ? 'AIProposed' : 'DraftSaved',
        { revision: expectedRevision }, { rubricProfileId: profile.profileId, criteria: clean, feedback: fb, scores: scored, deduction: deduction },
        actor, 'Rubric ' + profile.name + ' — teacher approval still required.');
      var states = {};
      Object.keys(clean).forEach(function (cid) {
        var row = existing.filter(function (r) { return text(r.CriterionID) === cid; })[0];
        var ai = opt.ai && opt.ai[cid];
        states[cid] = { checkboxes: clean[cid].checkboxes, teacherWrittenNote: clean[cid].teacherWrittenNote, review: clean[cid].review,
          aiProposal: ai ? { checks: ai.checks || {}, insufficient: !!ai.insufficient, unsupported: ai.unsupported || [] } : (row ? parseJson(row.AIProposedChecksJSON, null) : null) };
      });
      var flags = AssessmentFlags.compute({ profile: profile, states: states, scores: scored, attachments: attachmentsFor(ss, submissionId), feedback: fb });
      return { success: true, message: 'Draft saved to the workbook. Nothing has been approved or sent to Classroom.',
        revision: revision, assessmentId: assessmentId, scores: scored, deduction: deduction, flags: flags, savedAt: timestamp };
    });
  }

  /* Grounded feedback from the on-screen (possibly unsaved) selections. Writes nothing. */
  function buildFeedback(workbookId, taskName, submissionId, incoming, generalObservation) {
    var target = openTask(workbookId, taskName);
    var sub = selectSubmission(target.ss, target.task, submissionId);
    var a = assess(target.ss, target.task, sub);
    if (!a.profile) throw new Error('No rubric is active for this task.');
    var clean = validateChecks(a.profile, incoming || a.states);
    var scores = RubricEngine.scoreProfile(a.profile, clean);
    var built = FeedbackBuilder.build(a.profile, clean, scores, { generalObservation: generalObservation });
    return { success: true, data: built, message: 'Feedback rebuilt from ' + built.sources.ticked + ' ticked observable(s), ' + built.sources.targets +
      ' next-step target(s) and ' + built.sources.notes + ' teacher observation(s). Edit it before saving.' };
  }

  /* ---- approval, lock, unlock, reassessment ------------------------------ */

  function approve(workbookId, taskName, submissionId, expectedRevision, options) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var opt = normaliseOptions(options);
    var user = opt.actor || WebAccess.requireTeacher();
    return withLock(function () {
      var sub = selectSubmission(ss, target.task, submissionId);
      editable(sub);
      if (text(expectedRevision) !== lastRevision(ss, submissionId)) throw new Error('A newer assessment exists. Reload this student before approval.');
      var all = records(ss, 'CriterionAssessments').filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId); });
      if (!all.length) throw new Error('Save a complete draft before approving.');
      if (snapshotFor(ss, submissionId)) throw new Error('An approved snapshot already exists for this version. Reconcile the workbook rather than approving twice.');
      var a = assess(ss, target.task, sub);
      var rows = a.rows;
      var pinnedId = text(rows[0].RubricProfileID);
      if (!pinnedId || rows.some(function (r) { return text(r.RubricProfileID) !== pinnedId; })) throw new Error('Rubric version is missing or inconsistent.');
      var profile = a.profile;
      if (!profile || profile.taskName !== target.task) throw new Error('The pinned rubric version is unavailable.');
      profile.criteria.forEach(function (c) {
        if (!rows.some(function (r) { return text(r.CriterionID) === c.criterionId; })) throw new Error('Missing grade for ' + c.title + '.');
      });
      var blocking = a.flags.flags.filter(function (f) { return f.severity === 'blocking'; });
      if (blocking.length) {
        throw new Error('Resolve before approving: ' + blocking.map(function (f) { return f.message; }).join(' '));
      }
      var scored = a.scores;
      var deduction = RubricEngine.calculateDeduction(scored.totalPoints, scored.totalMaxMarks, opt.deductionRule, sub);
      var feedback = validateFeedback(a.feedback);
      var timestamp = now();
      var outcomeCoverage = RubricEngine.calculateOutcomeCoverage(profile, rows);

      /* Append the immutable approved snapshot FIRST; if it fails, the draft
         stays editable. Approval never contacts Classroom. */
      append(ss, 'ApprovedGrades', {
        GradeID: uid(), TaskName: target.task, SubmissionRecordID: submissionId,
        StudentUserID: sub.StudentUserID, StudentName: sub.StudentName, AssessmentID: a.assessmentId,
        RubricProfileID: pinnedId, Points: deduction.finalScore, MaxMarks: scored.totalMaxMarks,
        Percent: RubricEngine.normalisePercent(deduction.percent), Grade: scored.letter, FeedbackJSON: JSON.stringify(studentFeedback(feedback)),
        ApprovedAt: timestamp, ApprovedBy: user,
        AssessmentType: opt.assessmentType || text(sub.AssessmentType) || 'Summative',
        RawScore: scored.totalPoints, DeductionPoints: deduction.deductionPoints,
        DeductionReason: deduction.reason, FinalScore: deduction.finalScore,
        OutcomesJSON: JSON.stringify(outcomeCoverage)
      });
      rows.forEach(function (r) {
        patch(ss, 'CriterionAssessments', r._row, {
          DerivedGrade: scored.criteria[r.CriterionID].grade,
          FinalApprovedGrade: scored.criteria[r.CriterionID].grade,
          Status: 'Approved', ApprovedAt: timestamp, ApprovedBy: user
        });
      });
      patch(ss, 'Submissions', sub._row, {
        Status: 'Approved', ApprovedAt: timestamp, ApprovedBy: user,
        RawScore: scored.totalPoints, DeductionPoints: deduction.deductionPoints,
        DeductionReason: deduction.reason, FinalScore: deduction.finalScore,
        AssessmentType: opt.assessmentType || text(sub.AssessmentType) || 'Summative'
      });
      var revision = event(ss, submissionId, a.assessmentId, 'Approved', {},
        { rubricProfileId: pinnedId, scores: scored, deduction: deduction, feedback: feedback }, user,
        'Approved by teacher. Classroom was not changed.');
      if (typeof Refinement !== 'undefined') {
        try { Refinement.recordApproval(ss, target.task, sub, profile, a.states, feedback, user); }
        catch (e) { console.error('Refinement example not recorded: ' + errString(e)); }
      }
      return { success: true, message: 'Approved and saved to ApprovedGrades. Nothing was published to Classroom.', scores: scored, deduction: deduction, revision: revision, approvedAt: timestamp };
    });
  }

  function lockAssessment(workbookId, taskName, submissionId) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var user = WebAccess.requireTeacher();
    return withLock(function () {
      var sub = selectSubmission(ss, target.task, submissionId);
      if (text(sub.Status) !== 'Approved') throw new Error('Approve the assessment before locking it.');
      if (!isCurrent(sub)) throw new Error('Only the current submission can be locked.');
      if (!snapshotFor(ss, submissionId)) throw new Error('No web-approved grade snapshot exists for this submission. Reconcile legacy assessment data first.');
      var timestamp = now();
      patch(ss, 'Submissions', sub._row, { Status: 'Locked', LockedAt: timestamp, LockedBy: user });
      var rows = records(ss, 'CriterionAssessments').filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId); });
      rows.forEach(function (r) { patch(ss, 'CriterionAssessments', r._row, { Status: 'Locked' }); });
      var revision = event(ss, submissionId, rows.length ? rows[rows.length - 1].AssessmentID : '', 'Locked', {}, {}, user, 'Final lock by teacher.');
      return { success: true, message: 'Final-locked. Ordinary controls can no longer change this assessment.', revision: revision };
    });
  }

  /* Authorised unlock: Locked → Approved. Requires LESSON_GRADER_UNLOCK_EMAILS,
     a reason, and the typed word UNLOCK. The approval snapshot is unchanged;
     further change still requires a reassessment version. (SC41) */
  function unlock(workbookId, taskName, submissionId, reason, confirmation) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var user = WebAccess.requireTeacher();
    if (!WebAccess.canUnlock(user)) throw new Error('Your account is not authorised to unlock final assessments (LESSON_GRADER_UNLOCK_EMAILS).');
    var why = text(reason).trim();
    if (why.length < 10 || why.length > 500) throw new Error('Give a reason for unlocking (10–500 characters).');
    if (text(confirmation).trim() !== 'UNLOCK') throw new Error('Type UNLOCK to confirm.');
    upgrade(ss);
    return withLock(function () {
      var sub = selectSubmission(ss, target.task, submissionId);
      if (text(sub.Status) !== 'Locked') throw new Error('Only a final-locked assessment can be unlocked.');
      var timestamp = now();
      patch(ss, 'Submissions', sub._row, { Status: 'Approved', UnlockedAt: timestamp, UnlockedBy: user, UnlockReason: why });
      var rows = records(ss, 'CriterionAssessments').filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId) && text(r.Status) === 'Locked'; });
      rows.forEach(function (r) { patch(ss, 'CriterionAssessments', r._row, { Status: 'Approved' }); });
      var revision = event(ss, submissionId, rows.length ? rows[rows.length - 1].AssessmentID : '', 'Unlocked',
        { status: 'Locked', lockedAt: text(sub.LockedAt), lockedBy: text(sub.LockedBy) }, { status: 'Approved', reason: why }, user, 'Authorised unlock: ' + why);
      return { success: true, message: 'Unlocked. The approved result is unchanged; start a reassessment to change it.', revision: revision };
    });
  }

  function reassess(workbookId, taskName, submissionId) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var user = WebAccess.requireTeacher();
    upgrade(ss);
    return withLock(function () {
      var sub = selectSubmission(ss, target.task, submissionId);
      if (text(sub.Status) === 'Locked') throw new Error('This assessment is final-locked. An authorised teacher must unlock it before reassessment.');
      if (text(sub.Status) !== 'Approved') throw new Error('Only an approved assessment can start a reassessment.');
      if (!snapshotFor(ss, submissionId)) throw new Error('Reconcile this legacy assessment before creating a web reassessment.');
      if (!isCurrent(sub)) throw new Error('The selected version is not current.');
      if (records(ss, 'Submissions').some(function (r) { return text(r.ParentSubmissionRecordID) === text(submissionId); })) {
        throw new Error('A reassessment version already exists. Open it instead.');
      }
      var pending = parseJson(sub.PendingEvidenceJSON, null);
      var id = uid();
      var fields = {};
      SUB.forEach(function (key) { fields[key] = sub[key] === undefined ? '' : sub[key]; });
      fields.SubmissionRecordID = id;
      fields.SubmissionVersion = (Number(sub.SubmissionVersion) || 1) + 1;
      fields.SourceType = 'Reassessment';
      fields.ParentSubmissionRecordID = submissionId;
      /* Keep the new row non-official until its file-copy step succeeds. */
      fields.CurrentOfficial = false;
      fields.Status = 'New';
      fields.ApprovedAt = fields.ApprovedBy = fields.LockedAt = fields.LockedBy = '';
      fields.LastClassroomSyncAt = fields.LastSyncResult = fields.ClassroomAssignedGrade = '';
      fields.RawScore = fields.DeductionPoints = fields.DeductionReason = fields.FinalScore = '';
      fields.UpdateTime = now();
      if (pending) {
        fields.AttachmentFileIDsJSON = JSON.stringify(pending.fileIds || []);
        fields.AttachmentMetadataJSON = JSON.stringify(pending.attachments || []);
        fields.AttachmentSummary = (pending.attachments || []).length + ' Drive file(s)';
        fields.ClassroomState = text(pending.state) || fields.ClassroomState;
        fields.UpdateTime = text(pending.updateTime) || fields.UpdateTime;
      }
      var newRow = append(ss, 'Submissions', fields);
      if (pending) {
        (pending.attachments || []).forEach(function (f) {
          append(ss, 'SubmissionFiles', {
            SubmissionRecordID: id, FileRecordID: uid(), SourceType: 'driveFile', DriveFileID: text(f.id),
            FileName: text(f.title) || 'Drive file', MimeType: text(f.mimeType), AlternateLink: text(f.alternateLink),
            ThumbnailUrl: text(f.thumbnailUrl), EligibleForAI: DriveService.isEligibleForAi(text(f.mimeType)),
            AIReviewStatus: DriveService.isEligibleForAi(text(f.mimeType)) ? 'Pending' : 'Unsupported', CreatedAt: now()
          });
        });
      } else {
        records(ss, 'SubmissionFiles').filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId) && text(r.AIReviewStatus) !== 'Removed'; }).forEach(function (r) {
          var copy = {};
          FILES.forEach(function (key) { copy[key] = r[key] === undefined ? '' : r[key]; });
          copy.SubmissionRecordID = id;
          copy.FileRecordID = uid();
          copy.AIReviewStatus = text(r.AIReviewStatus) === 'Reviewed' ? 'Pending' : text(r.AIReviewStatus) || 'Pending';
          copy.CreatedAt = now();
          append(ss, 'SubmissionFiles', copy);
        });
      }
      patch(ss, 'Submissions', sub._row, { CurrentOfficial: false });
      if (pending) patch(ss, 'Submissions', sub._row, { PendingEvidenceJSON: '' });
      patch(ss, 'Submissions', newRow, { CurrentOfficial: true });
      event(ss, id, '', 'ReassessmentCreated', {}, { parentId: submissionId, version: fields.SubmissionVersion, newEvidence: !!pending }, user,
        'Earlier approval retained' + (pending ? '; uses evidence submitted after approval.' : '.'));
      return { success: true, message: 'Reassessment version ' + fields.SubmissionVersion + ' created. The earlier approval was not changed.', submissionRecordId: id };
    });
  }

  function previewRubric(workbookId, taskName, rubricFileId) {
    var target = openTask(workbookId, taskName);
    var sourceId = text(rubricFileId).trim();
    if (sourceId && !LessonGraderWeb.isSharedRubric(sourceId)) throw new Error('Choose a rubric from the Shared Rubrics library.');
    var rows = sourceId ? RubricEngine.readRubricRows(sourceId) : [];
    var profile = rows.length ? RubricEngine.buildProfileFromRubricRows(rows, {
      taskName: target.task, name: sourceId ? DriveApp.getFileById(sourceId).getName() : '',
      weights: RubricEngine.getGradeScale(target.ss).weights,
      bands: RubricEngine.getGradeScale(target.ss).bands
    }) : null;
    return { success: true, data: { rows: rows, preview: profile, validation: profile ? RubricEngine.validateProfile(profile) : null } };
  }

  function activateRubric(workbookId, taskName, name, sourceId, rows, customWeights, confirmed) {
    if (confirmed !== true) throw new Error('Review the rubric criteria, marks and grade-band mapping before activating.');
    var target = openTask(workbookId, taskName);
    if (!ready(target.ss)) throw new Error('Enable web grading before activating a rubric.');
    var fileId = text(sourceId).trim();
    if (fileId && !LessonGraderWeb.isSharedRubric(fileId)) throw new Error('Source rubric is not in Shared Rubrics.');
    if (!Array.isArray(rows) || !rows.length) throw new Error('Add at least one rubric observable.');
    var scale = RubricEngine.getGradeScale(target.ss);
    var weights = scale.weights;
    Object.keys(customWeights || {}).forEach(function (band) {
      /* Explicit mappings only for non-A..E bands. Never let client silently
         rewrite official Setup grade weights for standard bands. */
      if (RubricEngine.CORE_BANDS.indexOf(band) !== -1) return;
      var weight = Number(customWeights[band]);
      if (!text(customWeights[band]).trim() || !isFinite(weight) || weight < 0 || weight > 1) throw new Error('Band ' + band + ' needs a weight from 0 to 1.');
      weights[band.toUpperCase()] = weight;
    });
    var profile = RubricEngine.buildProfileFromRubricRows(rows, {
      taskName: target.task, name: text(name).trim().slice(0, 160),
      source: fileId ? 'sharedRubric' : 'manual', sourceRubricFileId: fileId,
      weights: weights, bands: scale.bands
    });
    var validation = RubricEngine.validateProfile(profile);
    if (!validation.ok) throw new Error(validation.errors.join(' '));
    var actor = WebAccess.requireTeacher();
    return withLock(function () {
      var saved = RubricEngine.saveProfile(target.ss, profile, actor);
      if (!saved.success) throw new Error(saved.message);
      writeRubricMirror(target.ss, target.task, saved.profile);
      return { success: true, message: 'Rubric activated for ' + target.task + ' and copied to the RubricRows sheet. Existing assessments remain pinned to their original rubric.', data: saved.summary };
    });
  }

  function writeRubricMirror(ss, task, profile) {
    var sheet = ensureSheet(ss, 'RubricRows', RUBRIC_ROWS);
    var data = sheet.getDataRange().getValues();
    var header = data[0].map(function (h) { return text(h).trim(); });
    var ti = header.indexOf('TaskName');
    var keep = data.slice(1).filter(function (r) { return text(r[ti]).trim() !== task && r.join('') !== ''; });
    var stamp = now();
    var fresh = RubricEngine.profileToRows(profile).map(function (r) {
      var v = { TaskName: task, Criterion: r.criterion, Part: r.part, Section: r.section, MaxMarks: r.maxMarks, Outcome: r.outcome,
        Band: r.band, Description: r.description, BandWeight: r.weight, ProfileID: profile.profileId, SyncedAt: stamp };
      return header.map(function (h) { return v[h] === undefined ? '' : v[h]; });
    });
    var all = [header].concat(keep, fresh);
    sheet.clearContents();
    sheet.getRange(1, 1, all.length, header.length).setValues(all);
  }

  function rowKey(r) {
    return [text(r.criterion).trim(), text(r.part).trim(), text(r.section).trim(), toNumber(r.maxMarks), text(r.outcome).trim(),
      text(r.band).trim().toUpperCase(), text(r.description).trim(), r.weight === '' || r.weight == null ? '' : toNumber(r.weight)];
  }

  /** Active rubric vs the workbook's RubricRows sheet for one task (SC13, SC14). */
  function rubricSource(workbookId, taskName) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var profile = RubricEngine.getActiveProfile(ss, target.task);
    var activeRows = profile ? RubricEngine.profileToRows(profile) : [];
    var sheetRows = ss.getSheetByName('RubricRows') ? records(ss, 'RubricRows').filter(function (r) { return text(r.TaskName).trim() === target.task && text(r.Criterion).trim(); }).map(function (r) {
      return { criterion: text(r.Criterion), part: text(r.Part), section: text(r.Section), maxMarks: r.MaxMarks, outcome: text(r.Outcome),
        band: text(r.Band), description: text(r.Description), weight: r.BandWeight };
    }) : [];
    var drift = !!profile && sheetRows.length > 0 && AssessmentFlags.canonical(activeRows.map(rowKey)) !== AssessmentFlags.canonical(sheetRows.map(rowKey));
    var versions = RubricEngine.readProfiles(ss).filter(function (p) { return text(p.row.TaskName) === target.task; }).map(function (p) {
      return { profileId: text(p.row.ProfileID), name: text(p.row.ProfileName), createdAt: text(p.row.CreatedAt), createdBy: text(p.row.CreatedBy), active: yes(p.row.Active) };
    });
    return { success: true, data: { profile: profile, summary: RubricEngine.profileSummary(profile), activeRows: activeRows,
      sheetRows: sheetRows, drift: drift, versions: versions, gradeScale: RubricEngine.getGradeScale(ss),
      referencePoints: RubricEngine.referencePointReport(RubricEngine.getGradeScale(ss).weights) } };
  }

  /* AI drafting lives in AiGateway.gs (provider calls, evidence validation,
     de-identification). Kept here as the stable entry point for callers. */
  function runAi(workbookId, taskName, submissionId, expectedRevision, options) {
    return AiGateway.draft(workbookId, taskName, submissionId, expectedRevision, options);
  }

  function report(workbookId, taskName, submissionId) {
    var d = detail(workbookId, taskName, submissionId).data;
    if (!d.profile || d.legacy) throw new Error('No pinned rubric is available for this assessment.');
    if (['Approved','Locked'].indexOf(d.submission.status) === -1) throw new Error('Approve the assessment before producing its final report.');
    var html = RubricEngine.renderReportHtml(d.profile, {
      submission: { StudentName: d.submission.studentName, Class: d.submission.className, SubmissionVersion: d.submission.version },
      criteriaMap: d.criteriaMap, feedback: d.feedback
    }, null, { taskName: taskName, date: text(d.submission.approvedAt).slice(0, 10) });
    var blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
    var safe = d.submission.studentName.replace(/[^a-z0-9_-]/gi, '_').slice(0, 64);
    var fileName = safe + '_' + text(taskName).replace(/[^a-z0-9_-]/gi, '_').slice(0, 64) + '_Assessment.pdf';
    return { success: true, fileName: fileName, dataUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(blob.getBytes()) };
  }

  /* For a project outside Classroom, Sheets itself becomes the submission
     source of truth. No Classroom ID is invented and no grade is synced. */
  function addManualSubmission(workbookId, taskName, studentName, studentEmail, className) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    if (!ready(ss)) throw new Error('Enable web grading before adding a student.');
    var name = text(studentName).trim();
    var email = text(studentEmail).trim().toLowerCase();
    var course = text(className).trim();
    if (!name || name.length > 160 || email.length > 180 || course.length > 100 ||
        (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))) {
      throw new Error('Enter a student name and, optionally, a valid school email and class.');
    }
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var duplicate = records(ss, 'Submissions').some(function (r) {
        return text(r.Task) === target.task &&
          (r.CurrentOfficial === '' || yes(r.CurrentOfficial)) &&
          ((email && text(r.StudentEmail).toLowerCase() === email) ||
           (!email && text(r.StudentName).trim().toLowerCase() === name.toLowerCase()));
      });
      if (duplicate) throw new Error('A current submission for this student and task already exists. Open it or create a reassessment.');
      var id = uid();
      append(ss, 'Submissions', {
        SubmissionRecordID: id, StudentUserID: email || ('manual:' + id),
        StudentName: name, StudentEmail: email, Class: course, Task: target.task,
        SubmissionVersion: 1, SourceType: 'Manual', Status: 'New', CurrentOfficial: true,
        TurnedInTime: now(), UpdateTime: now(), AttachmentSummary: 'Teacher-entered (no Drive attachments)',
        AttachmentFileIDsJSON: '[]', AttachmentMetadataJSON: '[]'
      });
      event(ss, id, '', 'ManualSubmissionCreated', {}, { taskName: target.task, studentName: name },
        WebAccess.requireTeacher(), 'Manually entered in Sheets; no Classroom link or sync.');
      return { success: true, submissionRecordId: id, message: 'Student added to this task in Sheets. No Classroom record or grade was created.' };
    } finally { lock.releaseLock(); }
  }

  /* Normalises Classroom attachments. Drive files are readable evidence;
     links, YouTube and Forms are recorded and disclosed as not readable. */
  function classroomAttachments(item) {
    var out = [];
    ((item.assignmentSubmission && item.assignmentSubmission.attachments) || []).forEach(function (a) {
      if (a.driveFile) out.push({ kind: 'driveFile', id: text(a.driveFile.id), title: text(a.driveFile.title), mimeType: text(a.driveFile.mimeType),
        alternateLink: text(a.driveFile.alternateLink), thumbnailUrl: text(a.driveFile.thumbnailUrl) });
      else if (a.link) out.push({ kind: 'link', id: 'link:' + text(a.link.url), title: text(a.link.title) || text(a.link.url), alternateLink: text(a.link.url) });
      else if (a.youTubeVideo) out.push({ kind: 'youTube', id: 'yt:' + text(a.youTubeVideo.id), title: text(a.youTubeVideo.title), alternateLink: text(a.youTubeVideo.alternateLink) });
      else if (a.form) out.push({ kind: 'form', id: 'form:' + text(a.form.formUrl), title: text(a.form.title), alternateLink: text(a.form.formUrl) });
    });
    return out;
  }

  function appendFileRow(ss, recordId, f) {
    var drive = f.kind === 'driveFile';
    var eligible = drive && DriveService.isEligibleForAi(text(f.mimeType));
    append(ss, 'SubmissionFiles', {
      SubmissionRecordID: recordId, FileRecordID: uid(), SourceType: f.kind || 'driveFile',
      DriveFileID: drive ? text(f.id) : '', FileName: text(f.title) || 'Attachment', MimeType: text(f.mimeType),
      AlternateLink: text(f.alternateLink), ThumbnailUrl: text(f.thumbnailUrl), EligibleForAI: eligible,
      AIReviewStatus: drive ? 'Pending' : 'Unsupported', Limitations: drive ? '' : 'Linked content (' + f.kind + ') cannot be read automatically; open it to review.',
      CreatedAt: now()
    });
  }

  /**
   * Imports or refreshes one task's Classroom submissions. Idempotent.
   *  - new student → new submission row (identity = course + coursework + Classroom userId)
   *  - New/draft rows: evidence updated in place; teacher drafts are kept and flagged
   *  - Approved/locked rows: never altered; new evidence is held as pending until the
   *    teacher deliberately starts a reassessment (SC12)
   */
  function importTask(workbookId, taskName, options) {
    var opt = options || {};
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    if (!ready(ss)) throw new Error('Enable web grading before importing.');
    var actor = opt.actor || WebAccess.requireTeacher();
    var links = records(ss, 'ClassroomConfig').filter(function (r) { return text(r.TaskName) === target.task && yes(r.Active) && r.CourseID && r.CourseWorkID; });
    if (!links.length) throw new Error('Link a Classroom assignment to this task first.');
    var cfg = links[links.length - 1];
    var raw = [];
    var token = null;
    do {
      var params = { pageSize: 100 };
      if (token) params.pageToken = token;
      var page = Classroom.Courses.CourseWork.StudentSubmissions.list(text(cfg.CourseID), text(cfg.CourseWorkID), params);
      raw = raw.concat(page && page.studentSubmissions || []);
      token = page && page.nextPageToken;
    } while (token);
    return withLock(function () {
      upgrade(ss);
      var before = records(ss, 'Submissions');
      var files = records(ss, 'SubmissionFiles');
      var roster = records(ss, 'ClassLists');
      var stats = { newCount: 0, updated: 0, heldForApproved: 0, unchanged: 0, total: raw.length };
      var seen = {};
      for (var i = 0; i < raw.length; i++) {
        var item = raw[i];
        var uidString = text(item.userId);
        if (!uidString || seen[uidString]) continue;
        seen[uidString] = true;
        var versions = before.filter(function (r) {
          return text(r.ClassroomCourseID) === text(cfg.CourseID) && text(r.ClassroomCourseWorkID) === text(cfg.CourseWorkID) && text(r.StudentUserID) === uidString;
        });
        var prior = versions.filter(isCurrent)[0] || versions.sort(function (a, b) { return (Number(b.SubmissionVersion) || 1) - (Number(a.SubmissionVersion) || 1); })[0];
        var attachments = classroomAttachments(item);
        var fileIds = attachments.map(function (a) { return a.id; }).sort();
        var state = text(item.state), updateTime = text(item.updateTime);

        if (!prior) {
          var identity = null;
          try { identity = Classroom.UserProfiles.get(uidString); } catch (e) { /* keep the verified ID; never invent a name */ }
          var name = identity && identity.name && identity.name.fullName || ('Student ' + uidString);
          var email = identity && identity.emailAddress || '';
          var match = roster.filter(function (r) {
            return !!email && yes(r.Active) && text(r.SchoolEmail).trim().toLowerCase() === email.toLowerCase() &&
              (!text(r.ClassID).trim() || text(r.ClassID).trim() === text(cfg.CourseName).trim());
          })[0];
          if (match) name = text(match.PreferredName || match.OfficialName || name);
          if (email && before.some(function (r) {
            return text(r.Task) === target.task && !text(r.ClassroomCourseWorkID) && text(r.StudentEmail).toLowerCase() === email.toLowerCase() && isCurrent(r);
          })) {
            throw new Error('A manual submission already exists for ' + name + ' in this task. Reconcile it before importing their Classroom work. No duplicate was created.');
          }
          var recordId = uid();
          var fields = {
            SubmissionRecordID: recordId, ClassroomCourseID: text(cfg.CourseID), ClassroomCourseWorkID: text(cfg.CourseWorkID),
            ClassroomSubmissionID: text(item.id), StudentUserID: uidString, StudentName: name, StudentEmail: email,
            Class: text(cfg.CourseName), Task: target.task, SubmissionVersion: 1, SourceType: 'Classroom',
            ClassroomState: state, TurnedInTime: text(item.creationTime), UpdateTime: updateTime || now(), Late: !!item.late,
            AttachmentSummary: attachments.length + ' attachment(s)', AttachmentFileIDsJSON: JSON.stringify(fileIds),
            AttachmentMetadataJSON: JSON.stringify(attachments), Status: 'New', ParentSubmissionRecordID: '', CurrentOfficial: true
          };
          fields._row = append(ss, 'Submissions', fields);
          attachments.forEach(function (f) { appendFileRow(ss, recordId, f); });
          before.push(fields);
          stats.newCount++;
          event(ss, recordId, '', 'ClassroomImported', {}, { courseWorkId: text(cfg.CourseWorkID), fileIds: fileIds }, actor, 'Classroom evidence imported; no grade changed.');
          continue;
        }

        var oldIds = parseJson(prior.AttachmentFileIDsJSON, []).map(text).sort();
        var filesChanged = JSON.stringify(oldIds) !== JSON.stringify(fileIds);
        var timeChanged = !!updateTime && updateTime !== text(prior.UpdateTime);
        var status = text(prior.Status);
        var meta = {};
        if (state && state !== text(prior.ClassroomState)) meta.ClassroomState = state;
        if (!!item.late !== yes(prior.Late) && FINAL.indexOf(status) === -1) meta.Late = !!item.late;
        if (text(item.id) && !text(prior.ClassroomSubmissionID)) meta.ClassroomSubmissionID = text(item.id);

        if (FINAL.indexOf(status) !== -1) {
          if (filesChanged || timeChanged) {
            var pending = { fileIds: fileIds, attachments: attachments, state: state, updateTime: updateTime };
            if (text(prior.PendingEvidenceJSON) !== JSON.stringify(pending)) {
              meta.PendingEvidenceJSON = JSON.stringify(pending);
              meta.EvidenceChangedAt = now();
              patch(ss, 'Submissions', prior._row, meta);
              event(ss, text(prior.SubmissionRecordID), '', 'EvidenceUpdatedAfterApproval', { fileIds: oldIds }, { fileIds: fileIds, updateTime: updateTime }, actor,
                'New Classroom evidence arrived after approval. The approved result is unchanged; start a reassessment to use it.');
              stats.heldForApproved++;
              continue;
            }
          }
          if (Object.keys(meta).length) patch(ss, 'Submissions', prior._row, meta);
          stats.unchanged++;
          continue;
        }

        if (filesChanged) {
          var priorId = text(prior.SubmissionRecordID);
          files.filter(function (f) { return text(f.SubmissionRecordID) === priorId && text(f.AIReviewStatus) !== 'Removed'; }).forEach(function (f) {
            var key = text(f.SourceType) === 'driveFile' || !text(f.SourceType) ? text(f.DriveFileID) : null;
            var stillThere = attachments.some(function (a) { return a.id === key || (a.kind !== 'driveFile' && text(a.alternateLink) === text(f.AlternateLink)); });
            if (!stillThere) patch(ss, 'SubmissionFiles', f._row, { AIReviewStatus: 'Removed', Limitations: 'Removed from the Classroom submission on ' + now().slice(0, 10) + '.' });
          });
          attachments.filter(function (a) { return oldIds.indexOf(a.id) === -1; }).forEach(function (a) { appendFileRow(ss, priorId, a); });
          meta.AttachmentFileIDsJSON = JSON.stringify(fileIds);
          meta.AttachmentMetadataJSON = JSON.stringify(attachments);
          meta.AttachmentSummary = attachments.length + ' attachment(s)';
        }
        if (filesChanged || timeChanged) {
          meta.UpdateTime = updateTime || now();
          if (status === 'InReview') meta.EvidenceChangedAt = now();
          patch(ss, 'Submissions', prior._row, meta);
          event(ss, text(prior.SubmissionRecordID), '', 'ClassroomEvidenceUpdated', { fileIds: oldIds }, { fileIds: fileIds, updateTime: updateTime }, actor,
            status === 'InReview' ? 'Evidence changed after the draft was saved; the draft was kept for review.' : 'Evidence refreshed from Classroom.');
          stats.updated++;
        } else {
          if (Object.keys(meta).length) patch(ss, 'Submissions', prior._row, meta);
          stats.unchanged++;
        }
      }
      var msg = 'Synced ' + stats.total + ' Classroom submission(s): ' + stats.newCount + ' new, ' + stats.updated + ' updated, ' +
        stats.heldForApproved + ' held for approved work, ' + stats.unchanged + ' unchanged. No marks changed.';
      return { success: true, message: msg, stats: stats };
    });
  }

  /* Whole-class drafting goes through the queue so it survives outages and
     the execution time limit; this call queues then processes what fits. */
  function bulkGrade(workbookId, taskName, submissionIds, options) {
    var queued = AiQueue.enqueue(workbookId, taskName, submissionIds, options);
    var processed = AiQueue.process(workbookId, taskName, 4 * 60 * 1000, options && options.actor);
    return { success: true, queued: queued.queued, alreadyQueued: queued.alreadyQueued, skipped: queued.skipped,
      counts: processed.counts, remaining: processed.remaining, message: queued.message + ' ' + processed.message };
  }

  function verifyGradingComplete(workbookId, taskName) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var all = records(ss, 'Submissions').filter(function (r) {
      return text(r.Task).trim() === target.task && !!r.SubmissionRecordID && (r.CurrentOfficial === '' || yes(r.CurrentOfficial));
    });
    var pending = [];
    var approved = [];
    for (var i = 0; i < all.length; i++) {
      var s = all[i];
      var st = text(s.Status);
      if (['Approved', 'Locked'].indexOf(st) !== -1) {
        approved.push({ id: text(s.SubmissionRecordID), name: text(s.StudentName), status: st });
      } else {
        pending.push({ id: text(s.SubmissionRecordID), name: text(s.StudentName), status: st || 'New' });
      }
    }
    return {
      success: true,
      taskName: target.task,
      complete: pending.length === 0 && all.length > 0,
      total: all.length,
      approvedCount: approved.length,
      pendingCount: pending.length,
      pending: pending
    };
  }

  function setGradingComplete(workbookId, taskName, confirmed) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var check = verifyGradingComplete(workbookId, taskName);
    if (!check.complete) {
      throw new Error('Cannot lock cohort: ' + check.pendingCount + ' student(s) still ungraded or pending review.');
    }
    if (!confirmed) {
      return { success: true, preview: true, message: 'All ' + check.total + ' submissions are approved. Confirm to lock cohort grading.' };
    }
    var user = WebAccess.requireTeacher();
    var ts = now();
    append(ss, 'AuditLog', {
      AuditID: uid(), TaskName: target.task, Action: 'GradingCompleted',
      TargetID: target.task, TargetURL: '', ActorEmail: user,
      Timestamp: ts, Status: 'Complete', DetailsJSON: JSON.stringify({ cohortSize: check.total })
    });
    return { success: true, message: 'Cohort grading locked and verified by ' + user + ' at ' + ts + '.' };
  }

  function suggestWorkSamples(workbookId, taskName, policy) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var grades = records(ss, 'ApprovedGrades').filter(function (r) { return text(r.TaskName).trim() === target.task; });
    if (!grades.length) throw new Error('No approved grades found for ' + target.task + '. Complete grading before selecting samples.');
    var subRows = records(ss, 'Submissions');
    var fileRows = records(ss, 'SubmissionFiles');

    var candidates = grades.map(function (g) {
      var s = subRows.filter(function (r) { return text(r.SubmissionRecordID) === text(g.SubmissionRecordID); })[0] || {};
      var files = fileRows.filter(function (r) { return text(r.SubmissionRecordID) === text(g.SubmissionRecordID); });
      var score = toNumber(g.FinalScore != null ? g.FinalScore : g.Points);
      var max = toNumber(g.MaxMarks) || 100;
      var pct = max > 0 ? (score / max) * 100 : toNumber(g.Percent);
      return {
        submissionRecordId: text(g.SubmissionRecordID),
        studentName: text(g.StudentName),
        studentUserId: text(g.StudentUserID),
        grade: text(g.Grade),
        points: score,
        maxMarks: max,
        percent: pct,
        files: files.map(function (f) {
          return { fileId: text(f.DriveFileID), fileName: text(f.FileName), mimeType: text(f.MimeType), link: text(f.AlternateLink) };
        })
      };
    }).filter(function (c) { return c.files.length > 0; });

    candidates.sort(function (a, b) { return b.percent - a.percent; });

    if (!candidates.length) throw new Error('No student submissions have accessible Drive folio files.');

    var high = candidates[0];
    var low = candidates[candidates.length - 1];
    var medIdx = Math.floor(candidates.length / 2);
    var med = candidates[medIdx];

    if (policy === 'grade_band') {
      var aBands = candidates.filter(function (c) { return c.grade === 'A'; });
      var cBands = candidates.filter(function (c) { return c.grade === 'C'; });
      var eBands = candidates.filter(function (c) { return ['D', 'E'].indexOf(c.grade) !== -1; });
      if (aBands.length) high = aBands[0];
      if (cBands.length) med = cBands[Math.floor(cBands.length / 2)];
      if (eBands.length) low = eBands[eBands.length - 1];
    }

    return {
      success: true,
      taskName: target.task,
      policy: policy || 'score_percentile',
      proposals: {
        high: {
          submissionRecordId: high.submissionRecordId, studentName: high.studentName,
          grade: high.grade, score: high.points, maxMarks: high.maxMarks, percent: high.percent,
          file: high.files[0], reason: 'Top-ranked submission in cohort (' + high.percent.toFixed(1) + '%, Grade ' + high.grade + ').'
        },
        med: {
          submissionRecordId: med.submissionRecordId, studentName: med.studentName,
          grade: med.grade, score: med.points, maxMarks: med.maxMarks, percent: med.percent,
          file: med.files[0], reason: 'Median representative submission (' + med.percent.toFixed(1) + '%, Grade ' + med.grade + ').'
        },
        low: {
          submissionRecordId: low.submissionRecordId, studentName: low.studentName,
          grade: low.grade, score: low.points, maxMarks: low.maxMarks, percent: low.percent,
          file: low.files[0], reason: 'Lowest score boundary submission (' + low.percent.toFixed(1) + '%, Grade ' + low.grade + ').'
        }
      },
      candidates: candidates
    };
  }

  function saveWorkSamples(workbookId, taskName, selections, destinations, confirmed) {
    if (!selections || !selections.high || !selections.med || !selections.low) {
      throw new Error('Select one HIGH, one MED, and one LOW sample.');
    }
    var target = openTask(workbookId, taskName);
    var ss = target.ss;

    if (!confirmed) {
      return {
        success: true,
        preview: true,
        selections: selections,
        destinations: destinations,
        message: 'Review sample selections and destination folders before saving.'
      };
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var user = WebAccess.requireTeacher();
      var results = [];
      var tiers = ['high', 'med', 'low'];
      for (var t = 0; t < tiers.length; t++) {
        var tier = tiers[t];
        var item = selections[tier];
        var folderId = destinations && destinations[tier + 'FolderId'];
        var label = tier.toUpperCase();
        var outcome = { tier: label, student: text(item.studentName), copied: false, fileId: '', link: '', message: '' };
        if (!folderId) outcome.message = 'No destination folder chosen; nothing copied.';
        else if (!item.file || !item.file.fileId) outcome.message = 'Student has no Drive file to copy.';
        else {
          try {
            var srcFile = DriveApp.getFileById(item.file.fileId);
            var destFolder = DriveApp.getFolderById(folderId);
            var copied = srcFile.makeCopy('[' + label + '] ' + text(item.studentName) + ' - ' + target.task + ' (Sample)', destFolder);
            outcome.copied = true; outcome.fileId = copied.getId(); outcome.link = copied.getUrl();
            outcome.message = 'Copied.';
          } catch (e) { outcome.message = 'Copy failed: ' + errString(e); }
        }
        append(ss, 'AuditLog', {
          AuditID: uid(), TaskName: target.task, Action: 'SampleSaved_' + label,
          TargetID: outcome.fileId || '', TargetURL: outcome.link, ActorEmail: user, Timestamp: now(),
          Status: outcome.copied ? 'Saved' : 'Failed',
          DetailsJSON: JSON.stringify({ tier: label, student: item.studentName, grade: item.grade, message: outcome.message })
        });
        results.push(outcome);
      }
      var copiedCount = results.filter(function (r) { return r.copied; }).length;
      return {
        success: copiedCount === tiers.length,
        partial: copiedCount > 0 && copiedCount < tiers.length,
        message: copiedCount === tiers.length ? 'Copied HIGH, MED and LOW samples into their folders.' :
          'Copied ' + copiedCount + ' of 3 samples. ' + results.filter(function (r) { return !r.copied; }).map(function (r) { return r.tier + ': ' + r.message; }).join(' '),
        results: results
      };
    } finally { lock.releaseLock(); }
  }

  /* Compatibility wrappers: the register now lives in EvaluationRegister.gs,
     which pre-fills only real workbook facts and writes a real destination. */
  function previewRegistrationPrefill(workbookId, taskName, docId) {
    var res = EvaluationRegister.preview(workbookId, taskName, docId);
    var c = res.data.context;
    return { success: true, taskName: c.taskName, docId: docId || '', destination: res.data.destination,
      proposed: { unitName: c.taskName, dates: c.dates, syllabusOutcomes: c.outcomes, gradeDistribution: c.gradeDistribution,
        className: c.className, courseCode: c.courseCode, teacherVariations: '', curriculumAdjustments: '', studentFeedbackSummary: '' } };
  }

  function saveRegistrationPrefill(workbookId, taskName, docId, fieldValues, confirmed) {
    var f = fieldValues || {};
    if (!confirmed) return { success: true, preview: true, fieldValues: f, message: 'Review field values and the destination before committing.' };
    return EvaluationRegister.commit(workbookId, taskName, {
      teacherEvaluation: f.teacherEvaluation || f.teacherVariations || '',
      teachingVariations: f.teachingVariations || (f.teacherEvaluation ? f.teacherVariations : '') || '',
      adjustments: f.adjustments || f.curriculumAdjustments || '',
      studentFeedbackSummary: f.studentFeedbackSummary || ''
    }, docId, true);
  }

  function getAuditDashboard(workbookId, params) {
    var ss = LessonGraderWeb.open(workbookId);
    var tasks = LessonGraderWeb.taskNames(ss);
    var auditLogs = records(ss, 'AuditLog');
    var grades = records(ss, 'ApprovedGrades');
    var list = [];
    for (var i = 0; i < tasks.length; i++) {
      var t = tasks[i];
      var tGrades = grades.filter(function (r) { return text(r.TaskName).trim() === t; });
      var tAudit = auditLogs.filter(function (r) { return text(r.TaskName).trim() === t; });
      var regSaved = records(ss, 'EvaluationRegister').some(function (r) { return text(r.TaskName).trim() === t && /Written|Recorded/.test(text(r.Status)); });
      function sampleSaved(tier) { return tAudit.some(function (r) { return text(r.Action) === 'SampleSaved_' + tier && text(r.Status) === 'Saved'; }); }
      var highSaved = sampleSaved('HIGH');
      var medSaved = sampleSaved('MED');
      var lowSaved = sampleSaved('LOW');
      var samplesCount = (highSaved ? 1 : 0) + (medSaved ? 1 : 0) + (lowSaved ? 1 : 0);

      list.push({
        taskName: t,
        approvedGradesCount: tGrades.length,
        registrationDocStatus: regSaved ? 'Complete' : (tGrades.length ? 'Pending Prefill' : 'Not Started'),
        sampleCounts: { high: highSaved ? 1 : 0, med: medSaved ? 1 : 0, low: lowSaved ? 1 : 0, total: samplesCount },
        auditComplete: regSaved && samplesCount === 3,
        misplacedFiles: []
      });
    }
    return {
      success: true,
      workbookName: ss.getName(),
      tasks: list
    };
  }

  function flagIncompleteSubmissions(workbookId, taskName) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var all = records(ss, 'Submissions').filter(function (r) {
      return text(r.Task).trim() === target.task && (r.CurrentOfficial === '' || yes(r.CurrentOfficial));
    });
    var files = records(ss, 'SubmissionFiles');
    var flagged = { notSubmitted: [], incomplete: [], unreadable: [], ready: [] };
    for (var i = 0; i < all.length; i++) {
      var s = all[i];
      var id = text(s.SubmissionRecordID);
      var subFiles = files.filter(function (f) { return text(f.SubmissionRecordID) === id; });
      var name = text(s.StudentName);
      var email = text(s.StudentEmail);
      if (text(s.ClassroomState) === 'CREATED' || !subFiles.length) {
        flagged.notSubmitted.push({ id: id, name: name, email: email, reason: 'No files submitted.' });
      } else if (subFiles.some(function (f) { return text(f.AIReviewStatus) === 'Inaccessible' || text(f.Limitations); })) {
        flagged.unreadable.push({ id: id, name: name, email: email, reason: 'Attached file could not be accessed.' });
      } else if (text(s.Status) === 'New') {
        flagged.ready.push({ id: id, name: name, email: email, reason: 'Ready for assessment.' });
      } else {
        flagged.incomplete.push({ id: id, name: name, email: email, reason: 'Draft in review.' });
      }
    }
    return { success: true, taskName: target.task, flags: flagged };
  }

  function draftClassroomMessages(workbookId, taskName, studentFlags, customNotes) {
    var target = openTask(workbookId, taskName);
    var flags = studentFlags || [];
    var drafts = [];
    for (var i = 0; i < flags.length; i++) {
      var s = flags[i];
      var note = customNotes || 'Please check your attached assignment files and resubmit.';
      var textMsg = 'Hi ' + s.name + ',\nRegarding your assignment for ' + target.task + ': ' + s.reason + '\n' + note;
      drafts.push({
        studentId: s.id,
        studentName: s.name,
        studentEmail: s.email,
        reason: s.reason,
        message: textMsg
      });
    }
    return { success: true, taskName: target.task, drafts: drafts };
  }

  /* Google Classroom has no API for messaging students, so nothing is sent
     from here. Drafts are recorded honestly for the teacher to copy into
     Classroom or school email. (REVIEW.md F02) */
  function sendClassroomMessages(workbookId, taskName, messages, confirmed) {
    if (!Array.isArray(messages) || !messages.length) throw new Error('No messages to record.');
    if (messages.length > 200) throw new Error('Record at most 200 message drafts at a time.');
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    if (!confirmed) {
      return { success: true, preview: true, count: messages.length, messages: messages,
        message: 'Review the drafts. Nothing will be sent: Classroom has no messaging API.' };
    }
    ensureSheet(ss, 'ClassroomMessages', MESSAGES);
    var user = WebAccess.requireTeacher();
    var links = records(ss, 'ClassroomConfig').filter(function (r) { return text(r.TaskName) === target.task && yes(r.Active); });
    var link = links[links.length - 1] || {};
    for (var i = 0; i < messages.length; i++) {
      var m = messages[i] || {};
      append(ss, 'ClassroomMessages', {
        MessageID: uid(), CourseID: text(link.CourseID), CourseWorkID: text(link.CourseWorkID),
        StudentUserID: text(m.studentId).slice(0, 200), StudentName: text(m.studentName).slice(0, 160),
        StudentEmail: text(m.studentEmail).slice(0, 180), Reason: text(m.reason).slice(0, 300),
        MessageText: text(m.message).slice(0, 4000), Status: 'Drafted — not sent', SentAt: '', SentBy: user
      });
    }
    return { success: true, sent: 0, recorded: messages.length,
      message: 'Recorded ' + messages.length + ' message draft(s) in ClassroomMessages. Nothing was sent — copy them into Classroom or email.' };
  }

  function getClassTableData(workbookId, taskName, filter) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var all = records(ss, 'Submissions').filter(function (r) {
      return text(r.Task).trim() === target.task && (r.CurrentOfficial === '' || yes(r.CurrentOfficial));
    });
    var grades = records(ss, 'ApprovedGrades').filter(function (r) { return text(r.TaskName).trim() === target.task; });
    var profile = RubricEngine.getActiveProfile(ss, target.task);

    var rows = all.map(function (s) {
      var g = grades.filter(function (r) { return text(r.SubmissionRecordID) === text(s.SubmissionRecordID); })[0];
      return {
        submissionRecordId: text(s.SubmissionRecordID),
        studentName: text(s.StudentName),
        studentEmail: text(s.StudentEmail),
        className: text(s.Class),
        taskName: target.task,
        assessmentType: text(s.AssessmentType) || (g && text(g.AssessmentType)) || 'Summative',
        rawScore: s.RawScore != null && s.RawScore !== '' ? toNumber(s.RawScore) : (g && g.RawScore != null && g.RawScore !== '' ? toNumber(g.RawScore) : null),
        deductionPoints: s.DeductionPoints != null && s.DeductionPoints !== '' ? toNumber(s.DeductionPoints) : (g && g.DeductionPoints != null && g.DeductionPoints !== '' ? toNumber(g.DeductionPoints) : 0),
        deductionReason: text(s.DeductionReason || (g && g.DeductionReason) || ''),
        finalScore: s.FinalScore != null && s.FinalScore !== '' ? toNumber(s.FinalScore) : (g && g.FinalScore != null && g.FinalScore !== '' ? toNumber(g.FinalScore) : (g ? toNumber(g.Points) : null)),
        grade: g ? text(g.Grade) : 'Not Approved',
        percent: g ? toNumber(g.Percent) : null,
        status: text(s.Status)
      };
    });

    return {
      success: true,
      taskName: target.task,
      profile: profile ? RubricEngine.profileSummary(profile) : null,
      rows: rows
    };
  }

  /* Writes the task's current results to ClassTracking. Course code, year and
     stage come from the workbook; the student key is the real StudentUserID. (F27) */
  function syncClassTable(workbookId, taskName, classMeta) {
    var target = openTask(workbookId, taskName);
    var ss = target.ss;
    var actor = WebAccess.requireTeacher();
    ensureSheet(ss, 'ClassTracking', HEADERS.ClassTracking);
    var meta = classMeta || {};
    var parsed = EvaluationRegister.parseWorkbookName(ss.getName());
    var subs = records(ss, 'Submissions').filter(function (r) { return text(r.Task).trim() === target.task && isCurrent(r); });
    var grades = records(ss, 'ApprovedGrades').filter(function (r) { return text(r.TaskName).trim() === target.task; });
    var links = records(ss, 'ClassroomConfig').filter(function (r) { return text(r.TaskName) === target.task; });
    var className = text(meta.className).trim() || text((links[links.length - 1] || {}).CourseName) || text((subs[0] || {}).Class);
    var termPair = text(meta.termPair).trim().slice(0, 40);
    return withLock(function () {
      var existing = records(ss, 'ClassTracking');
      var updated = 0, added = 0;
      subs.forEach(function (s) {
        var g = grades.filter(function (r) { return text(r.SubmissionRecordID) === text(s.SubmissionRecordID); })[0];
        var studentKey = text(s.StudentUserID) || text(s.SubmissionRecordID);
        var prior = existing.filter(function (r) {
          return text(r.TaskName) === target.task && (text(r.StudentUserID) === studentKey || text(r.StudentUserID) === text(s.SubmissionRecordID));
        })[0];
        var fields = {
          TrackingID: prior ? prior.TrackingID : uid(), Year: parsed.year, Stage: deriveStageFromCourseCode(parsed.courseCode) || '',
          CourseCode: parsed.courseCode, ClassID: className, TermPair: termPair || (prior ? text(prior.TermPair) : ''),
          StudentUserID: studentKey, StudentName: text(s.StudentName), TaskName: target.task,
          AssessmentType: text(s.AssessmentType) || (g && text(g.AssessmentType)) || 'Summative',
          RawScore: g ? g.RawScore : '', FinalScore: g ? (g.FinalScore !== '' ? g.FinalScore : g.Points) : '',
          MaxMarks: g ? g.MaxMarks : '', Percent: g ? g.Percent : '', Grade: g ? text(g.Grade) : 'Not approved',
          OutcomesJSON: g ? text(g.OutcomesJSON) : '', Status: text(s.Status), UpdatedAt: now()
        };
        if (prior) { patch(ss, 'ClassTracking', prior._row, fields); updated++; }
        else { append(ss, 'ClassTracking', fields); existing.push(fields); added++; }
      });
      return { success: true, message: 'Class tracking synced by ' + actor + ': ' + added + ' added, ' + updated + ' updated.' };
    });
  }

  return {
    HEADERS: HEADERS, EXTRA: EXTRA, prepare: prepare, upgrade: upgrade, ready: ready, listTask: listTask, detail: detail,
    saveDraft: saveDraft, buildFeedback: buildFeedback, approve: approve, lockAssessment: lockAssessment, unlock: unlock,
    reassess: reassess, records: records, append: append, patch: patch, ensureSheet: ensureSheet,
    openTask: openTask, event: event, lastRevision: lastRevision, assess: assess, isCurrent: isCurrent, selectSubmission: selectSubmission, withLock: withLock, previewRubric: previewRubric, activateRubric: activateRubric, rubricSource: rubricSource,
    runAi: runAi, report: report, importTask: importTask,
    addManualSubmission: addManualSubmission,
    bulkGrade: bulkGrade, verifyGradingComplete: verifyGradingComplete,
    setGradingComplete: setGradingComplete, suggestWorkSamples: suggestWorkSamples,
    saveWorkSamples: saveWorkSamples, previewRegistrationPrefill: previewRegistrationPrefill,
    saveRegistrationPrefill: saveRegistrationPrefill, getAuditDashboard: getAuditDashboard,
    flagIncompleteSubmissions: flagIncompleteSubmissions,
    draftClassroomMessages: draftClassroomMessages,
    sendClassroomMessages: sendClassroomMessages,
    getClassTableData: getClassTableData, syncClassTable: syncClassTable
  };
})();

/* Web RPC surface: every action validates identity AND workbook membership. */
function apiWebEnableGrading(workbookId) {
  WebAccess.requireTeacher();
  return WebGrading.prepare(LessonGraderWeb.open(workbookId), true);
}
function apiWebTaskDashboard(workbookId, taskName) {
  WebAccess.requireTeacher();
  return WebGrading.listTask(workbookId, taskName);
}
function apiWebGradeDetail(workbookId, taskName, submissionId) {
  WebAccess.requireTeacher();
  return WebGrading.detail(workbookId, taskName, submissionId);
}
function apiWebSaveDraft(workbookId, taskName, submissionId, criteria, feedback, revision) {
  WebAccess.requireTeacher();
  return WebGrading.saveDraft(workbookId, taskName, submissionId, criteria, feedback, revision, 'Teacher');
}
function apiWebBuildFeedback(workbookId, taskName, submissionId, criteria, generalObservation) {
  WebAccess.requireTeacher();
  return WebGrading.buildFeedback(workbookId, taskName, submissionId, criteria, generalObservation);
}
function apiWebEvidenceVersion(workbookId, taskName, submissionId, versionId) {
  WebAccess.requireTeacher();
  var target = WebGrading.openTask(workbookId, taskName);
  WebGrading.selectSubmission(target.ss, target.task, submissionId);
  return { success: true, data: EvidenceService.versionText(target.ss, submissionId, versionId) };
}
function apiWebApprove(workbookId, taskName, submissionId, revision) {
  WebAccess.requireTeacher();
  return WebGrading.approve(workbookId, taskName, submissionId, revision);
}
function apiWebLock(workbookId, taskName, submissionId) {
  WebAccess.requireTeacher();
  return WebGrading.lockAssessment(workbookId, taskName, submissionId);
}
function apiWebUnlock(workbookId, taskName, submissionId, reason, confirmation) {
  WebAccess.requireTeacher();
  return WebGrading.unlock(workbookId, taskName, submissionId, reason, confirmation);
}
function apiWebReassess(workbookId, taskName, submissionId) {
  WebAccess.requireTeacher();
  return WebGrading.reassess(workbookId, taskName, submissionId);
}
function apiWebPreviewRubric(workbookId, taskName, rubricId) {
  WebAccess.requireTeacher();
  return WebGrading.previewRubric(workbookId, taskName, rubricId);
}
function apiWebRubricSource(workbookId, taskName) {
  WebAccess.requireTeacher();
  return WebGrading.rubricSource(workbookId, taskName);
}
function apiWebGetGradeScale(workbookId) {
  WebAccess.requireTeacher();
  return { success: true, data: RubricEngine.getGradeScale(LessonGraderWeb.open(workbookId)) };
}
function apiWebActivateRubric(workbookId, taskName, name, sourceId, rows, customWeights, confirmed) {
  WebAccess.requireTeacher();
  return WebGrading.activateRubric(workbookId, taskName, name, sourceId, rows, customWeights, confirmed);
}
function apiWebRunAi(workbookId, taskName, submissionId, revision) {
  WebAccess.requireTeacher();
  return WebGrading.runAi(workbookId, taskName, submissionId, revision);
}
function apiWebReport(workbookId, taskName, submissionId) {
  WebAccess.requireTeacher();
  return WebGrading.report(workbookId, taskName, submissionId);
}
function apiWebImportTask(workbookId, taskName) {
  WebAccess.requireTeacher();
  return WebGrading.importTask(workbookId, taskName);
}
function apiWebAddManualSubmission(workbookId, taskName, name, email, className) {
  WebAccess.requireTeacher();
  return WebGrading.addManualSubmission(workbookId, taskName, name, email, className);
}
function apiWebListTeacherCourses() { WebAccess.requireTeacher(); return ClassroomService.listTeacherCourses(); }
function apiWebListCourseWork(courseId) { WebAccess.requireTeacher(); return ClassroomService.listCourseWork(courseId); }
function apiWebUploadRubric(name, document) {
  WebAccess.requireTeacher();
  if (!document || !document.base64 || document.base64.length > 11 * 1024 * 1024) {
    return { success: false, message: 'Choose a PDF, Markdown or text rubric under 8 MB.' };
  }
  return apiCreateRubricFromDocument(name, document);
}
function apiWebBulkGradeQueue(workbookId, taskName, submissionIds, options) {
  WebAccess.requireTeacher();
  return WebGrading.bulkGrade(workbookId, taskName, submissionIds, options);
}
function apiWebVerifyGradingComplete(workbookId, taskName) {
  WebAccess.requireTeacher();
  return WebGrading.verifyGradingComplete(workbookId, taskName);
}
function apiWebSetGradingComplete(workbookId, taskName, confirmed) {
  WebAccess.requireTeacher();
  return WebGrading.setGradingComplete(workbookId, taskName, confirmed);
}
function apiWebSuggestWorkSamples(workbookId, taskName, policy) {
  WebAccess.requireTeacher();
  return WebGrading.suggestWorkSamples(workbookId, taskName, policy);
}
function apiWebSaveWorkSamples(workbookId, taskName, selections, destinations, confirmed) {
  WebAccess.requireTeacher();
  return WebGrading.saveWorkSamples(workbookId, taskName, selections, destinations, confirmed);
}
function apiWebPreviewRegistrationPrefill(workbookId, taskName, docId, classMeta) {
  WebAccess.requireTeacher();
  return WebGrading.previewRegistrationPrefill(workbookId, taskName, docId, classMeta);
}
function apiWebSaveRegistrationPrefill(workbookId, taskName, docId, fieldValues, confirmed) {
  WebAccess.requireTeacher();
  return WebGrading.saveRegistrationPrefill(workbookId, taskName, docId, fieldValues, confirmed);
}
function apiWebGetAuditDashboard(workbookId, params) {
  WebAccess.requireTeacher();
  return WebGrading.getAuditDashboard(workbookId, params);
}
function apiWebFlagIncompleteSubmissions(workbookId, taskName) {
  WebAccess.requireTeacher();
  return WebGrading.flagIncompleteSubmissions(workbookId, taskName);
}
function apiWebDraftClassroomMessages(workbookId, taskName, studentFlags, customNotes) {
  WebAccess.requireTeacher();
  return WebGrading.draftClassroomMessages(workbookId, taskName, studentFlags, customNotes);
}
function apiWebSendClassroomMessages(workbookId, taskName, messages, confirmed) {
  WebAccess.requireTeacher();
  return WebGrading.sendClassroomMessages(workbookId, taskName, messages, confirmed);
}
function apiWebGetClassTableData(workbookId, taskName, filter) {
  WebAccess.requireTeacher();
  return WebGrading.getClassTableData(workbookId, taskName, filter);
}
function apiWebSyncClassTable(workbookId, taskName, classMeta) {
  WebAccess.requireTeacher();
  return WebGrading.syncClassTable(workbookId, taskName, classMeta);
}
