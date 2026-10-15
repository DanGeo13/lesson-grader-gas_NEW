/* ============================================================================
 * AiQueue.gs — queued AI drafting that survives outages and the 6-minute
 * execution limit. (SC28, SC48, SC49)
 *
 * One AiQueue sheet per workbook. A job drafts one submission.
 *   Queued → Running → Done | Failed | Skipped | Cancelled
 * - Enqueue is idempotent: a submission with a Queued/Running job is not queued again.
 * - A job records the assessment revision at enqueue; if the teacher saves in the
 *   meantime the job is Skipped (teacher work always wins).
 * - Jobs are claimed under a short lock, processed OUTSIDE the lock, and a
 *   Running claim older than 15 minutes is treated as abandoned and retried.
 * - Provider outage → job returns to Queued (attempt counted); after 5
 *   attempts it is Failed with the reason. Never approves, never publishes.
 * ============================================================================ */
var AiQueue = (function () {
  var SHEET = 'AiQueue';
  var HEADERS = ['JobID', 'TaskName', 'SubmissionRecordID', 'StudentName', 'Status', 'Attempts', 'ExpectedRevision',
    'ReplaceExisting', 'CreatedAt', 'CreatedBy', 'ClaimedAt', 'UpdatedAt', 'Message'];
  var ACTIVE = { Queued: true, Running: true };
  var STALE_MS = 15 * 60 * 1000;
  var MAX_ATTEMPTS = 5;
  var DEFAULT_BUDGET_MS = 4.5 * 60 * 1000;

  function text(v) { return v == null ? '' : String(v); }
  function now() { return new Date().toISOString(); }
  function sheet(ss) { return WebGrading.ensureSheet(ss, SHEET, HEADERS); }
  function jobs(ss) { return ss.getSheetByName(SHEET) ? WebGrading.records(ss, SHEET) : []; }

  /** Queue drafts for selected submissions. Returns counts; never duplicates active jobs. */
  function enqueue(workbookId, taskName, submissionIds, options) {
    var opt = options || {};
    var target = WebGrading.openTask(workbookId, taskName);
    var ss = target.ss;
    var actor = opt.actor || WebAccess.requireTeacher();
    var ids = Array.isArray(submissionIds) ? submissionIds.map(text).filter(Boolean) : [];
    if (!ids.length) throw new Error('Select at least one student.');
    if (ids.length > 200) throw new Error('Queue at most 200 students at a time.');
    return WebGrading.withLock(function () {
      sheet(ss);
      var existing = jobs(ss);
      var subs = WebGrading.records(ss, 'Submissions');
      var result = { queued: 0, alreadyQueued: 0, skipped: [] };
      ids.forEach(function (id) {
        var sub = subs.filter(function (s) { return text(s.SubmissionRecordID) === id && text(s.Task).trim() === target.task; })[0];
        if (!sub) { result.skipped.push({ id: id, reason: 'Not in this task.' }); return; }
        if (['Approved', 'Locked', 'Superseded'].indexOf(text(sub.Status)) !== -1 || !WebGrading.isCurrent(sub)) {
          result.skipped.push({ id: id, name: text(sub.StudentName), reason: 'Approved, locked or historical.' }); return;
        }
        if (existing.some(function (j) { return text(j.SubmissionRecordID) === id && ACTIVE[text(j.Status)]; })) { result.alreadyQueued++; return; }
        WebGrading.append(ss, SHEET, { JobID: Utilities.getUuid(), TaskName: target.task, SubmissionRecordID: id,
          StudentName: text(sub.StudentName), Status: 'Queued', Attempts: 0, ExpectedRevision: WebGrading.lastRevision(ss, id),
          ReplaceExisting: opt.replaceExisting === true, CreatedAt: now(), CreatedBy: actor, UpdatedAt: now(), Message: '' });
        result.queued++;
      });
      result.message = result.queued + ' queued' + (result.alreadyQueued ? ', ' + result.alreadyQueued + ' already queued' : '') +
        (result.skipped.length ? ', ' + result.skipped.length + ' skipped' : '') + '. Drafts appear as each one finishes.';
      return Object.assign({ success: true }, result);
    });
  }

  function claim(ss, taskName) {
    return WebGrading.withLock(function () {
      var list = jobs(ss);
      var t = Date.now();
      for (var i = 0; i < list.length; i++) {
        var j = list[i];
        if (taskName && text(j.TaskName) !== taskName) continue;
        var st = text(j.Status);
        var stale = st === 'Running' && (t - new Date(text(j.ClaimedAt)).getTime() > STALE_MS);
        if (st === 'Queued' || stale) {
          WebGrading.patch(ss, SHEET, j._row, { Status: 'Running', ClaimedAt: now(), UpdatedAt: now(),
            Attempts: (Number(j.Attempts) || 0) + (stale ? 1 : 0), Message: stale ? 'Retrying an abandoned run.' : '' });
          j.Attempts = (Number(j.Attempts) || 0) + (stale ? 1 : 0);
          return j;
        }
      }
      return null;
    });
  }

  function finish(ss, job, fields) {
    WebGrading.withLock(function () { WebGrading.patch(ss, SHEET, job._row, Object.assign({ UpdatedAt: now() }, fields)); });
  }

  function runOne(workbookId, ss, job, actor) {
    var id = text(job.SubmissionRecordID);
    try {
      var sub = WebGrading.selectSubmission(ss, text(job.TaskName), id);
      if (['Approved', 'Locked', 'Superseded'].indexOf(text(sub.Status)) !== -1) {
        finish(ss, job, { Status: 'Skipped', Message: 'Approved or locked before the draft ran.' }); return 'Skipped';
      }
      if (WebGrading.lastRevision(ss, id) !== text(job.ExpectedRevision)) {
        finish(ss, job, { Status: 'Skipped', Message: 'The teacher changed this assessment after it was queued; teacher work was kept.' }); return 'Skipped';
      }
      var res = AiGateway.draft(workbookId, text(job.TaskName), id, text(job.ExpectedRevision),
        { actor: actor, replaceExisting: job.ReplaceExisting === true || text(job.ReplaceExisting).toLowerCase() === 'true' });
      if (res.success) { finish(ss, job, { Status: 'Done', Message: res.message.slice(0, 400) }); return 'Done'; }
      var attempts = (Number(job.Attempts) || 0) + 1;
      if (res.unavailable && attempts < MAX_ATTEMPTS) {
        finish(ss, job, { Status: 'Queued', Attempts: attempts, Message: 'Waiting for AI service: ' + text(res.message).slice(0, 300) }); return 'Waiting';
      }
      finish(ss, job, { Status: 'Failed', Attempts: attempts, Message: text(res.message).slice(0, 400) }); return 'Failed';
    } catch (e) {
      finish(ss, job, { Status: 'Failed', Attempts: (Number(job.Attempts) || 0) + 1, Message: text(e && e.message).slice(0, 400) });
      return 'Failed';
    }
  }

  /**
   * Processes queued jobs for one workbook until the time budget is used.
   * Stops early when the AI service is unavailable (remaining jobs stay queued).
   */
  function process(workbookId, taskName, budgetMs, actor) {
    var ss = LessonGraderWeb.open(workbookId);
    var who = actor || WebAccess.requireTeacher();
    var started = Date.now();
    var budget = Math.min(Number(budgetMs) || DEFAULT_BUDGET_MS, DEFAULT_BUDGET_MS);
    var counts = { Done: 0, Skipped: 0, Failed: 0, Waiting: 0 };
    if (!ss.getSheetByName(SHEET)) return { success: true, processed: 0, counts: counts, remaining: 0, message: 'Nothing queued.' };
    while (Date.now() - started < budget) {
      var job = claim(ss, taskName ? text(taskName).trim() : '');
      if (!job) break;
      var outcome = runOne(workbookId, ss, job, who);
      counts[outcome]++;
      if (outcome === 'Waiting') break; // service down: stop hammering it
    }
    var left = jobs(ss).filter(function (j) { return (!taskName || text(j.TaskName) === text(taskName).trim()) && ACTIVE[text(j.Status)]; }).length;
    return { success: true, counts: counts, remaining: left,
      message: counts.Done + ' draft(s) ready, ' + counts.Skipped + ' skipped, ' + counts.Failed + ' failed' +
        (counts.Waiting ? '; AI service unavailable, ' + left + ' still queued' : left ? '; ' + left + ' still queued' : '') + '.' };
  }

  function status(workbookId, taskName) {
    var target = WebGrading.openTask(workbookId, taskName);
    var list = jobs(target.ss).filter(function (j) { return text(j.TaskName) === target.task; });
    var counts = { Queued: 0, Running: 0, Done: 0, Failed: 0, Skipped: 0, Cancelled: 0 };
    list.forEach(function (j) { counts[text(j.Status)] = (counts[text(j.Status)] || 0) + 1; });
    return { success: true, data: { counts: counts, jobs: list.slice(-100).reverse().map(function (j) {
      return { jobId: text(j.JobID), submissionRecordId: text(j.SubmissionRecordID), studentName: text(j.StudentName), status: text(j.Status),
        attempts: Number(j.Attempts) || 0, updatedAt: text(j.UpdatedAt), message: text(j.Message) };
    }) } };
  }

  function cancel(workbookId, taskName, jobIds) {
    var target = WebGrading.openTask(workbookId, taskName);
    var ids = (jobIds || []).map(text);
    return WebGrading.withLock(function () {
      var n = 0;
      jobs(target.ss).forEach(function (j) {
        if (ids.indexOf(text(j.JobID)) !== -1 && text(j.Status) === 'Queued') { WebGrading.patch(target.ss, SHEET, j._row, { Status: 'Cancelled', UpdatedAt: now() }); n++; }
      });
      return { success: true, message: n + ' job(s) cancelled.' };
    });
  }

  return { HEADERS: HEADERS, enqueue: enqueue, process: process, status: status, cancel: cancel, MAX_ATTEMPTS: MAX_ATTEMPTS };
})();

function apiWebQueueAiDrafts(workbookId, taskName, submissionIds, replaceExisting) {
  WebAccess.requireTeacher();
  return AiQueue.enqueue(workbookId, taskName, submissionIds, { replaceExisting: replaceExisting === true });
}
function apiWebProcessAiQueue(workbookId, taskName) {
  WebAccess.requireTeacher();
  return AiQueue.process(workbookId, taskName, 4 * 60 * 1000);
}
function apiWebAiQueueStatus(workbookId, taskName) { WebAccess.requireTeacher(); return AiQueue.status(workbookId, taskName); }
function apiWebCancelAiJobs(workbookId, taskName, jobIds) { WebAccess.requireTeacher(); return AiQueue.cancel(workbookId, taskName, jobIds); }
