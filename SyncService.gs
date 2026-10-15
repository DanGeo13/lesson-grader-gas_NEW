/* ============================================================================
 * SyncService.gs — automatic and manual Classroom synchronisation for every
 * workbook in the teacher's library, plus trigger management. (SC08–SC12)
 *
 * Automatic (hourly trigger, runs as the teacher who installed it):
 *   - only links WITH a due date, and only after that due date has passed;
 *   - first run marks the link AutoImported; later hours keep syncing, so late
 *     and resubmitted work arrives without teacher action.
 * Manual: "Sync now" (apiWebImportTask) works for every link, dated or not,
 *   whether or not the trigger is installed.
 * Every sync is idempotent and never changes marks, feedback or approvals.
 * ============================================================================ */
var SyncService = (function () {
  var SYNC_HANDLER = 'lessonGraderHourlySync';
  var QUEUE_HANDLER = 'lessonGraderAiQueueTick';
  var BUDGET_MS = 5 * 60 * 1000;

  function text(v) { return v == null ? '' : String(v); }
  function yes(v) { return v === true || text(v).toLowerCase() === 'true'; }

  /**
   * Classroom due dates are UTC. With no due time we use 23:59 in Australia/Sydney
   * (end of the school day locally) rather than 23:59 UTC (≈10am next day). (F46)
   */
  function classroomDueIso(dueDate, dueTime) {
    if (!dueDate || !dueDate.year || !dueDate.month || !dueDate.day) return '';
    if (dueTime && (dueTime.hours != null || dueTime.minutes != null)) {
      return new Date(Date.UTC(dueDate.year, dueDate.month - 1, dueDate.day, dueTime.hours || 0, dueTime.minutes || 0)).toISOString();
    }
    var local = dueDate.year + '-' + ('0' + dueDate.month).slice(-2) + '-' + ('0' + dueDate.day).slice(-2) + ' 23:59';
    return Utilities.parseDate(local, AppConfig.TIME_ZONE, 'yyyy-MM-dd HH:mm').toISOString();
  }

  /** Links eligible for automatic import now. Pure given rows and a clock. */
  function dueLinks(configRows, nowMs) {
    return (configRows || []).filter(function (r) {
      if (!yes(r.Active) || !text(r.TaskName).trim() || !text(r.CourseID) || !text(r.CourseWorkID)) return false;
      var due = new Date(text(r.DueDate)).getTime();
      return isFinite(due) && due <= nowMs;
    });
  }

  function recordLinkResult(ss, row, fields) {
    try { WebGrading.ensureSheet(ss, 'ClassroomConfig', ['LastSyncedAt', 'LastSyncResult']); WebGrading.patch(ss, 'ClassroomConfig', row._row, fields); }
    catch (e) { console.error('Could not record sync result: ' + e); }
  }

  /** Trigger handler body. Returns a summary for logging/tests. */
  function runHourly(nowMs) {
    var actor = WebAccess.requireTriggerOwner();
    var started = Date.now();
    var t = nowMs || Date.now();
    var listing = LessonGraderWeb.list();
    var summary = { workbooks: 0, links: 0, imported: 0, errors: [] };
    if (!listing.success) { summary.errors.push(listing.message); return summary; }
    for (var i = 0; i < listing.data.length; i++) {
      if (Date.now() - started > BUDGET_MS) { summary.errors.push('Time budget reached; remaining workbooks sync next hour.'); break; }
      var wb = listing.data[i];
      try {
        var ss = LessonGraderWeb.open(wb.id);
        if (!WebGrading.ready(ss)) continue;
        summary.workbooks++;
        dueLinks(WebGrading.records(ss, 'ClassroomConfig'), t).forEach(function (link) {
          summary.links++;
          try {
            var res = WebGrading.importTask(wb.id, text(link.TaskName), { actor: actor, automatic: true });
            summary.imported += (res.stats.newCount || 0) + (res.stats.updated || 0);
            var fields = { LastSyncedAt: new Date().toISOString(), LastSyncResult: res.message.slice(0, 300) };
            if (!yes(link.AutoImported)) fields.AutoImported = true;
            recordLinkResult(ss, link, fields);
          } catch (e) {
            summary.errors.push(wb.name + ' / ' + text(link.TaskName) + ': ' + (e && e.message));
            recordLinkResult(ss, link, { LastSyncedAt: new Date().toISOString(), LastSyncResult: 'Failed: ' + text(e && e.message).slice(0, 280) });
          }
        });
      } catch (err) {
        summary.errors.push(wb.name + ': ' + (err && err.message));
      }
    }
    return summary;
  }

  /** Processes queued AI drafts across the library within one execution budget. */
  function runQueue() {
    var actor = WebAccess.requireTriggerOwner();
    var started = Date.now();
    var listing = LessonGraderWeb.list();
    var done = 0;
    if (!listing.success) return { done: 0 };
    for (var i = 0; i < listing.data.length; i++) {
      var remaining = BUDGET_MS - (Date.now() - started);
      if (remaining < 30000) break;
      try {
        var ss = LessonGraderWeb.open(listing.data[i].id);
        if (!ss.getSheetByName('AiQueue')) continue;
        var res = AiQueue.process(listing.data[i].id, '', remaining - 15000, actor);
        done += res.counts.Done;
        if (res.counts.Waiting) break; // provider down: try again next tick
      } catch (e) { console.error('Queue tick: ' + e); }
    }
    return { done: done };
  }

  function triggerStatus() {
    WebAccess.requireTeacher();
    var triggers = ScriptApp.getProjectTriggers();
    function has(h) { return triggers.some(function (t) { return t.getHandlerFunction() === h; }); }
    return { success: true, data: { hourlySync: has(SYNC_HANDLER), aiQueue: has(QUEUE_HANDLER),
      legacyAutoImport: has('checkAndAutoImportDueAssignments'),
      note: 'Triggers run as the teacher who installed them and only see that teacher\'s triggers here.' } };
  }

  function installTriggers() {
    WebAccess.requireTeacher();
    var triggers = ScriptApp.getProjectTriggers();
    var created = [];
    if (!triggers.some(function (t) { return t.getHandlerFunction() === SYNC_HANDLER; })) {
      ScriptApp.newTrigger(SYNC_HANDLER).timeBased().everyHours(1).create(); created.push('hourly Classroom sync');
    }
    if (!triggers.some(function (t) { return t.getHandlerFunction() === QUEUE_HANDLER; })) {
      ScriptApp.newTrigger(QUEUE_HANDLER).timeBased().everyMinutes(15).create(); created.push('15-minute AI queue');
    }
    return { success: true, message: created.length ? 'Installed: ' + created.join(', ') + '.' : 'Triggers were already installed.', data: triggerStatus().data };
  }

  function removeTriggers() {
    WebAccess.requireTeacher();
    var n = 0;
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if ([SYNC_HANDLER, QUEUE_HANDLER].indexOf(t.getHandlerFunction()) !== -1) { ScriptApp.deleteTrigger(t); n++; }
    });
    return { success: true, message: n + ' trigger(s) removed. Manual "Sync now" still works.', data: triggerStatus().data };
  }

  return { SYNC_HANDLER: SYNC_HANDLER, QUEUE_HANDLER: QUEUE_HANDLER, classroomDueIso: classroomDueIso, dueLinks: dueLinks,
    runHourly: runHourly, runQueue: runQueue, triggerStatus: triggerStatus, installTriggers: installTriggers, removeTriggers: removeTriggers };
})();

/* Trigger handlers (names are referenced by installed triggers — do not rename). */
function lessonGraderHourlySync() { return SyncService.runHourly(); }
function lessonGraderAiQueueTick() { return SyncService.runQueue(); }

function apiWebTriggerStatus() { return SyncService.triggerStatus(); }
function apiWebInstallTriggers() { return SyncService.installTriggers(); }
function apiWebRemoveTriggers() { return SyncService.removeTriggers(); }
