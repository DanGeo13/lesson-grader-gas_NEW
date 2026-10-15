/* ============================================================================
 * HealthCheck.gs — read-only self-test. Never writes to a workbook, never
 * publishes, never changes a grade and never returns a secret.
 * Each check: { area, check, status: 'pass'|'warn'|'fail'|'info', detail }.
 * ============================================================================ */
var HealthCheck = (function () {
  function text(v) { return v == null ? '' : String(v); }
  function yes(v) { return v === true || text(v).toLowerCase() === 'true'; }

  function run(workbookId) {
    var user = WebAccess.requireTeacher();
    var P = AppConfig.PROPS;
    var out = [];
    function add(area, check, status, detail) { out.push({ area: area, check: check, status: status, detail: detail || '' }); }

    /* Configuration */
    add('Configuration', 'Teacher allowlist', AppConfig.emailList(P.ALLOWLIST).length ? 'pass' : 'fail',
      AppConfig.emailList(P.ALLOWLIST).length + ' teacher account(s) allowed.');
    add('Configuration', 'Authorised unlock list', AppConfig.emailList(P.UNLOCK_ALLOWLIST).length ? 'pass' : 'warn',
      AppConfig.emailList(P.UNLOCK_ALLOWLIST).length ? 'Unlock is restricted to the configured accounts.' :
        'LESSON_GRADER_UNLOCK_EMAILS is empty: final-locked assessments cannot be unlocked by anyone.');
    var tz = '';
    try { tz = Session.getScriptTimeZone(); } catch (e) { tz = ''; }
    add('Configuration', 'Time zone', tz === AppConfig.TIME_ZONE ? 'pass' : 'warn', 'Script time zone: ' + (tz || 'unknown') + ' (expected ' + AppConfig.TIME_ZONE + ').');
    add('Configuration', 'Launcher embedding', 'info', AppConfig.flag(P.ALLOW_EMBED) ? 'Embedding allowed (LESSON_GRADER_ALLOW_EMBED=true).' : 'Embedding blocked (default). The launcher opens the app in a new tab.');

    /* AI */
    var providers = WebAccess.providerStatus();
    var settings = AppConfig.readSettings().aiProviders;
    var health = AiGateway.readHealth();
    function aiLine(name, configured) {
      var h = health[name];
      var status = !configured ? 'warn' : h && h.ok === false ? 'warn' : 'pass';
      add('AI', name.toUpperCase() + ' provider', status, (configured ? 'Configured' : 'Not configured') +
        (h ? '; last check ' + h.at.slice(0, 16).replace('T', ' ') + ' UTC: ' + (h.ok ? 'OK' : h.message) : '; not yet checked') + '.');
    }
    aiLine('gemini', providers.gemini.configured);
    aiLine('qwen', providers.qwen.configured);
    aiLine('jev', providers.jev.configured);
    var draftReady = settings.draftProvider === 'gemini' ? providers.gemini.configured : providers.jev.configured;
    add('AI', 'Drafting provider', draftReady ? 'pass' : 'warn', 'Drafting uses ' + settings.draftProvider +
      (draftReady ? '.' : ', which is not configured. Teachers can still mark manually.'));
    add('AI', 'Active prompt', 'info', 'Prompt version ' + PromptRegistry.active().id + '.');

    /* Triggers */
    var triggers = [];
    try { triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }); } catch (e) { triggers = []; }
    add('Sync', 'Hourly Classroom sync trigger', triggers.indexOf(SyncService.SYNC_HANDLER) !== -1 ? 'pass' : 'warn',
      triggers.indexOf(SyncService.SYNC_HANDLER) !== -1 ? 'Installed for ' + user + '.' : 'Not installed for ' + user + '. Due-dated assignments will not import automatically; "Sync now" still works.');
    add('Sync', 'AI queue trigger', triggers.indexOf(SyncService.QUEUE_HANDLER) !== -1 ? 'pass' : 'warn',
      triggers.indexOf(SyncService.QUEUE_HANDLER) !== -1 ? 'Installed.' : 'Not installed: queued drafts run only while a teacher presses "Process queue".');

    /* Classroom access (read-only) */
    try {
      var courses = ClassroomService.listTeacherCourses();
      add('Classroom', 'Classroom access', courses.success ? 'pass' : 'fail', courses.success ? courses.courses.length + ' active class(es) visible.' : courses.message);
    } catch (e) { add('Classroom', 'Classroom access', 'fail', text(e && e.message)); }

    /* Grading rules */
    add('Grading rules', 'Teacher-selected criterion level', 'info', 'Not implemented by design: the rule for how a chosen level interacts with observables is unresolved (REVIEW.md §4). Levels are derived from observables.');
    add('Grading rules', 'B and D contributions', 'warn', 'Not confirmed by school policy. Current defaults B 87.5%, D 57.5% remain in use for existing workbooks.');

    if (workbookId) {
      try {
        var ss = LessonGraderWeb.open(workbookId);
        add('Workbook', 'Opened', 'pass', ss.getName());
        add('Workbook', 'Grading tables', WebGrading.ready(ss) ? 'pass' : 'fail', WebGrading.ready(ss) ? 'All required sheets and columns present.' : 'Run "Enable grading" (takes a backup first).');
        var scale = RubricEngine.getGradeScale(ss);
        var thresholds = RubricEngine.validateThresholds(scale.bands);
        add('Grading rules', 'Grade boundaries', thresholds.ok ? 'pass' : 'fail', scale.bands.map(function (b) { return b.letter + ' ≥ ' + b.min; }).join(', ') + (thresholds.ok ? '' : ' — ' + thresholds.errors.join(' ')));
        var ref = RubricEngine.referencePointReport(scale.weights);
        add('Grading rules', 'Reference points A/C/E (SC19)', ref.matchesTarget ? 'pass' : 'warn', ref.message);
        LessonGraderWeb.taskNames(ss).forEach(function (task) {
          var profile = RubricEngine.getActiveProfile(ss, task);
          add('Task: ' + task, 'Active rubric', profile ? 'pass' : 'warn', profile ? profile.name + ' (' + profile.criteria.length + ' criteria, ' + profile.totalMaxMarks + ' marks).' : 'No rubric activated.');
          var links = WebGrading.records(ss, 'ClassroomConfig').filter(function (r) { return text(r.TaskName) === task && yes(r.Active); });
          if (!links.length) { add('Task: ' + task, 'Classroom link', 'info', 'Not linked (manual students only; publishing unavailable).'); return; }
          var link = links[links.length - 1];
          add('Task: ' + task, 'Classroom link', 'pass', text(link.AssignmentTitle) + (text(link.DueDate) ? ' · due ' + text(link.DueDate).slice(0, 10) : ' · no due date (manual sync)') +
            (text(link.LastSyncedAt) ? ' · last auto-sync ' + text(link.LastSyncedAt).slice(0, 16) : ''));
          try {
            var cw = Classroom.Courses.CourseWork.get(text(link.CourseID), text(link.CourseWorkID));
            var max = cw.maxPoints == null ? 0 : Number(cw.maxPoints);
            add('Task: ' + task, 'Publishing readiness', max > 0 ? 'warn' : 'fail', max > 0 ?
              'Assignment worth ' + max + ' points. Publishing works only if this script\'s Cloud project created the assignment; otherwise Classroom refuses and the grade must be entered manually.' :
              'Assignment is ungraded in Classroom; there is no grade field to publish to.');
          } catch (e2) { add('Task: ' + task, 'Publishing readiness', 'fail', 'Assignment could not be read: ' + text(e2 && e2.message).slice(0, 120)); }
        });
      } catch (err) {
        add('Workbook', 'Opened', 'fail', text(err && err.message));
      }
    }
    var counts = { pass: 0, warn: 0, fail: 0, info: 0 };
    out.forEach(function (c) { counts[c.status]++; });
    return { success: true, data: { checks: out, counts: counts, ranAt: new Date().toISOString() } };
  }

  return { run: run };
})();

function apiWebHealthCheck(workbookId) { return HealthCheck.run(workbookId || ''); }
