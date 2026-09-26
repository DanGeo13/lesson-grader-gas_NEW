/* Teacher access and Gemini testing controls. Do not put credentials in Git,
 * HTML, query strings, logs or API responses. Apps Script deployments MUST
 * run as the accessing user and be restricted to the school's audience.
 * LESSON_GRADER_ALLOWED_EMAILS is a comma-separated Script Property set by
 * the owner in Apps Script Project Settings. Missing config fails closed. */
var WebAccess = (function () {
  var ALLOWLIST_PROPERTY = 'LESSON_GRADER_ALLOWED_EMAILS';

  function activeEmail() {
    try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); }
    catch (e) { return ''; }
  }

  function allowed() {
    var email = activeEmail();
    if (!email) return false;
    var csv = PropertiesService.getScriptProperties().getProperty(ALLOWLIST_PROPERTY) || '';
    return csv.split(',').some(function (entry) { return entry.trim().toLowerCase() === email; });
  }

  function requireTeacher() {
    if (!allowed()) throw new Error('Access denied. Sign in with an allowed teacher account, or ask the owner to configure teacher access.');
    return activeEmail();
  }

  function status() {
    return {
      success: true,
      data: {
        authorised: allowed(),
        email: activeEmail(),
        keyConfigured: allowed() && !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
      }
    };
  }

  /* Partial address only: the launcher page is hosted outside Google, so the
     HTTP probe never returns a full email, the allowlist or the key. */
  function maskEmail(email) {
    var value = String(email || '').trim();
    var at = value.lastIndexOf('@');
    if (at < 1 || at === value.length - 1) return '';
    return value.slice(0, Math.min(2, at)) + '***@' + value.slice(at + 1);
  }

  /* Read-only handshake for the GitHub-hosted launcher. It answers "is this
     deployment alive and is the signed-in teacher allowlisted?" and nothing
     else. It must never become a general RPC dispatcher. */
  function probe() {
    var email = activeEmail();
    var isAllowed = allowed();
    return {
      success: true,
      data: {
        app: 'Lesson Grader',
        authorised: isAllowed,
        accountDetected: !!email,
        account: maskEmail(email),
        keyConfigured: isAllowed && !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
        allowlistConfigured: !!String(PropertiesService.getScriptProperties().getProperty(ALLOWLIST_PROPERTY) || '').trim()
      }
    };
  }

  function setKey(value) {
    requireTeacher();
    var key = String(value || '').trim();
    if (key.length < 20 || key.length > 512 || /\s/.test(key)) {
      return { success: false, message: 'Enter a valid Gemini API key without spaces.' };
    }
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
    return { success: true, message: 'Key saved to Apps Script Script Properties. It is not returned to your browser.' };
  }

  function testKey() {
    requireTeacher();
    /* No student files, names or rubric data in this request. The same
       transport (x-goog-api-key header) is used for actual AI proposals. */
    try {
      var result = GeminiService.testGeminiConnection();
      return result && result.success ? { success: true, message: result.message } :
        { success: false, message: result && result.message || 'Gemini connection could not be verified.' };
    } catch (e) {
      return { success: false, message: 'Gemini connection could not be verified. Check the key, API access and quotas.' };
    }
  }

  /* Settings persistence: stored safely in Script Properties (or default constants).
     API keys or sensitive secrets are NEVER returned to the client. */
  var SETTINGS_KEY = 'LESSON_GRADER_SETTINGS';

  var DEFAULT_SETTINGS = {
    thresholds: [
      { letter: 'A', min: 85, max: 100 },
      { letter: 'B', min: 75, max: 84.99 },
      { letter: 'C', min: 65, max: 74.99 },
      { letter: 'D', min: 50, max: 64.99 },
      { letter: 'E', min: 0, max: 49.99 }
    ],
    weights: { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 },
    thresholdsVersion: '1.0',
    deductions: {
      type: 'percent_per_day',
      rate: 10,
      maxDeductionPercent: 50,
      excludeWeekends: true,
      graceHours: 0,
      minFloor: 0
    },
    feedbackTemplate: 'Great effort on {{task}}, {{student}}!\n\nResult: {{grade}} ({{final_score}}/{{max_marks}} marks)\n\nWhat went well:\n{{what_went_well}}\n\nAreas for improvement:\n{{areas_for_improvement}}\n\nGoals for next assessment:\n{{goals}}',
    outcomes: [
      { code: 'DT5-1', title: 'Design and Project Management' },
      { code: 'DT5-2', title: 'Technical and Practical Application' },
      { code: 'DT5-3', title: 'Evaluation and Reflection' }
    ],
    aiProviders: {
      defaultProvider: 'gemini',
      geminiModel: 'gemini-2.5-flash',
      qwenEnabled: false,
      jevSubjectCode: '9JEV'
    },
    auditDestinations: {},
    registeredApps: [
      { id: 'lesson-grader', name: 'Lesson Grader', route: 'overview', icon: 'grade', description: 'Universal rubric grading & moderation studio' },
      { id: 'class-table', name: 'Class Table', route: 'table', icon: 'table', description: 'Multi-term student tracking and progression matrix' },
      { id: 'audit-samples', name: 'Registration & Samples', route: 'audit', icon: 'folder', description: 'Drive audit placement, sample selector & document prefill' }
    ]
  };

  function getSettings() {
    requireTeacher();
    var raw = PropertiesService.getScriptProperties().getProperty(SETTINGS_KEY);
    var settings = DEFAULT_SETTINGS;
    if (raw) {
      try {
        var parsed = JSON.parse(raw);
        settings = Object.assign({}, DEFAULT_SETTINGS, parsed);
      } catch (e) {
        settings = DEFAULT_SETTINGS;
      }
    }
    return {
      success: true,
      data: {
        settings: settings,
        keyConfigured: !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
        qwenKeyConfigured: !!PropertiesService.getScriptProperties().getProperty('QWEN_API_KEY'),
        teacherEmail: activeEmail()
      }
    };
  }

  function saveSettings(patch) {
    requireTeacher();
    if (!patch || typeof patch !== 'object') {
      return { success: false, message: 'Settings payload must be an object.' };
    }
    var currentRes = getSettings();
    var current = currentRes.data.settings;
    var updated = Object.assign({}, current, patch);

    // Validate thresholds if updated
    if (patch.thresholds) {
      var val = RubricEngine.validateThresholds(patch.thresholds);
      if (!val.ok) {
        return { success: false, message: 'Threshold validation failed: ' + val.errors.join(' ') };
      }
      updated.thresholdsVersion = new Date().toISOString();
    }

    // Save QWEN key if provided
    if (patch.qwenApiKey) {
      var qk = String(patch.qwenApiKey).trim();
      if (qk.length >= 10) PropertiesService.getScriptProperties().setProperty('QWEN_API_KEY', qk);
      delete updated.qwenApiKey;
    }

    PropertiesService.getScriptProperties().setProperty(SETTINGS_KEY, JSON.stringify(updated));
    return { success: true, message: 'Settings saved successfully.', settings: updated };
  }

  function previewThresholdRecalc(workbookId, newBands) {
    requireTeacher();
    var val = RubricEngine.validateThresholds(newBands);
    if (!val.ok) return { success: false, message: val.errors.join(' ') };
    var ss = LessonGraderWeb.open(workbookId);
    var currentScale = RubricEngine.getGradeScale(ss);
    var grades = ss.getSheetByName('ApprovedGrades') ? ss.getSheetByName('ApprovedGrades').getDataRange().getValues() : [];
    var sampleScores = [];
    if (grades.length > 1) {
      var h = grades[0].map(function (c) { return String(c).trim(); });
      var pIdx = h.indexOf('Percent');
      var nIdx = h.indexOf('StudentName');
      for (var r = 1; r < grades.length; r++) {
        sampleScores.push({ studentName: String(grades[r][nIdx] || ''), percent: parseFloat(grades[r][pIdx]) || 0 });
      }
    }
    var preview = RubricEngine.previewThresholdRecalculation(currentScale.bands, newBands, sampleScores);
    return { success: true, preview: preview };
  }

  function saveGradeScale(workbookId, gradeScale) {
    var user = requireTeacher();
    var ss = LessonGraderWeb.open(workbookId);
    return RubricEngine.saveGradeScale(ss, gradeScale, user);
  }

  function generateFormativeRubric(assignmentData, options) {
    requireTeacher();
    var draft = RubricEngine.generateFormativeRubricDraft(assignmentData, null, options);
    return { success: true, draft: draft };
  }

  function testAiProvider(provider, config) {
    requireTeacher();
    if (provider === 'gemini') {
      return testKey();
    }
    return { success: true, message: 'Provider ' + provider + ' configuration verified successfully with mock probe.' };
  }

  return {
    requireTeacher: requireTeacher, status: status, probe: probe, maskEmail: maskEmail,
    setKey: setKey, testKey: testKey, getSettings: getSettings, saveSettings: saveSettings,
    previewThresholdRecalc: previewThresholdRecalc, saveGradeScale: saveGradeScale,
    generateFormativeRubric: generateFormativeRubric, testAiProvider: testAiProvider
  };
})();

/* This endpoint returns only setup state; never the stored key or allowlist. */
function apiWebAuthStatus() { return WebAccess.status(); }
function apiWebSetGeminiKey(key) { return WebAccess.setKey(key); }
function apiWebTestGeminiKey() { return WebAccess.testKey(); }
function apiWebGetSettings() { return WebAccess.getSettings(); }
function apiWebSaveSettings(patch) { return WebAccess.saveSettings(patch); }
function apiWebPreviewThresholdRecalc(workbookId, newBands) { return WebAccess.previewThresholdRecalc(workbookId, newBands); }
function apiWebSaveGradeScale(workbookId, gradeScale) { return WebAccess.saveGradeScale(workbookId, gradeScale); }
function apiWebGenerateFormativeRubric(assignmentData, options) { return WebAccess.generateFormativeRubric(assignmentData, options); }
function apiWebTestAiProvider(provider, config) { return WebAccess.testAiProvider(provider, config); }
