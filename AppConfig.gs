/* ============================================================================
 * AppConfig.gs — single home for deployment configuration, property names,
 * assessment state vocabulary and validated teacher settings.
 *
 * Load-time pure: no Apps Script service is touched until a function runs.
 * Secrets (API keys) live ONLY in Script Properties under the names below and
 * are never part of the settings JSON or any value returned to a browser.
 * ============================================================================ */
var AppConfig = (function () {
  var TIME_ZONE = 'Australia/Sydney';

  /* Script Property names. Values marked (secret) must never be returned. */
  var PROPS = {
    ALLOWLIST: 'LESSON_GRADER_ALLOWED_EMAILS',      // comma-separated teacher emails
    UNLOCK_ALLOWLIST: 'LESSON_GRADER_UNLOCK_EMAILS', // who may unlock final assessments
    SETTINGS: 'LESSON_GRADER_SETTINGS',             // validated JSON (no secrets)
    WEB_CATEGORY: 'LESSON_GRADER_WEB_CATEGORY',     // default Drive category folder
    ALLOW_EMBED: 'LESSON_GRADER_ALLOW_EMBED',       // 'true' lets the launcher iframe the app
    GEMINI_KEY: 'GEMINI_API_KEY',                   // (secret)
    QWEN_ENDPOINT: 'QWEN_ENDPOINT_URL',
    QWEN_KEY: 'QWEN_API_KEY',                       // (secret)
    JEV_ENDPOINT: 'JEV_ENDPOINT_URL',
    JEV_KEY: 'JEV_API_KEY',                         // (secret)
    PROMPTS: 'LESSON_GRADER_PROMPTS',               // prompt versions + activation record
    AI_HEALTH: 'LESSON_GRADER_AI_HEALTH'            // last known provider health (no secrets)
  };
  var SECRET_PROPS = [PROPS.GEMINI_KEY, PROPS.QWEN_KEY, PROPS.JEV_KEY];

  /* Persisted Submissions.Status values (unchanged vocabulary; see REVIEW.md §1.7). */
  var STATUS = { NEW: 'New', IN_REVIEW: 'InReview', APPROVED: 'Approved', LOCKED: 'Locked', SUPERSEDED: 'Superseded' };
  var FINAL_STATUSES = [STATUS.APPROVED, STATUS.LOCKED, STATUS.SUPERSEDED];

  /* Teacher-facing label for a submission version. Reassessment is a version
     property (SourceType/ParentSubmissionRecordID), not a separate status. */
  function stateLabel(status, isReassessment, version) {
    var s = String(status || STATUS.NEW);
    var base = s === STATUS.LOCKED ? 'Final (locked)' : s === STATUS.APPROVED ? 'Teacher-approved' :
      s === STATUS.SUPERSEDED ? 'Superseded' : s === STATUS.IN_REVIEW ? 'Draft' : 'Not started';
    return isReassessment ? base + ' · Reassessment v' + (Number(version) || 2) : base;
  }

  var GRADE_DISPLAYS = ['letter', 'mark', 'percent', 'letter+mark', 'letter+percent'];
  var CLASSROOM_GRADE_MODES = ['scaled-points', 'raw-points', 'percent', 'none'];
  var PROVIDERS = { draft: ['gemini', 'jev'], vision: ['none', 'gemini', 'qwen'] };

  var DEFAULT_SETTINGS = {
    outputFormat: {
      gradeDisplay: 'letter+mark',
      useFirstName: true,
      classroomGradeMode: 'scaled-points',
      includeGradeInComment: true
    },
    feedbackTemplate: '{{opening}}\n\nResult: {{grade_display}}\n\nWhat went well:\n{{what_went_well}}\n\nAreas for improvement:\n{{areas_for_improvement}}\n\nGoals for next assessment:\n{{goals}}',
    aiProviders: {
      draftProvider: 'gemini',
      visionProvider: 'none',
      geminiModel: 'gemini-2.5-flash',
      qwenModel: 'qwen3-vl',
      jevModel: '',
      jevSubjectCode: '9JEV'
    },
    /* Stored for compatibility only. NOT applied to grades: no late-penalty
       rule has been confirmed (REVIEW.md §4). */
    deductions: { type: 'none', rate: 0, maxDeductionPercent: 50, excludeWeekends: true, graceHours: 0, minFloor: 0 },
    outcomes: [
      { code: 'DT5-1', title: 'Design and Project Management' },
      { code: 'DT5-2', title: 'Technical and Practical Application' },
      { code: 'DT5-3', title: 'Evaluation and Reflection' }
    ]
  };

  var SECRET_KEY_RE = /(api[_-]?key|secret|token|password|credential)/i;

  function clone(v) { return JSON.parse(JSON.stringify(v)); }

  /* Remove secret-looking keys at any depth (protects against legacy settings
     JSON that once stored geminiApiKey — REVIEW.md F01). */
  function scrubSecrets(value) {
    if (Array.isArray(value)) return value.map(scrubSecrets);
    if (!value || typeof value !== 'object') return value;
    var out = {};
    Object.keys(value).forEach(function (k) {
      if (SECRET_KEY_RE.test(k)) return;
      out[k] = scrubSecrets(value[k]);
    });
    return out;
  }

  function str(v, max) { return String(v == null ? '' : v).slice(0, max || 200); }
  function oneOf(v, list, fallback) { return list.indexOf(v) !== -1 ? v : fallback; }

  /* Merge a teacher patch into current settings. Unknown keys and anything
     secret-looking are rejected, never silently stored. Returns {ok, settings, errors}. */
  function mergeSettings(current, patch) {
    var errors = [];
    var base = clone(current || DEFAULT_SETTINGS);
    if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return { ok: false, errors: ['Settings must be an object.'] };
    Object.keys(patch).forEach(function (key) {
      if (SECRET_KEY_RE.test(key)) { errors.push('API keys are saved with their own control and never stored in settings.'); return; }
      var v = patch[key];
      switch (key) {
        case 'outputFormat':
          if (!v || typeof v !== 'object') { errors.push('outputFormat must be an object.'); return; }
          base.outputFormat = {
            gradeDisplay: oneOf(v.gradeDisplay, GRADE_DISPLAYS, base.outputFormat.gradeDisplay),
            useFirstName: v.useFirstName === undefined ? base.outputFormat.useFirstName : !!v.useFirstName,
            classroomGradeMode: oneOf(v.classroomGradeMode, CLASSROOM_GRADE_MODES, base.outputFormat.classroomGradeMode),
            includeGradeInComment: v.includeGradeInComment === undefined ? base.outputFormat.includeGradeInComment : !!v.includeGradeInComment
          };
          return;
        case 'feedbackTemplate':
          if (typeof v !== 'string' || v.length > 4000) { errors.push('Feedback template must be text under 4,000 characters.'); return; }
          base.feedbackTemplate = v;
          return;
        case 'aiProviders':
          if (!v || typeof v !== 'object') { errors.push('aiProviders must be an object.'); return; }
          if (Object.keys(v).some(function (k) { return SECRET_KEY_RE.test(k); })) {
            errors.push('API keys are saved with their own control and never stored in settings.'); return;
          }
          base.aiProviders = {
            draftProvider: oneOf(v.draftProvider, PROVIDERS.draft, base.aiProviders.draftProvider),
            visionProvider: oneOf(v.visionProvider, PROVIDERS.vision, base.aiProviders.visionProvider),
            geminiModel: /^[a-z0-9.\-]{3,60}$/.test(String(v.geminiModel || '')) ? String(v.geminiModel) : base.aiProviders.geminiModel,
            qwenModel: str(v.qwenModel != null ? v.qwenModel : base.aiProviders.qwenModel, 80),
            jevModel: str(v.jevModel != null ? v.jevModel : base.aiProviders.jevModel, 80),
            jevSubjectCode: str(v.jevSubjectCode != null ? v.jevSubjectCode : base.aiProviders.jevSubjectCode, 20)
          };
          return;
        case 'deductions':
          if (!v || typeof v !== 'object') { errors.push('deductions must be an object.'); return; }
          base.deductions = {
            type: oneOf(v.type, ['none', 'percent_per_day', 'fixed_marks', 'percent_of_achieved'], 'none'),
            rate: Math.max(0, Math.min(100, Number(v.rate) || 0)),
            maxDeductionPercent: Math.max(0, Math.min(100, Number(v.maxDeductionPercent) || 50)),
            excludeWeekends: !!v.excludeWeekends,
            graceHours: Math.max(0, Number(v.graceHours) || 0),
            minFloor: Math.max(0, Number(v.minFloor) || 0)
          };
          return;
        case 'outcomes':
          if (!Array.isArray(v) || v.length > 60) { errors.push('outcomes must be a list of up to 60 items.'); return; }
          base.outcomes = v.map(function (o) { return { code: str(o && o.code, 20), title: str(o && o.title, 160) }; })
            .filter(function (o) { return o.code; });
          return;
        case 'thresholds':
        case 'weights':
        case 'thresholdsVersion':
          errors.push('Grade boundaries and band weights are saved to each workbook\'s Setup sheet, not to app settings.');
          return;
        default:
          errors.push('Unknown setting: ' + key);
      }
    });
    return { ok: errors.length === 0, settings: base, errors: errors };
  }

  /* Stored JSON may predate validation: keep only known keys, drop secrets. */
  function normaliseStored(raw) {
    var parsed = {};
    try { parsed = raw ? JSON.parse(raw) : {}; } catch (e) { parsed = {}; }
    parsed = scrubSecrets(parsed || {});
    var allowed = {};
    ['outputFormat', 'feedbackTemplate', 'aiProviders', 'deductions', 'outcomes'].forEach(function (k) {
      if (parsed[k] !== undefined) allowed[k] = parsed[k];
    });
    var merged = mergeSettings(DEFAULT_SETTINGS, allowed);
    return merged.settings;
  }

  function props() { return PropertiesService.getScriptProperties(); }
  function prop(name) { return props().getProperty(name); }
  function flag(name) { return String(prop(name) || '').trim().toLowerCase() === 'true'; }

  function readSettings() { return normaliseStored(prop(PROPS.SETTINGS)); }
  function writeSettings(settings) { props().setProperty(PROPS.SETTINGS, JSON.stringify(scrubSecrets(settings))); }

  function emailList(name) {
    return String(prop(name) || '').split(',').map(function (e) { return e.trim().toLowerCase(); }).filter(Boolean);
  }

  return {
    TIME_ZONE: TIME_ZONE, PROPS: PROPS, SECRET_PROPS: SECRET_PROPS, STATUS: STATUS, FINAL_STATUSES: FINAL_STATUSES,
    GRADE_DISPLAYS: GRADE_DISPLAYS, CLASSROOM_GRADE_MODES: CLASSROOM_GRADE_MODES, PROVIDERS: PROVIDERS,
    DEFAULT_SETTINGS: DEFAULT_SETTINGS, stateLabel: stateLabel, scrubSecrets: scrubSecrets,
    mergeSettings: mergeSettings, normaliseStored: normaliseStored, readSettings: readSettings,
    writeSettings: writeSettings, prop: prop, flag: flag, emailList: emailList
  };
})();
