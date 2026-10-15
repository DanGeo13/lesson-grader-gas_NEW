/* Teacher access, credential handling and validated settings.
 *
 * Do not put credentials in Git, HTML, query strings, logs or API responses.
 * Deployments MUST run as the accessing user and be restricted to the school
 * domain. LESSON_GRADER_ALLOWED_EMAILS is a comma-separated Script Property set
 * by the owner in Apps Script Project Settings. Missing config fails closed. */
var WebAccess = (function () {
  /* Resolved lazily: Apps Script file load order is not guaranteed. */
  function P() { return AppConfig.PROPS; }

  function activeEmail() {
    try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); }
    catch (e) { return ''; }
  }

  function effectiveEmail() {
    try { return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); }
    catch (e) { return ''; }
  }

  function isAllowed(email) {
    return !!email && AppConfig.emailList(P().ALLOWLIST).indexOf(email) !== -1;
  }

  function allowed() { return isAllowed(activeEmail()); }

  /** Throws unless the signed-in user is an allowlisted teacher; returns their email. */
  function requireTeacher() {
    if (!allowed()) throw new Error('Access denied. Sign in with an allowed teacher account, or ask the owner to configure teacher access.');
    return activeEmail();
  }

  /** Time-driven triggers run as their owner; getActiveUser can be blank there. */
  function requireTriggerOwner() {
    var email = activeEmail() || effectiveEmail();
    if (!isAllowed(email)) throw new Error('Trigger owner is not an allowed teacher account.');
    return email;
  }

  /** Unlocking a final assessment is restricted to LESSON_GRADER_UNLOCK_EMAILS (fails closed). */
  function canUnlock(email) {
    var e = String(email || '').toLowerCase();
    return isAllowed(e) && AppConfig.emailList(P().UNLOCK_ALLOWLIST).indexOf(e) !== -1;
  }

  function hasProp(name) { return !!String(AppConfig.prop(name) || '').trim(); }

  function hostOf(url) {
    var m = /^https:\/\/([^\/?#:]+)/i.exec(String(url || ''));
    return m ? m[1] : '';
  }

  function providerStatus() {
    return {
      gemini: { configured: hasProp(P().GEMINI_KEY) },
      qwen: { configured: hasProp(P().QWEN_ENDPOINT), host: hostOf(AppConfig.prop(P().QWEN_ENDPOINT)), keyConfigured: hasProp(P().QWEN_KEY) },
      jev: { configured: hasProp(P().JEV_ENDPOINT), host: hostOf(AppConfig.prop(P().JEV_ENDPOINT)), keyConfigured: hasProp(P().JEV_KEY) }
    };
  }

  function status() {
    var ok = allowed();
    return {
      success: true,
      data: {
        authorised: ok,
        email: ok ? activeEmail() : '',
        keyConfigured: ok && hasProp(P().GEMINI_KEY),
        canUnlock: ok && canUnlock(activeEmail())
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
    var isOk = isAllowed(email);
    return {
      success: true,
      data: {
        app: 'Lesson Grader',
        authorised: isOk,
        accountDetected: !!email,
        account: maskEmail(email),
        keyConfigured: isOk && hasProp(P().GEMINI_KEY),
        allowlistConfigured: hasProp(P().ALLOWLIST)
      }
    };
  }

  function setKey(value) {
    requireTeacher();
    var key = String(value || '').trim();
    if (key.length < 20 || key.length > 512 || /\s/.test(key)) {
      return { success: false, message: 'Enter a valid Gemini API key without spaces.' };
    }
    PropertiesService.getScriptProperties().setProperty(P().GEMINI_KEY, key);
    return { success: true, message: 'Key saved to Script Properties. It is not returned to your browser.' };
  }

  /** Saves a Qwen or JEV endpoint (https only) and optional key. Never echoes the key. */
  function setProviderConfig(provider, config) {
    requireTeacher();
    var names = provider === 'qwen' ? { url: P().QWEN_ENDPOINT, key: P().QWEN_KEY } :
      provider === 'jev' ? { url: P().JEV_ENDPOINT, key: P().JEV_KEY } : null;
    if (!names) return { success: false, message: 'Unknown AI provider.' };
    var c = config || {};
    var props = PropertiesService.getScriptProperties();
    if (c.clear === true) {
      props.deleteProperty(names.url); props.deleteProperty(names.key);
      return { success: true, message: provider.toUpperCase() + ' configuration removed.', data: providerStatus() };
    }
    var url = String(c.endpoint || '').trim();
    if (url) {
      if (!/^https:\/\/[^\s]{4,300}$/i.test(url)) return { success: false, message: 'Endpoint must be an https:// URL.' };
      props.setProperty(names.url, url);
    }
    var key = String(c.key || '').trim();
    if (key) {
      if (key.length < 8 || key.length > 512 || /\s/.test(key)) return { success: false, message: 'Enter a key without spaces.' };
      props.setProperty(names.key, key);
    }
    return { success: true, message: provider.toUpperCase() + ' configuration saved. Keys are not returned to the browser.', data: providerStatus() };
  }

  function testKey() {
    requireTeacher();
    /* No student files, names or rubric data in this request. */
    try {
      var result = GeminiService.testGeminiConnection();
      return result && result.success ? { success: true, message: result.message } :
        { success: false, message: result && result.message || 'Gemini connection could not be verified.' };
    } catch (e) {
      return { success: false, message: 'Gemini connection could not be verified. Check the key, API access and quotas.' };
    }
  }

  function getSettings() {
    var email = requireTeacher();
    return {
      success: true,
      data: {
        settings: AppConfig.readSettings(),
        providers: providerStatus(),
        keyConfigured: hasProp(P().GEMINI_KEY),
        teacherEmail: email,
        canUnlock: canUnlock(email)
      }
    };
  }

  function saveSettings(patch) {
    requireTeacher();
    var merged = AppConfig.mergeSettings(AppConfig.readSettings(), patch);
    if (!merged.ok) return { success: false, message: merged.errors.join(' ') };
    AppConfig.writeSettings(merged.settings);
    return { success: true, message: 'Settings saved.', data: { settings: merged.settings } };
  }

  function previewThresholdRecalc(workbookId, newBands) {
    requireTeacher();
    var val = RubricEngine.validateThresholds(newBands);
    if (!val.ok) return { success: false, message: val.errors.join(' ') };
    var ss = LessonGraderWeb.open(workbookId);
    var currentScale = RubricEngine.getGradeScale(ss);
    var sheet = ss.getSheetByName('ApprovedGrades');
    var grades = sheet ? sheet.getDataRange().getValues() : [];
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
    return { success: true, preview: preview,
      message: 'Approved results are snapshots and will NOT change. Only drafts recalculated after saving would use the new scale.' };
  }

  function saveGradeScale(workbookId, gradeScale) {
    var user = requireTeacher();
    var ss = LessonGraderWeb.open(workbookId);
    return RubricEngine.saveGradeScale(ss, gradeScale, user);
  }

  function generateFormativeRubric(assignmentData, options) {
    requireTeacher();
    var draft = RubricEngine.generateFormativeRubricDraft(assignmentData, null, options);
    return { success: true, draft: draft, message: 'Generic template only. It is not derived from your syllabus and must be rewritten before use.' };
  }

  /* Real connectivity checks only. No student data is sent. */
  function testAiProvider(provider) {
    requireTeacher();
    if (provider === 'gemini') return testKey();
    if (typeof AiGateway === 'undefined') return { success: false, message: 'Provider ' + provider + ' is not available in this deployment.' };
    return AiGateway.healthCheck(provider);
  }

  return {
    requireTeacher: requireTeacher, requireTriggerOwner: requireTriggerOwner, canUnlock: canUnlock,
    activeEmail: activeEmail, isAllowed: isAllowed, status: status, probe: probe, maskEmail: maskEmail,
    setKey: setKey, setProviderConfig: setProviderConfig, providerStatus: providerStatus, testKey: testKey,
    getSettings: getSettings, saveSettings: saveSettings, previewThresholdRecalc: previewThresholdRecalc,
    saveGradeScale: saveGradeScale, generateFormativeRubric: generateFormativeRubric, testAiProvider: testAiProvider
  };
})();

/* These endpoints return setup state only; never a stored key or the allowlist. */
function apiWebAuthStatus() { return WebAccess.status(); }
function apiWebSetGeminiKey(key) { return WebAccess.setKey(key); }
function apiWebTestGeminiKey() { return WebAccess.testKey(); }
function apiWebSetAiProviderConfig(provider, config) { return WebAccess.setProviderConfig(provider, config); }
function apiWebGetSettings() { return WebAccess.getSettings(); }
function apiWebSaveSettings(patch) { return WebAccess.saveSettings(patch); }
function apiWebPreviewThresholdRecalc(workbookId, newBands) { return WebAccess.previewThresholdRecalc(workbookId, newBands); }
function apiWebSaveGradeScale(workbookId, gradeScale) { return WebAccess.saveGradeScale(workbookId, gradeScale); }
function apiWebGenerateFormativeRubric(assignmentData, options) { return WebAccess.generateFormativeRubric(assignmentData, options); }
function apiWebTestAiProvider(provider) { return WebAccess.testAiProvider(provider); }
