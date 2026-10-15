/* ============================================================================
 * AiGateway.gs — AI evidence drafting behind one boundary.
 *
 * Providers
 *   gemini  Google Gemini (existing). Receives the files themselves.
 *   jev     owner-hosted text evaluator, OpenAI-compatible chat endpoint
 *           (JEV_ENDPOINT_URL). Receives ONLY de-identified Markdown + rubric.
 *   qwen    owner-hosted Qwen 3-VL, OpenAI-compatible (QWEN_ENDPOINT_URL).
 *           Optional "vision" step that describes images as text for JEV.
 * Rules
 *   - AI output is untrusted data: parsed defensively, validated against the
 *     pinned rubric, never executed or rendered as HTML.
 *   - A tick is accepted only with an evidence note citing a supplied file.
 *     Uncited ticks are discarded and reported; a criterion with no accepted
 *     tick is flagged "insufficient evidence" for the teacher.
 *   - Marks, percentages or grades in AI output are ignored. Scores are always
 *     recalculated by RubricEngine from the accepted ticks (SC17).
 *   - Provider failures return {success:false, unavailable:true} so the
 *     teacher can queue the work or continue manually (SC48).
 * ============================================================================ */
var AiGateway = (function () {
  var VISION_PROMPT = 'Describe, as plain Markdown, the visible design and practical evidence in this image for a teacher marking a school project. ' +
    'Do not name or describe any person. Do not guess. Transcribe headings and labels you can read.';

  function text(v) { return v == null ? '' : String(v); }

  /* ---- pure helpers ------------------------------------------------------- */

  /** Extracts a JSON object from model text (tolerates ```json fences). Returns null if invalid. */
  function parseModelJson(raw) {
    var t = text(raw).trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
    var start = t.indexOf('{'), end = t.lastIndexOf('}');
    if (start === -1 || end <= start) return null;
    try { var obj = JSON.parse(t.slice(start, end + 1)); return obj && typeof obj === 'object' ? obj : null; }
    catch (e) { return null; }
  }

  /**
   * Validates a model proposal against the pinned rubric and supplied file labels.
   * @return {Object} {criterionId: {checks, evidence, unsupported, insufficient, grade}}
   */
  function validateProposal(profile, parsed, fileLabels) {
    if (!parsed || !Array.isArray(parsed.criteria)) throw new Error('The AI reply was not in the expected format. No draft was changed.');
    var ids = RubricEngine.knownObservableIds(profile);
    var labels = (fileLabels || []).map(function (l) { return text(l).trim().toLowerCase(); });
    var scale = RubricEngine.normalizeGradeScale(null, profile.gradeScaleBands);
    var out = {};
    profile.criteria.forEach(function (c) {
      var p = parsed.criteria.filter(function (x) { return x && x.criterionId === c.criterionId; })[0] || {};
      var notes = Array.isArray(p.evidenceNotes) ? p.evidenceNotes : [];
      var accepted = {}, unsupported = [], evidence = [];
      (Array.isArray(p.tickedCheckboxes) ? p.tickedCheckboxes : []).slice(0, 200).forEach(function (id) {
        if (typeof id !== 'string' || ids[id] !== c.criterionId || id === c.mdOverride.id || accepted[id]) return;
        var note = notes.filter(function (n) {
          return n && n.checkboxId === id && text(n.note).trim() && labels.indexOf(text(n.fileName).trim().toLowerCase()) !== -1;
        })[0];
        if (note) {
          accepted[id] = true;
          evidence.push({ checkboxId: id, fileName: text(note.fileName).trim().slice(0, 180), note: text(note.note).trim().slice(0, 1000) });
        } else {
          unsupported.push(id);
        }
      });
      out[c.criterionId] = {
        checks: accepted, evidence: evidence, unsupported: unsupported,
        insufficient: !Object.keys(accepted).length,
        grade: RubricEngine.scoreCriterion(c, accepted, scale).grade
      };
    });
    return out;
  }

  /* ---- health ------------------------------------------------------------- */

  function readHealth() {
    try { return JSON.parse(AppConfig.prop(AppConfig.PROPS.AI_HEALTH) || '{}') || {}; } catch (e) { return {}; }
  }

  function recordHealth(provider, ok, message) {
    try {
      var h = readHealth();
      h[provider] = { ok: !!ok, at: new Date().toISOString(), message: text(message).slice(0, 200) };
      PropertiesService.getScriptProperties().setProperty(AppConfig.PROPS.AI_HEALTH, JSON.stringify(h));
    } catch (e) { /* health is advisory only */ }
  }

  /* ---- transport ---------------------------------------------------------- */

  function unavailableCode(code) { return code === 0 || code === 408 || code === 429 || code >= 500; }

  function callOpenAiCompatible(provider, messages, opts) {
    var P = AppConfig.PROPS;
    var endpoint = text(AppConfig.prop(provider === 'qwen' ? P.QWEN_ENDPOINT : P.JEV_ENDPOINT)).trim();
    var key = text(AppConfig.prop(provider === 'qwen' ? P.QWEN_KEY : P.JEV_KEY)).trim();
    var settings = AppConfig.readSettings().aiProviders;
    if (!endpoint) return { success: false, unavailable: true, notConfigured: true, message: provider.toUpperCase() + ' endpoint is not configured (Settings → AI providers).' };
    var payload = { model: (opts && opts.model) || (provider === 'qwen' ? settings.qwenModel : settings.jevModel) || undefined,
      messages: messages, temperature: 0.1, max_tokens: 4000 };
    if (opts && opts.json) payload.response_format = { type: 'json_object' };
    var headers = {};
    if (key) headers.Authorization = 'Bearer ' + key;
    try {
      var res = UrlFetchApp.fetch(endpoint, { method: 'post', contentType: 'application/json', headers: headers,
        payload: JSON.stringify(payload), muteHttpExceptions: true });
      var code = res.getResponseCode();
      if (code !== 200) {
        recordHealth(provider, false, 'HTTP ' + code);
        return { success: false, unavailable: unavailableCode(code), message: provider.toUpperCase() + ' returned HTTP ' + code + '.' };
      }
      var body = JSON.parse(res.getContentText());
      var content = body && body.choices && body.choices[0] && body.choices[0].message && body.choices[0].message.content;
      if (Array.isArray(content)) content = content.map(function (p) { return text(p.text); }).join('');
      recordHealth(provider, true, 'OK');
      return { success: true, text: text(content), modelUsed: text(body.model || payload.model || provider) };
    } catch (e) {
      recordHealth(provider, false, 'unreachable');
      return { success: false, unavailable: true, message: provider.toUpperCase() + ' could not be reached. Check that the host is online.' };
    }
  }

  function callGemini(parts, requireJson) {
    try {
      var model = AppConfig.readSettings().aiProviders.geminiModel;
      var res = GeminiService.callGeminiWithFallback([{ role: 'user', parts: parts }], requireJson, model);
      recordHealth('gemini', res.success, res.success ? 'OK' : res.message);
      if (!res.success) return { success: false, unavailable: res.unavailable !== false, message: text(res.message) };
      return res;
    } catch (e) {
      recordHealth('gemini', false, 'not configured');
      return { success: false, unavailable: true, notConfigured: true, message: 'Gemini is not configured (Settings → AI providers).' };
    }
  }

  /** Connectivity check with no student data. */
  function healthCheck(provider) {
    if (provider === 'gemini') {
      var g = GeminiService.testGeminiConnection();
      recordHealth('gemini', g.success, g.message);
      return { success: !!g.success, message: g.message };
    }
    if (provider !== 'qwen' && provider !== 'jev') return { success: false, message: 'Unknown provider.' };
    var r = callOpenAiCompatible(provider, [{ role: 'user', content: 'Reply with the single word OK.' }], {});
    return r.success ? { success: true, message: provider.toUpperCase() + ' responded (' + r.modelUsed + ').' } : { success: false, message: r.message };
  }

  /* ---- drafting ----------------------------------------------------------- */

  function identityFor(ss, sub) {
    var roster = WebGrading.records(ss, 'Submissions').map(function (r) { return text(r.StudentName); }).filter(Boolean);
    return { names: [text(sub.StudentName)], rosterNames: roster,
      ids: [text(sub.StudentUserID), text(sub.StudentEmail), text(sub.SubmissionRecordID), text(sub.ClassroomSubmissionID)].filter(Boolean) };
  }

  function describeImage(provider, f) {
    if (provider === 'qwen') {
      if (!/^image\//.test(f.inlineData.mimeType)) return { success: false, unavailable: false, message: 'Qwen vision reads images only; ' + f.fileName + ' is a PDF.' };
      return callOpenAiCompatible('qwen', [{ role: 'user', content: [
        { type: 'text', text: VISION_PROMPT },
        { type: 'image_url', image_url: { url: 'data:' + f.inlineData.mimeType + ';base64,' + f.inlineData.data } }] }], {});
    }
    return callGemini([{ text: VISION_PROMPT }, { inlineData: f.inlineData }], false);
  }

  /**
   * Builds and sends one drafting request. With dryRun nothing is written.
   * @return {{success:boolean, unavailable?:boolean, message?:string, profile, proposals, files, provider, model, promptId, payloadPreview}}
   */
  function proposeForSubmission(ss, taskName, submissionId, options) {
    var opt = options || {};
    var sub = WebGrading.selectSubmission(ss, taskName, submissionId);
    var a = WebGrading.assess(ss, taskName, sub);
    if (!a.profile) throw new Error('No pinned rubric is available for this submission.');
    var settings = AppConfig.readSettings().aiProviders;
    var provider = opt.provider || settings.draftProvider;
    var vision = settings.visionProvider;
    var prompt = opt.promptText ? { id: opt.promptId || 'candidate', text: opt.promptText } : PromptRegistry.active();
    var fileRows = WebGrading.records(ss, 'SubmissionFiles').filter(function (r) {
      return text(r.SubmissionRecordID) === text(submissionId) && text(r.AIReviewStatus) !== 'Removed';
    });
    var ev = EvidenceService.collect(fileRows, provider === 'gemini' ? { wantInline: true, wantText: true } : { wantInline: vision !== 'none', wantText: true });
    var rubricText = RubricEngine.serializeForAi(a.profile);
    var reviewed = ev.files.filter(function (f) { return f.status === 'Reviewed'; });
    var labels = [], labelToName = {};
    var res, payloadPreview = '';

    if (provider === 'gemini') {
      var parts = [{ text: PromptRegistry.render(prompt.text, taskName, rubricText) }];
      reviewed.forEach(function (f) {
        labels.push(f.fileName); labelToName[f.fileName.toLowerCase()] = f.fileName;
        if (f.markdown != null) parts.push({ text: '=== FILE: ' + f.fileName + ' ===\n' + f.markdown });
        else { parts.push({ text: '=== FILE: ' + f.fileName + ' (attached below) ===' }); parts.push({ inlineData: f.inlineData }); }
      });
      if (!labels.length) return { success: false, unavailable: false, files: ev.files, message: 'No student file could be read. Assess manually, or check file access.' };
      res = callGemini(parts, true);
    } else if (provider === 'jev') {
      var identity = identityFor(ss, sub);
      var chunks = [];
      for (var i = 0; i < reviewed.length; i++) {
        var f = reviewed[i];
        var label = 'File ' + (i + 1);
        var md = f.markdown;
        if (md == null && f.inlineData) {
          var d = describeImage(vision, f);
          if (!d.success) {
            if (d.unavailable) return { success: false, unavailable: true, provider: vision, files: ev.files, message: 'Vision provider unavailable: ' + d.message };
            f.status = 'Unsupported'; f.limitations = d.message; continue;
          }
          md = d.text;
          f.limitations = 'Image described by ' + vision + '; the description, not the image, was evaluated.';
        }
        if (md == null) continue;
        labels.push(label); labelToName[label.toLowerCase()] = f.fileName;
        chunks.push('=== FILE: ' + label + ' ===\n' + Deidentify.apply(md, identity).text);
      }
      if (!labels.length) return { success: false, unavailable: false, files: ev.files, message: 'No student file could be converted to text for JEV. Assess manually.' };
      var system = PromptRegistry.render(prompt.text, taskName, rubricText);
      var user = chunks.join('\n\n');
      Deidentify.assertClean(system + '\n' + user, identity);
      payloadPreview = user.slice(0, 2000);
      res = callOpenAiCompatible('jev', [{ role: 'system', content: system }, { role: 'user', content: user }], { json: true });
    } else {
      throw new Error('Unknown drafting provider: ' + provider);
    }

    if (!res.success) return { success: false, unavailable: !!res.unavailable, provider: provider, files: ev.files, message: res.message };
    var parsed = parseModelJson(res.text);
    var proposals = validateProposal(a.profile, parsed, labels);
    Object.keys(proposals).forEach(function (cid) {
      proposals[cid].provider = provider;
      proposals[cid].evidence.forEach(function (e) { e.fileName = labelToName[e.fileName.toLowerCase()] || e.fileName; });
    });
    return { success: true, profile: a.profile, assessment: a, sub: sub, proposals: proposals, files: ev.files, provider: provider,
      model: text(res.modelUsed), promptId: prompt.id, payloadPreview: payloadPreview, deidentified: provider === 'jev' };
  }

  function logRun(ss, submissionId, assessmentId, out, status, message, started, actor) {
    WebGrading.append(ss, 'AIAssessments', {
      AssessmentID: assessmentId || '', SubmissionRecordID: submissionId, RunType: 'EvidenceProposal', RunTimestamp: new Date().toISOString(),
      ModelUsed: text(out && out.model), PromptVersion: text(out && out.promptId),
      InputFilesJSON: JSON.stringify(((out && out.files) || []).map(function (f) { return { fileName: f.fileName, status: f.status, limitations: f.limitations }; })),
      PromptSummary: (out && out.provider ? out.provider : '') + (out && out.deidentified ? ' (de-identified Markdown only)' : ''),
      ResponseJSON: out && out.proposals ? JSON.stringify({ proposals: out.proposals, payloadPreview: out.payloadPreview || '' }).slice(0, 45000) : '',
      ExecutionStatus: status, ErrorMessage: text(message).slice(0, 450), LatencySeconds: (Date.now() - started) / 1000,
      TriggeredBy: actor, TeacherOutcome: status === 'Success' ? 'Pending teacher review' : 'Not applied'
    });
  }

  /**
   * Drafts one submission and saves the proposal as a draft (never approves).
   * If the teacher has already saved selections, they are kept and the AI
   * ticks are stored only as suggestions unless options.replaceExisting.
   */
  function draft(workbookId, taskName, submissionId, expectedRevision, options) {
    var opt = options || {};
    var target = WebGrading.openTask(workbookId, taskName);
    var ss = target.ss;
    var actor = opt.actor || WebAccess.requireTeacher();
    var sub = WebGrading.selectSubmission(ss, target.task, submissionId);
    if (['Approved', 'Locked', 'Superseded'].indexOf(text(sub.Status)) !== -1) throw new Error('This assessment is approved or locked; AI drafting is not allowed.');
    if (text(expectedRevision) !== WebGrading.lastRevision(ss, submissionId)) throw new Error('A newer assessment exists. Reload this student before drafting.');
    var started = Date.now();
    var out = proposeForSubmission(ss, target.task, submissionId, { actor: actor });
    WebGrading.upgrade(ss);
    if (out.files) EvidenceService.recordStatuses(ss, out.files);
    if (!out.success) {
      logRun(ss, submissionId, '', out, 'Failed', out.message, started, actor);
      return { success: false, unavailable: !!out.unavailable, message: out.message,
        files: (out.files || []).map(function (f) { return { fileName: f.fileName, status: f.status, limitations: f.limitations }; }) };
    }
    var runId = Utilities.getUuid();
    EvidenceService.recordVersions(ss, submissionId, target.task, out.files, runId, actor);
    var current = out.assessment.states;
    var teacherHasSelections = Object.keys(current).some(function (cid) { return AssessmentFlags.ticked(current[cid].checkboxes).length; });
    var apply = !teacherHasSelections || opt.replaceExisting === true;
    var incoming = {};
    out.profile.criteria.forEach(function (c) {
      var st = current[c.criterionId] || {};
      incoming[c.criterionId] = { checkboxes: apply ? out.proposals[c.criterionId].checks : (st.checkboxes || {}),
        teacherWrittenNote: st.teacherWrittenNote || '', review: st.review || {} };
    });
    var saved = WebGrading.saveDraft(workbookId, target.task, submissionId, incoming, out.assessment.feedback, expectedRevision, 'AI',
      { ai: out.proposals, actor: actor });
    logRun(ss, submissionId, saved.assessmentId, out, 'Success', '', started, actor);
    var unreadable = out.files.filter(function (f) { return f.status !== 'Reviewed'; });
    return { success: true, applied: apply, revision: saved.revision, flags: saved.flags,
      message: (apply ? 'AI draft saved' : 'AI suggestions saved alongside your selections (your ticks were kept)') +
        ' using ' + out.provider + '. Review every criterion before approval.' +
        (unreadable.length ? ' ' + unreadable.length + ' file(s) were not read.' : ''),
      files: out.files.map(function (f) { return { fileName: f.fileName, status: f.status, limitations: f.limitations }; }) };
  }

  function status() {
    WebAccess.requireTeacher();
    var settings = AppConfig.readSettings().aiProviders;
    return { success: true, data: { settings: settings, providers: WebAccess.providerStatus(), health: readHealth(), activePrompt: PromptRegistry.active().id } };
  }

  return { parseModelJson: parseModelJson, validateProposal: validateProposal, callOpenAiCompatible: callOpenAiCompatible,
    healthCheck: healthCheck, readHealth: readHealth, proposeForSubmission: proposeForSubmission, draft: draft, status: status,
    identityFor: identityFor };
})();

function apiWebAiStatus() { return AiGateway.status(); }
