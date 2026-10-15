/* ============================================================================
 * AssessmentFlags.gs — pure rules deciding what must be resolved before a
 * teacher can approve, and what is merely advisory. (SC22, SC23, SC27, SC32,
 * SC38; REVIEW.md F12, F13)
 *
 * BLOCKING (approval refused until resolved):
 *   NO_SELECTION            criterion has no ticked observable and the teacher has not
 *                           confirmed "no creditable evidence" (0 marks)
 *   AI_INSUFFICIENT_EVIDENCE AI found no supported evidence for the criterion and the
 *                           teacher has not marked the criterion reviewed
 *   UNREADABLE_EVIDENCE     one or more submitted files could not be read or are
 *                           unsupported, and the teacher has not acknowledged checking them
 *   FEEDBACK_NOT_REVIEWED   feedback or teacher notes exist but the feedback was written
 *                           before the latest change to selections/notes
 * ADVISORY (shown, never blocks):
 *   CAPPED                  ticks in several bands exceeded the criterion maximum
 *   AI_UNSUPPORTED_TICKS    AI ticks were discarded because no evidence was cited
 *   TEACHER_CHANGED_AI      final selections differ from the AI proposal
 *   FEEDBACK_EMPTY          no student feedback has been written
 *   EVIDENCE_CHANGED        Classroom evidence changed after the draft was saved
 *
 * Load-time pure and shared by the browser (mirrored, see tests/flags.test.cjs).
 * ============================================================================ */
var AssessmentFlags = (function () {
  var UNREADABLE = { Unsupported: true, Unreadable: true, Inaccessible: true, TooLarge: true };

  function ticked(checks) {
    return Object.keys(checks || {}).filter(function (k) { return checks[k] === true || checks[k] === 'true'; }).sort();
  }

  /* Deterministic djb2 hash of a canonical (sorted-key) JSON string. Used to
     tell whether feedback was written against the current assessment. */
  function canonical(value) {
    if (Array.isArray(value)) return '[' + value.map(canonical).join(',') + ']';
    if (value && typeof value === 'object') {
      return '{' + Object.keys(value).sort().map(function (k) { return JSON.stringify(k) + ':' + canonical(value[k]); }).join(',') + '}';
    }
    return JSON.stringify(value === undefined ? null : value);
  }

  function hash(value) {
    var str = canonical(value), h = 5381;
    for (var i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
    return 'h' + (h >>> 0).toString(36);
  }

  /** The assessment facts feedback must agree with: selections, notes, observation. */
  function feedbackBasis(profile, states, generalObservation) {
    var basis = { o: String(generalObservation || '').trim() };
    (profile && profile.criteria || []).forEach(function (c) {
      var st = (states || {})[c.criterionId] || {};
      basis[c.criterionId] = { t: ticked(st.checkboxes), n: String(st.teacherWrittenNote || '').trim(), z: !!(st.review && st.review.noEvidence) };
    });
    return hash(basis);
  }

  function hasText(fb) {
    return ['whatWentWell', 'areasForImprovement', 'goalsForNextAssessment'].some(function (k) { return String((fb || {})[k] || '').trim(); });
  }

  /**
   * @param {{profile:Object, states:Object, scores:Object, attachments:Array, feedback:Object}} input
   * @return {{flags:Array, blocking:number, advisory:number, byCriterion:Object}}
   */
  function compute(input) {
    var profile = input.profile;
    var states = input.states || {};
    var scores = input.scores || { criteria: {} };
    var fb = input.feedback || {};
    var flags = [];
    function add(code, severity, criterionId, message) {
      flags.push({ code: code, severity: severity, criterionId: criterionId || '', message: message });
    }
    if (!profile) return { flags: [], blocking: 0, advisory: 0, byCriterion: {} };

    var anyNote = false;
    profile.criteria.forEach(function (c) {
      var st = states[c.criterionId] || {};
      var review = st.review || {};
      var picks = ticked(st.checkboxes);
      var ai = st.aiProposal || null;
      if (String(st.teacherWrittenNote || '').trim()) anyNote = true;
      if (!picks.length && !review.noEvidence) {
        add('NO_SELECTION', 'blocking', c.criterionId, c.title + ': tick the evidenced observables, or confirm there is no creditable evidence (0 marks).');
      }
      if (ai && ai.insufficient && !review.reviewed && !review.noEvidence) {
        add('AI_INSUFFICIENT_EVIDENCE', 'blocking', c.criterionId, c.title + ': the AI found no supported evidence. Check the work yourself, then mark this criterion reviewed.');
      }
      if (ai && ai.unsupported && ai.unsupported.length) {
        add('AI_UNSUPPORTED_TICKS', 'advisory', c.criterionId, c.title + ': ' + ai.unsupported.length + ' AI suggestion(s) discarded because no evidence was cited.');
      }
      if (ai && ai.checks && canonical(ticked(ai.checks)) !== canonical(picks)) {
        add('TEACHER_CHANGED_AI', 'advisory', c.criterionId, c.title + ': your selections differ from the AI proposal.');
      }
      var sc = scores.criteria && scores.criteria[c.criterionId];
      if (sc && sc.rawPoints > sc.maxMarks + 1e-9) {
        add('CAPPED', 'advisory', c.criterionId, c.title + ': ticks across bands total ' + sc.rawPoints.toFixed(2) + ', capped at ' + sc.maxMarks + '.');
      }
    });

    var unreadable = (input.attachments || []).filter(function (f) { return UNREADABLE[f.status]; });
    if (unreadable.length && !fb.acknowledgedUnreadable) {
      add('UNREADABLE_EVIDENCE', 'blocking', '', unreadable.length + ' file(s) could not be read automatically (' +
        unreadable.map(function (f) { return f.fileName; }).slice(0, 3).join(', ') + '). Open and check them, then acknowledge.');
    }
    var current = feedbackBasis(profile, states, fb.generalObservation);
    if ((hasText(fb) || anyNote) && fb.basisHash !== current) {
      add('FEEDBACK_NOT_REVIEWED', 'blocking', '', 'Feedback was written before the latest change to selections or notes. Update it, or confirm it still matches.');
    }
    if (!hasText(fb)) add('FEEDBACK_EMPTY', 'advisory', '', 'No student feedback has been written.');
    if (input.evidenceChangedAt && !input.final) {
      add('EVIDENCE_CHANGED', 'advisory', '', 'Classroom evidence changed after this draft was saved. Re-check the work before approving.');
    }

    var byCriterion = {};
    flags.forEach(function (f) { if (f.criterionId) (byCriterion[f.criterionId] = byCriterion[f.criterionId] || []).push(f); });
    return {
      flags: flags,
      blocking: flags.filter(function (f) { return f.severity === 'blocking'; }).length,
      advisory: flags.filter(function (f) { return f.severity === 'advisory'; }).length,
      byCriterion: byCriterion,
      feedbackBasis: current
    };
  }

  return { compute: compute, feedbackBasis: feedbackBasis, hash: hash, canonical: canonical, ticked: ticked, UNREADABLE: UNREADABLE };
})();
