/* ============================================================================
 * FeedbackBuilder.gs — deterministic, grounded student feedback. (SC31–SC35)
 *
 * Built ONLY from: observables the teacher has ticked, the next-band
 * observables not yet evidenced, the teacher's own observations, and the
 * calculated result. Nothing is invented, no AI is involved, so it works
 * during an AI outage. The teacher edits the result before approval.
 * Load-time pure.
 * ============================================================================ */
var FeedbackBuilder = (function () {
  function text(v) { return v == null ? '' : String(v); }
  function sentence(s) { var t = text(s).trim().replace(/\s+/g, ' '); if (!t) return ''; t = t.charAt(0).toUpperCase() + t.slice(1); return /[.!?]$/.test(t) ? t : t + '.'; }
  function lowerFirst(s) { var t = text(s).trim(); return t ? t.charAt(0).toLowerCase() + t.slice(1).replace(/[.]$/, '') : ''; }

  function firstName(fullName) {
    var parts = text(fullName).trim().split(/\s+/);
    return parts[0] && !/^student$/i.test(parts[0]) ? parts[0] : '';
  }

  /* Highest band (best first) with at least one tick, and the band above it. */
  function bandPosition(c, checks) {
    for (var b = 0; b < c.bandOrder.length; b++) {
      var items = c.bands[c.bandOrder[b]] || [];
      if (items.some(function (o) { return checks[o.id]; })) return b;
    }
    return c.bandOrder.length; // nothing ticked
  }

  /**
   * @param {Object} profile pinned rubric profile
   * @param {Object} states {criterionId:{checkboxes, teacherWrittenNote, review}}
   * @param {Object} scores RubricEngine.scoreProfile result
   * @param {{generalObservation?:string}} extra
   * @return {{whatWentWell:string, areasForImprovement:string, goalsForNextAssessment:string, sources:Object}}
   */
  function build(profile, states, scores, extra) {
    var well = [], improve = [], goalPool = [];
    var sources = { ticked: 0, targets: 0, notes: 0 };
    profile.criteria.forEach(function (c) {
      var st = states[c.criterionId] || {};
      var checks = st.checkboxes || {};
      var res = scores.criteria[c.criterionId] || { percent: 0 };
      var pos = bandPosition(c, checks);
      if (checks[c.mdOverride.id]) {
        well.push(c.title + ': ' + sentence(c.mdOverride.text.replace(/^Distinct evidence:\s*/i, '')));
        sources.ticked++;
      } else if (pos < c.bandOrder.length) {
        var ticked = (c.bands[c.bandOrder[pos]] || []).filter(function (o) { return checks[o.id]; }).slice(0, 3);
        well.push(c.title + ': ' + ticked.map(function (o) { return sentence(o.text); }).join(' '));
        sources.ticked += ticked.length;
      }
      var note = text(st.teacherWrittenNote).trim();
      if (note) { well.push(c.title + ' (observed): ' + sentence(note)); sources.notes++; }
      if (res.percent < 100 - 1e-9) {
        var targetBand = pos > 0 ? c.bandOrder[Math.min(pos, c.bandOrder.length) - 1] : null;
        var sameBandMissing = pos < c.bandOrder.length ? (c.bands[c.bandOrder[pos]] || []).filter(function (o) { return !checks[o.id]; }) : [];
        var targets = (targetBand ? (c.bands[targetBand] || []) : sameBandMissing).filter(function (o) { return !checks[o.id]; }).slice(0, 2);
        if (!targets.length) targets = sameBandMissing.slice(0, 2);
        if (pos === c.bandOrder.length && st.review && st.review.noEvidence) {
          improve.push(c.title + ': no evidence was submitted for this part of the task.');
        }
        if (targets.length) {
          improve.push(c.title + ': next, aim to show that you ' + targets.map(function (o) { return lowerFirst(o.text); }).join(', and ') + '.');
          goalPool.push({ percent: res.percent, title: c.title, target: targets[0] });
          sources.targets += targets.length;
        }
      }
    });
    var observation = text(extra && extra.generalObservation).trim();
    if (observation) { well.push('Observed in class: ' + sentence(observation)); sources.notes++; }
    goalPool.sort(function (a, b) { return a.percent - b.percent; });
    var goals = goalPool.slice(0, 2).map(function (g, i) { return (i + 1) + '. ' + g.title + ' — ' + sentence(g.target.text); });
    return { whatWentWell: well.join('\n'), areasForImprovement: improve.join('\n'), goalsForNextAssessment: goals.join('\n'), sources: sources };
  }

  function gradeDisplay(mode, r) {
    var letter = text(r.letter), pts = Number(r.points || 0).toFixed(1).replace(/\.0$/, ''), max = text(r.maxMarks);
    var pct = Number(r.percent || 0).toFixed(1).replace(/\.0$/, '') + '%';
    switch (mode) {
      case 'letter': return letter;
      case 'mark': return pts + ' / ' + max;
      case 'percent': return pct;
      case 'letter+percent': return letter + ' (' + pct + ')';
      default: return letter + ' (' + pts + ' / ' + max + ')';
    }
  }

  /**
   * Renders the student-facing comment using the configured template.
   * The first name, when enabled, is inserted ONCE, in the opening line.
   */
  function render(template, feedback, ctx, outputFormat) {
    var fmt = outputFormat || {};
    var name = fmt.useFirstName === false ? '' : firstName(ctx.studentName);
    var opening = name ? name + ', here is your feedback for ' + text(ctx.taskName) + '.' : 'Here is your feedback for ' + text(ctx.taskName) + '.';
    var display = gradeDisplay(fmt.gradeDisplay, ctx);
    var t = text(template) || '{{opening}}\n\nResult: {{grade_display}}\n\n{{what_went_well}}\n\n{{areas_for_improvement}}\n\n{{goals}}';
    var warnings = [];
    var out = t.replace(/\{\{([a-z_]+)\}\}/g, function (m, key) {
      switch (key) {
        case 'opening': return opening;
        case 'first_name': return name;
        case 'student': return name; // never the full name in student-facing text
        case 'task': return text(ctx.taskName);
        case 'grade': return text(ctx.letter);
        case 'grade_display': return fmt.includeGradeInComment === false ? '' : display;
        case 'final_score': case 'score': return Number(ctx.points || 0).toFixed(1);
        case 'max_marks': return text(ctx.maxMarks);
        case 'percent': return Number(ctx.percent || 0).toFixed(1) + '%';
        case 'what_went_well': return text(feedback.whatWentWell);
        case 'areas_for_improvement': return text(feedback.areasForImprovement);
        case 'goals': return text(feedback.goalsForNextAssessment);
        default: warnings.push('Unknown template field {{' + key + '}}'); return m;
      }
    }).replace(/\n{3,}/g, '\n\n').replace(/Result:\s*\n/g, '').trim();
    if (name) {
      var count = (out.match(new RegExp('\\b' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'g')) || []).length;
      if (count > 1) warnings.push('The first name appears ' + count + ' times; once in the opening is usually enough.');
    }
    return { text: out, warnings: warnings, gradeDisplay: display };
  }

  return { build: build, render: render, firstName: firstName, gradeDisplay: gradeDisplay };
})();
