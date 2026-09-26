/* ============================================================================
 * RubricEngine.gs — Universal, task-agnostic rubric profile + scoring engine.
 *
 * Purpose: remove the hardcoded Year 9 Jewellery dependency from grading.
 * A "rubric profile" is a JSON document describing ANY task's criteria,
 * bands and observables. Profiles are built from:
 *   - a Shared Rubric spreadsheet (Rubric tab rows),
 *   - an uploaded PDF/Markdown/TXT rubric (via Gemini extraction upstream),
 *   - the legacy workbook CriteriaConfig sheet (one-click adoption), or
 *   - the legacy hardcoded Config.CRITERIA_DEFINITIONS (seeded profile).
 *
 * Design rules:
 *   - ALL scoring maths lives here as pure functions (unit-tested in
 *     tests/rubric-engine.test.cjs). AI never computes marks.
 *   - Sheet-touching functions take an explicit `ss` (Spreadsheet) argument,
 *     so the engine is workbook-scoped and safe from any context (web app,
 *     dialog, trigger). No SpreadsheetApp.getActiveSpreadsheet() in here.
 *   - This file must stay load-time pure (no GAS calls at parse time).
 *
 * Profile JSON schema (SCHEMA_VERSION 1.0):
 * {
 *   schemaVersion, profileId, taskName, name, source, sourceRubricFileId,
 *   totalMaxMarks, parts: [..],
 *   criteria: [{
 *     criterionId, title, part, section, maxMarks, outcome,
 *     mdOverride: {id, text},
 *     bandOrder: ['A','B','C','D','E'] (best -> worst, may contain other keys),
 *     bandWeights: {A: 1.0, ...},
 *     bands: {A: [{id, text}, ...], ...}
 *   }],
 *   createdAt, createdBy
 * }
 * ============================================================================ */
var RubricEngine = (function () {
  var SCHEMA_VERSION = '1.0';
  var CORE_BANDS = ['A', 'B', 'C', 'D', 'E'];
  var DEFAULT_WEIGHTS = { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 };
  var DEFAULT_BANDS = [
    { letter: 'A', min: 85, max: 100 },
    { letter: 'B', min: 75, max: 84.99 },
    { letter: 'C', min: 65, max: 74.99 },
    { letter: 'D', min: 50, max: 64.99 },
    { letter: 'E', min: 0, max: 49.99 }
  ];

  var PROFILES_HEADERS = [
    'ProfileID', 'TaskName', 'ProfileName', 'Source', 'SourceRubricFileID',
    'JSONDefinition', 'SchemaVersion', 'Active', 'CreatedAt', 'CreatedBy'
  ];

  /* ------------------------------------------------------------------ *
   * Pure helpers
   * ------------------------------------------------------------------ */

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function cidFromIndex(i) { return 'C' + pad2(i + 1); }

  function toNumber(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = parseFloat(String(v == null ? '' : v).replace(/[^\d.\-]/g, ''));
    return isFinite(n) ? n : 0;
  }

  function normalizeGradeScale(weights, bands) {
    var w = {};
    var src = weights || {};
    for (var i = 0; i < CORE_BANDS.length; i++) {
      var L = CORE_BANDS[i];
      w[L] = (typeof src[L] === 'number' && isFinite(src[L])) ? src[L] : DEFAULT_WEIGHTS[L];
    }
    var b = [];
    if (bands && bands.length) {
      for (var k = 0; k < bands.length; k++) {
        var row = bands[k];
        if (!row || !row.letter) continue;
        b.push({
          letter: String(row.letter).toUpperCase(),
          min: toNumber(row.min),
          max: row.max == null ? 100 : toNumber(row.max)
        });
      }
      b.sort(function (x, y) { return y.min - x.min; });
    }
    if (!b.length) b = DEFAULT_BANDS.slice();
    return { weights: w, bands: b };
  }

  function overallLetter(percent, bands) {
    var b = (bands && bands.length) ? bands : DEFAULT_BANDS;
    for (var i = 0; i < b.length; i++) {
      if (percent >= b[i].min) return b[i].letter;
    }
    return b.length ? b[b.length - 1].letter : 'E';
  }

  /* NEVER invent a weight for an unfamiliar band label (e.g. "Level 4").
     Such rubrics require Dan to supply an explicit 0..1 weight when reviewing. */
  function bandWeightsFor(bandOrder, weights) {
    var w = {};
    for (var i = 0; i < bandOrder.length; i++) {
      var key = bandOrder[i];
      if (weights && typeof weights[key] === 'number' && isFinite(weights[key])) {
        w[key] = weights[key];
      } else if (DEFAULT_WEIGHTS[key] !== undefined) {
        w[key] = DEFAULT_WEIGHTS[key];
      }
    }
    return w;
  }

  /* Order band keys best -> worst: core letters first (A..E when present),
     then any other keys in order of first appearance. */
  function orderBandKeys(seenKeys) {
    var ordered = [];
    for (var i = 0; i < CORE_BANDS.length; i++) {
      if (seenKeys.indexOf(CORE_BANDS[i]) !== -1) ordered.push(CORE_BANDS[i]);
    }
    for (var k = 0; k < seenKeys.length; k++) {
      if (ordered.indexOf(seenKeys[k]) === -1) ordered.push(seenKeys[k]);
    }
    return ordered;
  }

  /* ------------------------------------------------------------------ *
   * Profile construction (pure)
   * ------------------------------------------------------------------ */

  /* rows: [{criterion, part, section, maxMarks, outcome, band, description}] */
  function buildProfileFromRubricRows(rows, meta) {
    meta = meta || {};
    var order = [];
    var byTitle = Object.create(null);
    var weights = meta.weights || DEFAULT_WEIGHTS;
    if (!Array.isArray(rows) || rows.length > 2000) throw new Error('Supply at most 2,000 rubric observable rows.');

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || {};
      var title = String(r.criterion == null ? '' : r.criterion).trim();
      var band = String(r.band == null ? '' : r.band).trim().toUpperCase();
      var desc = String(r.description == null ? '' : r.description).trim();
      if (!title && !band && !desc) continue; // blank spreadsheet row
      if (!title || !band || !desc) throw new Error('Rubric row ' + (i + 1) + ' needs a criterion, band and observable.');
      if (title.length > 180 || band.length > 24 || desc.length > 1500) throw new Error('Rubric row ' + (i + 1) + ' is too long.');

      /* The same criterion title can appear in two parts of a project.
         Identity is title + part + section, not just the text of a label. */
      var part = String(r.part == null ? '' : r.part).trim();
      var section = String(r.section == null ? '' : r.section).trim();
      var key = JSON.stringify([part, section, title]);
      if (!byTitle[key]) {
        byTitle[key] = {
          title: title,
          part: part,
          section: section,
          maxMarks: toNumber(r.maxMarks),
          outcome: String(r.outcome == null ? '' : r.outcome).trim(),
          bandKeysSeen: [],
          bands: {}
        };
        order.push(key);
      }
      var c = byTitle[key];
      var marks = toNumber(r.maxMarks);
      if (marks && c.maxMarks && marks !== c.maxMarks) throw new Error('Inconsistent marks for ' + title + '. Review the imported rubric.');
      if (!c.maxMarks) c.maxMarks = marks;
      if (!c.part && r.part) c.part = String(r.part).trim();
      if (!c.section && r.section) c.section = String(r.section).trim();
      if (!c.outcome && r.outcome) c.outcome = String(r.outcome).trim();
      if (c.bandKeysSeen.indexOf(band) === -1) c.bandKeysSeen.push(band);
      if (!c.bands[band]) c.bands[band] = [];
      c.bands[band].push(desc);
    }

    var criteria = [];
    var parts = [];
    var totalMax = 0;
    for (var t = 0; t < order.length; t++) {
      var src = byTitle[order[t]];
      var cid = cidFromIndex(t);
      var bandOrder = orderBandKeys(src.bandKeysSeen);
      var bands = {};
      for (var b = 0; b < bandOrder.length; b++) {
        var key = bandOrder[b];
        var items = src.bands[key] || [];
        bands[key] = [];
        for (var o = 0; o < items.length; o++) {
          bands[key].push({ id: cid + '-' + key + '-' + (o + 1), text: items[o] });
        }
      }
      if (src.part && parts.indexOf(src.part) === -1) parts.push(src.part);
      totalMax += src.maxMarks;
      criteria.push({
        criterionId: cid,
        title: src.title,
        part: src.part,
        section: src.section,
        maxMarks: src.maxMarks,
        outcome: src.outcome,
        mdOverride: {
          id: cid + '-MD',
          text: 'Distinct evidence: outstanding work at the highest band standard, beyond the listed observables.'
        },
        bandOrder: bandOrder,
        bandWeights: bandWeightsFor(bandOrder, weights),
        bands: bands
      });
    }

    return {
      schemaVersion: SCHEMA_VERSION,
      profileId: meta.profileId || '',
      taskName: meta.taskName || '',
      name: meta.name || (criteria.length ? criteria[0].title : 'Untitled rubric'),
      source: meta.source || 'rubricFile',
      sourceRubricFileId: meta.sourceRubricFileId || '',
      totalMaxMarks: totalMax,
      parts: parts,
      gradeScaleBands: meta.bands || DEFAULT_BANDS,
      criteria: criteria,
      createdAt: meta.createdAt || '',
      createdBy: meta.createdBy || ''
    };
  }

  /* Legacy Config.CRITERIA_DEFINITIONS -> profile (keeps the jewellery
     workbook bit-identical; observable IDs and MD text are preserved). */
  function profileFromLegacyDefinitions(defs, meta) {
    meta = meta || {};
    var weights = normalizeGradeScale(meta.weights, null).weights;
    var criteria = [];
    var parts = [];
    var totalMax = 0;
    for (var i = 0; i < (defs || []).length; i++) {
      var d = defs[i];
      var seen = [];
      var bands = {};
      for (var b = 0; b < CORE_BANDS.length; b++) {
        var L = CORE_BANDS[b];
        var items = d.bands && d.bands[L] ? d.bands[L] : [];
        if (!items.length) continue;
        seen.push(L);
        bands[L] = [];
        for (var o = 0; o < items.length; o++) {
          bands[L].push({ id: items[o][0], text: items[o][1] });
        }
      }
      if (d.part && parts.indexOf(d.part) === -1) parts.push(d.part);
      totalMax += toNumber(d.maxMarks);
      criteria.push({
        criterionId: d.criterionId,
        title: d.title,
        part: d.part || '',
        section: d.section || '',
        maxMarks: toNumber(d.maxMarks),
        outcome: d.outcome || '',
        mdOverride: { id: d.mdOverride.id, text: d.mdOverride.text },
        bandOrder: seen,
        bandWeights: bandWeightsFor(seen, weights),
        bands: bands
      });
    }
    return {
      schemaVersion: SCHEMA_VERSION,
      profileId: meta.profileId || '',
      taskName: meta.taskName || '',
      name: meta.name || 'Legacy criteria',
      source: meta.source || 'legacy',
      sourceRubricFileId: meta.sourceRubricFileId || '',
      totalMaxMarks: totalMax,
      parts: parts,
      gradeScaleBands: meta.bands || DEFAULT_BANDS,
      criteria: criteria,
      createdAt: meta.createdAt || '',
      createdBy: meta.createdBy || ''
    };
  }

  /* Legacy flat CriteriaConfig sheet rows -> profile. */
  function profileFromCriteriaConfigRows(rows, meta) {
    meta = meta || {};
    var order = [];
    var byId = {};
    for (var i = 0; i < (rows || []).length; i++) {
      var r = rows[i] || {};
      var cid = String(r.CriterionID || '').trim();
      var label = String(r.CheckboxLabel || '').trim();
      var band = String(r.Band || '').trim().toUpperCase();
      if (!cid || !band || !label) continue;
      if (!byId[cid]) {
        byId[cid] = {
          criterionId: cid,
          title: String(r.CriterionTitle || '').trim(),
          part: String(r.Part || '').trim(),
          section: String(r.Section || '').trim(),
          maxMarks: toNumber(r.MaxMarks),
          outcome: String(r.Outcome || '').trim(),
          bandKeysSeen: [],
          bands: {},
          mdOverride: null
        };
        order.push(cid);
      }
      var c = byId[cid];
      if (String(r.IsMissingDistinctOverride).toLowerCase() === 'true') {
        c.mdOverride = { id: String(r.CheckboxID || (cid + '-MD')).trim(), text: label };
        continue;
      }
      if (c.bandKeysSeen.indexOf(band) === -1) c.bandKeysSeen.push(band);
      if (!c.bands[band]) c.bands[band] = [];
      c.bands[band].push({ id: String(r.CheckboxID || (cid + '-' + band + '-' + c.bands[band].length)).trim(), text: label });
    }

    var weights = normalizeGradeScale(meta.weights, null).weights;
    var criteria = [];
    var parts = [];
    var totalMax = 0;
    for (var t = 0; t < order.length; t++) {
      var src = byId[order[t]];
      var bandOrder = orderBandKeys(src.bandKeysSeen);
      var bands = {};
      for (var b = 0; b < bandOrder.length; b++) bands[bandOrder[b]] = src.bands[bandOrder[b]] || [];
      if (!src.mdOverride) {
        src.mdOverride = {
          id: src.criterionId + '-MD',
          text: 'Distinct evidence: outstanding work at the highest band standard, beyond the listed observables.'
        };
      }
      if (src.part && parts.indexOf(src.part) === -1) parts.push(src.part);
      totalMax += src.maxMarks;
      criteria.push({
        criterionId: src.criterionId,
        title: src.title,
        part: src.part,
        section: src.section,
        maxMarks: src.maxMarks,
        outcome: src.outcome,
        mdOverride: src.mdOverride,
        bandOrder: bandOrder,
        bandWeights: bandWeightsFor(bandOrder, weights),
        bands: bands
      });
    }

    return {
      schemaVersion: SCHEMA_VERSION,
      profileId: meta.profileId || '',
      taskName: meta.taskName || '',
      name: meta.name || 'CriteriaConfig import',
      source: meta.source || 'criteriaConfig',
      sourceRubricFileId: meta.sourceRubricFileId || '',
      totalMaxMarks: totalMax,
      parts: parts,
      gradeScaleBands: meta.bands || DEFAULT_BANDS,
      criteria: criteria,
      createdAt: meta.createdAt || '',
      createdBy: meta.createdBy || ''
    };
  }

  function validateProfile(profile) {
    var errors = [];
    if (!profile || typeof profile !== 'object') return { ok: false, errors: ['Profile is not an object.'] };
    if (!profile.taskName || !String(profile.taskName).trim()) errors.push('Choose a task for this rubric.');
    if (!profile.name || !String(profile.name).trim()) errors.push('Name the rubric.');
    if (!Array.isArray(profile.criteria) || !profile.criteria.length || profile.criteria.length > 50) {
      errors.push('Rubric must contain 1 to 50 criteria.');
    }
    var ids = Object.create(null);
    var obsIds = Object.create(null);
    var total = 0;
    for (var i = 0; i < (profile.criteria || []).length; i++) {
      var c = profile.criteria[i];
      if (!c.criterionId) errors.push('Criterion at index ' + i + ' has no id.');
      if (ids[c.criterionId]) errors.push('Duplicate criterion id: ' + c.criterionId);
      ids[c.criterionId] = true;
      if (!c.title) errors.push(c.criterionId + ': missing title.');
      if (!(typeof c.maxMarks === 'number' && isFinite(c.maxMarks) && c.maxMarks > 0 && c.maxMarks <= 1000)) {
        errors.push(c.criterionId + ': enter a mark allocation between 0 and 1,000.');
      }
      total += c.maxMarks;
      if (!c.bandOrder || !c.bandOrder.length) errors.push(c.criterionId + ': no bands.');
      if (!c.mdOverride || !c.mdOverride.id) errors.push(c.criterionId + ': missing distinct override.');
      if (c.mdOverride && c.mdOverride.id) obsIds[c.mdOverride.id] = true;
      for (var b = 0; b < (c.bandOrder || []).length; b++) {
        var label = c.bandOrder[b];
        var items = (c.bands && c.bands[label]) || [];
        var weight = c.bandWeights && c.bandWeights[label];
        if (!(typeof weight === 'number' && isFinite(weight) && weight >= 0 && weight <= 1)) {
          errors.push(c.criterionId + ': set an explicit 0–1 weight for band "' + label + '".');
        }
        if (!items.length) errors.push(c.criterionId + ': band ' + label + ' is empty.');
        for (var o = 0; o < items.length; o++) {
          if (!items[o].id || !items[o].text) errors.push(c.criterionId + ': observable without id/text.');
          if (obsIds[items[o].id]) errors.push('Duplicate observable id: ' + items[o].id);
          obsIds[items[o].id] = true;
        }
      }
    }
    if (Math.abs(total - profile.totalMaxMarks) > 0.001) errors.push('Rubric total does not match criterion marks.');
    return { ok: errors.length === 0, errors: errors };
  }

  function profileSummary(profile) {
    if (!profile) return null;
    var observableCount = 0;
    for (var i = 0; i < profile.criteria.length; i++) {
      var c = profile.criteria[i];
      for (var b = 0; b < c.bandOrder.length; b++) {
        observableCount += (c.bands[c.bandOrder[b]] || []).length;
      }
    }
    return {
      profileId: profile.profileId,
      taskName: profile.taskName,
      name: profile.name,
      source: profile.source,
      criteriaCount: profile.criteria.length,
      observableCount: observableCount,
      totalMaxMarks: profile.totalMaxMarks,
      parts: profile.parts
    };
  }

  /* ------------------------------------------------------------------ *
   * Scoring (pure, deterministic — mirrors legacy deriveGradeFromChecks)
   * ------------------------------------------------------------------ */

  function scoreCriterion(criterion, checksMap, gradeScale) {
    var scale = gradeScale || normalizeGradeScale(null, null);
    var bands = scale.bands;
    var checks = checksMap || {};
    var maxMarks = toNumber(criterion.maxMarks);
    var bandOrder = criterion.bandOrder || [];
    var weights = criterion.bandWeights || bandWeightsFor(bandOrder, scale.weights);

    function ticked(id) { return checks[id] === true || checks[id] === 'true'; }

    if (criterion.mdOverride && ticked(criterion.mdOverride.id)) {
      return {
        grade: overallLetter(100, bands),
        score: 1,
        points: maxMarks,
        maxMarks: maxMarks,
        percent: 100,
        explanation: 'Distinct top-band evidence selected — full marks awarded.',
        bandBreakdown: {},
        incompleteMinimumEvidence: false,
        distinctOverride: true
      };
    }

    var totalPoints = 0;
    var totalTicked = 0;
    var breakdown = {};
    for (var b = 0; b < bandOrder.length; b++) {
      var key = bandOrder[b];
      var items = criterion.bands[key] || [];
      var tickCount = 0;
      for (var o = 0; o < items.length; o++) {
        if (ticked(items[o].id)) tickCount++;
      }
      breakdown[key] = { ticked: tickCount, total: items.length, fraction: items.length ? tickCount / items.length : 0 };
      totalTicked += tickCount;
      if (items.length) {
        if (typeof weights[key] !== 'number') throw new Error('No weight configured for band ' + key + '.');
        totalPoints += tickCount * ((maxMarks * weights[key]) / items.length);
      }
    }

    /* Multiple supported bands can sum above 100%. Cap instead of exporting
       an impossible mark; retain the uncapped value for teacher review. */
    var rawPoints = totalPoints;
    totalPoints = Math.min(maxMarks, totalPoints);
    var percent = maxMarks > 0 ? (totalPoints / maxMarks) * 100 : 0;
    return {
      grade: overallLetter(percent, bands),
      score: maxMarks > 0 ? totalPoints / maxMarks : 0,
      points: totalPoints,
      rawPoints: rawPoints,
      maxMarks: maxMarks,
      percent: percent,
      explanation: 'Total points: ' + totalPoints.toFixed(2) + ' / ' + maxMarks +
        ' (' + Math.round(percent) + '%) → nominal grade ' + overallLetter(percent, bands) +
        (rawPoints > maxMarks ? ' (multiple bands capped at full marks).' : '.'),
      bandBreakdown: breakdown,
      incompleteMinimumEvidence: totalTicked === 0,
      distinctOverride: false
    };
  }

  /* checksByCriterion: {C01: {checksMap}, ...} */
  function scoreProfile(profile, checksByCriterion, gradeScale) {
    var scale = gradeScale || normalizeGradeScale(null, profile.gradeScaleBands || null);
    var perCriterion = {};
    var partTotals = {};
    var totalPoints = 0;
    var totalMax = 0;
    var src = checksByCriterion || {};
    for (var i = 0; i < profile.criteria.length; i++) {
      var c = profile.criteria[i];
      var entry = src[c.criterionId] || {};
      var res = scoreCriterion(c, entry.checkboxes || entry.checks || {}, scale);
      perCriterion[c.criterionId] = res;
      var partKey = c.part || 'Overall';
      partTotals[partKey] = (partTotals[partKey] || 0) + res.points;
      totalPoints += res.points;
      totalMax += toNumber(c.maxMarks);
    }
    var percent = totalMax > 0 ? (totalPoints / totalMax) * 100 : 0;
    return {
      criteria: perCriterion,
      partTotals: partTotals,
      totalPoints: totalPoints,
      totalMaxMarks: totalMax,
      percent: percent,
      letter: overallLetter(percent, scale.bands)
    };
  }

  /* ------------------------------------------------------------------ *
   * AI prompt serialisation (pure)
   * ------------------------------------------------------------------ */

  function serializeForAi(profile) {
    var out = [];
    for (var i = 0; i < profile.criteria.length; i++) {
      var c = profile.criteria[i];
      var head = 'CRITERION ' + c.criterionId + ': ' + c.title + ' (' +
        (c.part ? c.part + ', ' : '') + '/' + c.maxMarks + ' Marks)\n';
      for (var b = 0; b < c.bandOrder.length; b++) {
        var key = c.bandOrder[b];
        var items = c.bands[key] || [];
        var parts = [];
        for (var o = 0; o < items.length; o++) parts.push('[' + items[o].id + '] ' + items[o].text);
        head += '  - Band ' + key + ': ' + parts.join('; ') + '\n';
      }
      head += '  - Distinct override: [' + c.mdOverride.id + '] ' + c.mdOverride.text + '\n';
      out.push(head);
    }
    return out.join('\n');
  }

  function knownObservableIds(profile) {
    var ids = {};
    for (var i = 0; i < profile.criteria.length; i++) {
      var c = profile.criteria[i];
      ids[c.mdOverride.id] = c.criterionId;
      for (var b = 0; b < c.bandOrder.length; b++) {
        var items = c.bands[c.bandOrder[b]] || [];
        for (var o = 0; o < items.length; o++) ids[items[o].id] = c.criterionId;
      }
    }
    return ids;
  }

  /* ------------------------------------------------------------------ *
   * Generic PDF report (pure string builder — no task-specific wording)
   * ------------------------------------------------------------------ */

  function renderReportHtml(profile, detail, gradeScale, meta) {
    meta = meta || {};
    var scale = gradeScale || normalizeGradeScale(null, profile.gradeScaleBands || null);
    var sub = detail.submission || {};
    var cMap = detail.criteriaMap || {};
    var fb = detail.feedback || {};
    var partTotals = {};
    var totalPoints = 0;
    var rowsHtml = '';

    for (var i = 0; i < profile.criteria.length; i++) {
      var c = profile.criteria[i];
      var cd = cMap[c.criterionId] || {};
      var res = scoreCriterion(c, cd.checkboxes || {}, scale);
      var partKey = c.part || 'Overall';
      partTotals[partKey] = (partTotals[partKey] || 0) + res.points;
      totalPoints += res.points;

      var note = cd.teacherWrittenNote || cd.teacherAudioTranscript || res.explanation;
      rowsHtml += '<tr>' +
        '<td style="font-weight:bold;padding:6px;border:1px solid #cbd5e1;">' + escapeHtml(c.criterionId) + '</td>' +
        '<td style="padding:6px;border:1px solid #cbd5e1;"><b>' + escapeHtml(c.title) + '</b><br/>' +
          '<span style="font-size:10px;color:#64748b;">' + escapeHtml((c.part ? c.part + ' \u2022 ' : '') + '/' + c.maxMarks + ' marks' + (c.outcome ? ' \u2022 ' + c.outcome : '')) + '</span></td>' +
        '<td style="text-align:center;font-weight:bold;padding:6px;border:1px solid #cbd5e1;">' + escapeHtml(res.grade) + '</td>' +
        '<td style="text-align:center;padding:6px;border:1px solid #cbd5e1;">' + res.points.toFixed(2) + ' / ' + c.maxMarks + '</td>' +
        '<td style="padding:11px;font-size:11px;border:1px solid #cbd5e1;">' + escapeHtml(note) + '</td>' +
      '</tr>';
    }

    var totalMax = profile.totalMaxMarks || totalPoints || 100;
    var percent = totalMax > 0 ? (totalPoints / totalMax) * 100 : 0;
    var letter = overallLetter(percent, scale.bands);

    var partBoxes = '';
    for (var p = 0; p < (profile.parts || []).length; p++) {
      var pk = profile.parts[p];
      partBoxes += '<div><b>' + escapeHtml(pk) + ':</b> ' + (partTotals[pk] || 0).toFixed(2) + '</div>';
    }

    return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<style>' +
      'body{font-family:Arial,Helvetica,sans-serif;margin:20px;color:#0f172a;line-height:1.35;font-size:12px;}' +
      'h1{margin:0;font-size:17px;color:#1e3a8a;}' +
      '.header-box{display:flex;justify-content:space-between;border-bottom:2px solid #1e3a8a;padding-bottom:10px;margin-bottom:14px;}' +
      '.score-card{background:#0f172a;color:#fff;padding:8px 14px;border-radius:6px;text-align:right;}' +
      '.details-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;background:#f8fafc;border:1px solid #cbd5e1;padding:8px 12px;border-radius:6px;margin-bottom:14px;font-size:11px;}' +
      'table{width:100%;border-collapse:collapse;margin-bottom:16px;}' +
      'th{background:#0f172a;color:#fff;padding:6px;font-size:11px;border:1px solid #0f172a;text-align:left;}' +
      '.fb-box{background:#f8fafc;border:1px solid #2563eb;border-radius:6px;padding:10px 14px;margin-top:10px;}' +
      '.fb-title{font-weight:bold;color:#1e3a8a;margin-bottom:4px;font-size:11px;}' +
      '</style></head><body>' +
      '<div class="header-box">' +
        '<div>' +
          '<h1>' + escapeHtml(meta.reportTitle || 'ASSESSMENT REPORT') + '</h1>' +
          '<div style="font-size:13px;font-weight:bold;color:#2563eb;margin-top:2px;">' + escapeHtml(meta.taskName || profile.taskName || profile.name || '') + '</div>' +
        '</div>' +
        '<div class="score-card">' +
          '<div style="font-size:10px;color:#94a3b8;">OVERALL RESULT</div>' +
          '<div style="font-size:18px;font-weight:bold;color:#38bdf8;">' + totalPoints.toFixed(1) + ' / ' + totalMax + ' (' + letter + ')</div>' +
        '</div>' +
      '</div>' +
      '<div class="details-grid">' +
        '<div><b>Student:</b> ' + escapeHtml(sub.StudentName || '') + '</div>' +
        '<div><b>Class:</b> ' + escapeHtml(sub.Class || meta.courseName || '') + '</div>' +
        '<div><b>Submission version:</b> v' + escapeHtml(sub.SubmissionVersion || 1) + '</div>' +
        partBoxes +
        '<div><b>Approved:</b> ' + escapeHtml(meta.date || '') + '</div>' +
      '</div>' +
      '<table>' +
        '<thead><tr><th>Code</th><th>Criterion</th><th style="text-align:center;">Grade</th><th style="text-align:center;">Mark</th><th>Teacher observations</th></tr></thead>' +
        '<tbody>' + rowsHtml + '</tbody>' +
      '</table>' +
      '<div class="fb-box">' +
        '<div style="font-size:13px;font-weight:bold;color:#0f172a;margin-bottom:8px;border-bottom:1px solid #cbd5e1;padding-bottom:4px;">Teacher Feedback &amp; Goals</div>' +
        '<div class="fb-title">What went well:</div>' +
        '<div style="margin-bottom:8px;white-space:pre-wrap;">' + escapeHtml(fb.whatWentWell || '') + '</div>' +
        '<div class="fb-title">Areas for improvement:</div>' +
        '<div style="margin-bottom:8px;white-space:pre-wrap;">' + escapeHtml(fb.areasForImprovement || '') + '</div>' +
        '<div class="fb-title">Goals for next assessment:</div>' +
        '<div style="white-space:pre-wrap;">' + escapeHtml(fb.goalsForNextAssessment || '') + '</div>' +
      '</div>' +
      '</body></html>';
  }

  /* ------------------------------------------------------------------ *
   * Workbook-scoped store (GAS calls confined inside functions)
   * ------------------------------------------------------------------ */

  function getGradeScale(ss) {
    var out = normalizeGradeScale(null, null);
    try {
      var sheet = ss.getSheetByName('Setup');
      if (!sheet) return out;
      var data = sheet.getDataRange().getValues();

      function findHeaderRow(keyword) {
        for (var r = 0; r < data.length; r++) {
          if (data[r].join(' ').toUpperCase().indexOf(keyword) !== -1) return r;
        }
        return -1;
      }
      function extractLetterAndNumbers(row) {
        var letter = null;
        var nums = [];
        for (var c = 0; c < row.length; c++) {
          var cell = row[c];
          if (letter === null && typeof cell === 'string' && /^[A-Z]$/i.test(cell.trim())) {
            letter = cell.trim().toUpperCase();
            continue;
          }
          if (typeof cell === 'number') nums.push(cell);
          else if (typeof cell === 'string' && /^-?\d+(\.\d+)?%?$/.test(cell.trim())) nums.push(parseFloat(cell.replace('%', '')));
        }
        return { letter: letter, nums: nums };
      }

      var wRow = findHeaderRow('GRADE SCALE');
      if (wRow !== -1) {
        var weights = {};
        for (var r = wRow + 1; r < data.length; r++) {
          if (!data[r].join('').trim()) break;
          var p = extractLetterAndNumbers(data[r]);
          if (p.letter && p.nums.length) weights[p.letter] = p.nums[0] > 1 ? p.nums[0] / 100 : p.nums[0];
        }
        if (Object.keys(weights).length) out.weights = normalizeGradeScale(weights, null).weights;
      }
      var bRow = findHeaderRow('OVERALL GRADE BAND');
      if (bRow !== -1) {
        var bands = [];
        for (var r2 = bRow + 1; r2 < data.length; r2++) {
          if (!data[r2].join('').trim()) break;
          var p2 = extractLetterAndNumbers(data[r2]);
          if (p2.letter && p2.nums.length >= 2) bands.push({ letter: p2.letter, min: p2.nums[0], max: p2.nums[1] });
          else if (p2.letter && p2.nums.length === 1) bands.push({ letter: p2.letter, min: p2.nums[0], max: 100 });
        }
        if (bands.length) out.bands = normalizeGradeScale(null, bands).bands;
      }
    } catch (err) {
      /* fall through to defaults */
    }
    return out;
  }

  function ensureProfilesSheet(ss) {
    var sheet = ss.getSheetByName('RubricProfiles');
    if (!sheet) {
      sheet = ss.insertSheet('RubricProfiles');
      sheet.getRange(1, 1, 1, PROFILES_HEADERS.length).setValues([PROFILES_HEADERS]);
      sheet.getRange(1, 1, 1, PROFILES_HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      return sheet;
    }
    /* Additive header migration (legacy sheet had only 4 columns). */
    var lastCol = sheet.getLastColumn();
    var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
    var missing = PROFILES_HEADERS.filter(function (n) { return headers.indexOf(n) === -1; });
    if (missing.length) {
      var start = headers.length + 1;
      var target = start + missing.length - 1;
      if (sheet.getMaxColumns() < target) sheet.insertColumnsAfter(sheet.getMaxColumns(), target - sheet.getMaxColumns());
      sheet.getRange(1, start, 1, missing.length).setValues([missing]);
    }
    return sheet;
  }

  function readProfiles(ss) {
    var sheet = ss.getSheetByName('RubricProfiles');
    if (!sheet || sheet.getLastRow() < 2) return []; // reads never migrate data
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function (h) { return String(h).trim(); });
    var out = [];
    for (var r = 1; r < data.length; r++) {
      var row = {};
      for (var c = 0; c < headers.length; c++) row[headers[c]] = data[r][c];
      if (!row.ProfileID) continue;
      var profile = null;
      try { profile = JSON.parse(String(row.JSONDefinition || '')); } catch (e) { profile = null; }
      out.push({ rowNumber: r + 1, row: row, profile: profile });
    }
    return out;
  }

  function getActiveProfile(ss, taskName) {
    var entries = readProfiles(ss);
    for (var i = entries.length - 1; i >= 0; i--) {
      var e = entries[i];
      var active = e.row.Active === true || String(e.row.Active).toUpperCase() === 'TRUE';
      if (active && String(e.row.TaskName) === String(taskName) && e.profile && validateProfile(e.profile).ok) return e.profile;
    }
    return null;
  }

  function getProfileById(ss, profileId) {
    var entries = readProfiles(ss);
    for (var i = 0; i < entries.length; i++) {
      if (String(entries[i].row.ProfileID) === String(profileId) && entries[i].profile && validateProfile(entries[i].profile).ok) return entries[i].profile;
    }
    return null;
  }

  function saveProfile(ss, profile, user) {
    var validation = validateProfile(profile);
    if (!validation.ok) return { success: false, message: 'Invalid rubric profile: ' + validation.errors.join(' ') };
    var sheet = ensureProfilesSheet(ss);
    var nowIso = new Date().toISOString();
    /* Server generates the ID: never let a browser overwrite a prior version. */
    profile.profileId = Utilities.getUuid();
    profile.createdAt = nowIso;
    profile.createdBy = user || '';

    var entries = readProfiles(ss);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    function column(name) { return headers.indexOf(name) + 1; }
    var values = {
      ProfileID: profile.profileId, TaskName: profile.taskName, ProfileName: profile.name,
      Source: profile.source, SourceRubricFileID: profile.sourceRubricFileId || '',
      JSONDefinition: JSON.stringify(profile), SchemaVersion: SCHEMA_VERSION,
      Active: false, CreatedAt: nowIso, CreatedBy: user || ''
    };
    var row = headers.map(function (h) { return Object.prototype.hasOwnProperty.call(values, h) ? values[h] : ''; });
    var newRowNumber = sheet.getLastRow() + 1;
    sheet.getRange(newRowNumber, 1, 1, headers.length).setValues([row]);
    /* Append first, then deactivate previous versions, then activate new.
       Old JSON is never mutated or discarded (even on legacy 4-col sheets). */
    for (var i = 0; i < entries.length; i++) {
      if (String(entries[i].row.TaskName) === String(profile.taskName)) {
        sheet.getRange(entries[i].rowNumber, column('Active')).setValue(false);
      }
    }
    sheet.getRange(newRowNumber, column('Active')).setValue(true);
    return { success: true, profile: profile, summary: profileSummary(profile) };
  }

  function readRubricRows(rubricFileId) {
    var ss = SpreadsheetApp.openById(rubricFileId);
    var sheet = ss.getSheetByName('Rubric');
    if (!sheet || sheet.getLastRow() < 2) throw new Error('The shared rubric has no Rubric rows.');
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function (h) { return String(h).trim(); });
    ['Criterion', 'MaxMarks', 'Band', 'Description'].forEach(function (h) {
      if (headers.indexOf(h) === -1) throw new Error('The shared rubric is missing the ' + h + ' column.');
    });
    var rows = [];
    for (var r = 1; r < data.length; r++) {
      var row = {};
      for (var c = 0; c < headers.length; c++) row[headers[c]] = data[r][c];
      if (!String(row['Criterion'] || '').trim() && !String(row['Description'] || '').trim()) continue;
      rows.push({
        criterion: String(row['Criterion'] || ''),
        part: String(row['Part'] || ''),
        section: String(row['Section'] || ''),
        maxMarks: row['MaxMarks'],
        outcome: String(row['Outcome'] || ''),
        band: String(row['Band'] || ''),
        description: String(row['Description'] || '')
      });
    }
    return rows;
  }

  function buildProfileFromRubricFile(rubricFileId, meta) {
    return buildProfileFromRubricRows(readRubricRows(rubricFileId), meta);
  }

  function buildProfileFromWorkbookCriteriaConfig(ss, meta) {
    var sheet = ss.getSheetByName('CriteriaConfig');
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function (h) { return String(h).trim(); });
    var rows = [];
    for (var r = 1; r < data.length; r++) {
      var row = {};
      for (var c = 0; c < headers.length; c++) row[headers[c]] = data[r][c];
      rows.push(row);
    }
    if (!rows.length) return null;
    return profileFromCriteriaConfigRows(rows, meta);
  }

  /* ------------------------------------------------------------------ *
   * Settings & Foundations: Thresholds, Outcomes, Deductions, Templates
   * ------------------------------------------------------------------ */

  function validateThresholds(bands) {
    var errors = [];
    if (!Array.isArray(bands) || bands.length < 2) {
      return { ok: false, errors: ['Grade scale must contain at least 2 bands (e.g. A–E).'] };
    }
    var sorted = bands.slice().sort(function (a, b) { return toNumber(b.min) - toNumber(a.min); });
    if (Math.abs(toNumber(sorted[0].max) - 100) > 0.05) {
      errors.push('Highest band max must be 100% (found ' + sorted[0].max + '%).');
    }
    if (Math.abs(toNumber(sorted[sorted.length - 1].min) - 0) > 0.05) {
      errors.push('Lowest band min must be 0% (found ' + sorted[sorted.length - 1].min + '%).');
    }
    for (var i = 0; i < sorted.length; i++) {
      var b = sorted[i];
      var letter = String(b.letter || '').trim().toUpperCase();
      if (!letter) errors.push('Band at index ' + (i + 1) + ' is missing a letter identifier.');
      var min = toNumber(b.min);
      var max = toNumber(b.max);
      if (min >= max) {
        errors.push('Band ' + letter + ' has invalid range: min (' + min + ') must be less than max (' + max + ').');
      }
      if (min < 0 || max > 100) {
        errors.push('Band ' + letter + ' range must be within 0% to 100%.');
      }
      if (i < sorted.length - 1) {
        var next = sorted[i + 1];
        var nextMax = toNumber(next.max);
        var gap = min - nextMax;
        if (min < nextMax) {
          errors.push('Band ' + letter + ' overlaps with band ' + next.letter + ' (' + min + ' < ' + nextMax + ').');
        } else if (gap > 0.05) {
          errors.push('Gap detected between band ' + letter + ' and ' + next.letter + ' (' + min + ' vs ' + nextMax + ').');
        }
      }
    }
    return { ok: errors.length === 0, errors: errors };
  }

  function previewThresholdRecalculation(oldBands, newBands, sampleScores) {
    var oldScale = normalizeGradeScale(null, oldBands).bands;
    var newScale = normalizeGradeScale(null, newBands).bands;
    var scores = Array.isArray(sampleScores) ? sampleScores : [];
    var changes = [];
    var oldDist = {};
    var newDist = {};
    for (var i = 0; i < scores.length; i++) {
      var item = scores[i] || {};
      var percent = typeof item === 'number' ? item : toNumber(item.percent || (item.points && item.maxMarks ? (item.points / item.maxMarks) * 100 : 0));
      var oldGrade = overallLetter(percent, oldScale);
      var newGrade = overallLetter(percent, newScale);
      oldDist[oldGrade] = (oldDist[oldGrade] || 0) + 1;
      newDist[newGrade] = (newDist[newGrade] || 0) + 1;
      if (oldGrade !== newGrade) {
        changes.push({
          studentName: item.studentName || item.StudentName || ('Student ' + (i + 1)),
          percent: percent,
          oldGrade: oldGrade,
          newGrade: newGrade
        });
      }
    }
    return {
      totalEvaluated: scores.length,
      changedCount: changes.length,
      changes: changes,
      oldDistribution: oldDist,
      newDistribution: newDist
    };
  }

  function calculateDeduction(rawScore, maxMarks, rule, submissionMeta) {
    var raw = toNumber(rawScore);
    var max = toNumber(maxMarks) || 100;
    var r = rule || { type: 'none' };
    var meta = submissionMeta || {};

    if (!r.type || r.type === 'none') {
      return {
        rawScore: raw,
        deductionPoints: 0,
        finalScore: raw,
        percent: max > 0 ? (raw / max) * 100 : 0,
        lateDays: 0,
        penaltyApplied: false,
        reason: 'No deduction applied.'
      };
    }

    var turnedIn = meta.turnedInTime ? new Date(meta.turnedInTime) : null;
    var due = meta.dueDate ? new Date(meta.dueDate) : null;
    var extension = meta.extensionDate ? new Date(meta.extensionDate) : null;
    var effectiveDue = extension || due;

    var lateDays = 0;
    if (meta.lateDays != null && isFinite(toNumber(meta.lateDays))) {
      lateDays = Math.max(0, toNumber(meta.lateDays));
    } else if (turnedIn && effectiveDue && !isNaN(turnedIn.getTime()) && !isNaN(effectiveDue.getTime())) {
      var graceMs = (toNumber(r.graceHours) || 0) * 3600000;
      var diffMs = turnedIn.getTime() - (effectiveDue.getTime() + graceMs);
      if (diffMs > 0) {
        if (r.excludeWeekends) {
          var cur = new Date(effectiveDue.getTime() + graceMs);
          while (cur < turnedIn) {
            var day = cur.getDay();
            if (day !== 0 && day !== 6) lateDays++;
            cur.setDate(cur.getDate() + 1);
          }
        } else {
          lateDays = Math.ceil(diffMs / (24 * 3600000));
        }
      }
    } else if (meta.isLate || meta.late) {
      lateDays = 1;
    }

    var deduction = 0;
    var reason = '';
    var rate = toNumber(r.rate);
    var maxDeduction = (toNumber(r.maxDeductionPercent) || 100) / 100 * max;
    var minFloor = toNumber(r.minFloor) || 0;

    if (r.type === 'percent_per_day') {
      if (lateDays > 0) {
        var base = r.base === 'raw' ? raw : max;
        deduction = (rate / 100) * base * lateDays;
        reason = lateDays + ' day(s) late @ ' + rate + '%/day (' + deduction.toFixed(2) + ' marks deduction)';
      } else {
        reason = 'Submission on time. No late penalty.';
      }
    } else if (r.type === 'fixed_marks') {
      deduction = rate;
      reason = 'Fixed deduction: -' + rate + ' marks (' + (r.reason || 'Penalty') + ')';
    } else if (r.type === 'percent_of_achieved') {
      deduction = (rate / 100) * raw;
      reason = rate + '% deduction of achieved mark (-' + deduction.toFixed(2) + ' marks)';
    }

    deduction = Math.min(deduction, maxDeduction);
    deduction = Math.min(deduction, raw);
    var finalScore = Math.max(minFloor, raw - deduction);
    var percent = max > 0 ? (finalScore / max) * 100 : 0;

    return {
      rawScore: raw,
      deductionPoints: deduction,
      finalScore: finalScore,
      percent: percent,
      lateDays: lateDays,
      penaltyApplied: deduction > 0,
      reason: reason
    };
  }

  function renderFeedbackTemplate(template, data) {
    var t = String(template || '').trim();
    if (!t) {
      return {
        text: (data && data.whatWentWell ? data.whatWentWell + '\n' + (data.areasForImprovement || '') : ''),
        warnings: []
      };
    }
    var d = data || {};
    var warnings = [];
    var rendered = t.replace(/\{\{([a-zA-Z0-9_]+)\}\}/g, function (match, token) {
      var key = token.toLowerCase();
      switch (key) {
        case 'teacher': return d.teacher || d.teacherName || '';
        case 'student': return d.student || d.studentName || '';
        case 'class': return d.class || d.className || '';
        case 'task': return d.task || d.taskName || '';
        case 'grade': return d.grade || '';
        case 'raw_score': return d.rawScore != null ? Number(d.rawScore).toFixed(1) : (d.score != null ? Number(d.score).toFixed(1) : '');
        case 'final_score': return d.finalScore != null ? Number(d.finalScore).toFixed(1) : (d.points != null ? Number(d.points).toFixed(1) : '');
        case 'score': return d.finalScore != null ? Number(d.finalScore).toFixed(1) : (d.points != null ? Number(d.points).toFixed(1) : '');
        case 'max_marks': return d.maxMarks != null ? String(d.maxMarks) : '';
        case 'percent': return d.percent != null ? Number(d.percent).toFixed(1) + '%' : '';
        case 'what_went_well': return d.whatWentWell || '';
        case 'areas_for_improvement': return d.areasForImprovement || '';
        case 'goals': return d.goalsForNextAssessment || d.goals || '';
        case 'date': return d.date || new Date().toISOString().slice(0, 10);
        case 'assessment_type': return d.assessmentType || 'Summative';
        case 'outcomes': return Array.isArray(d.outcomes) ? d.outcomes.join(', ') : (d.outcomes || '');
        default:
          warnings.push('Unknown template variable: {{' + token + '}}');
          return '{{' + token + '}}';
      }
    });
    return { text: rendered, warnings: warnings };
  }

  function validateOutcomeMappings(criteria, registeredOutcomes) {
    var list = Array.isArray(criteria) ? criteria : [];
    var registered = Array.isArray(registeredOutcomes) ? registeredOutcomes : [];
    var regMap = {};
    for (var r = 0; r < registered.length; r++) {
      var code = typeof registered[r] === 'string' ? registered[r] : registered[r].code;
      if (code) regMap[code.trim().toUpperCase()] = registered[r];
    }
    var mapped = 0;
    var unmapped = [];
    var coverage = {};
    for (var i = 0; i < list.length; i++) {
      var c = list[i];
      var outCode = String(c.outcome || '').trim().toUpperCase();
      if (outCode) {
        mapped++;
        coverage[outCode] = (coverage[outCode] || 0) + 1;
      } else {
        unmapped.push(c.criterionId || c.title || ('Criterion ' + (i + 1)));
      }
    }
    return {
      valid: unmapped.length === 0,
      totalCriteria: list.length,
      mappedCriteria: mapped,
      unmappedCriteria: unmapped,
      coverage: coverage
    };
  }

  function calculateOutcomeCoverage(profile, assessments) {
    var critMap = {};
    for (var i = 0; i < (profile && profile.criteria || []).length; i++) {
      var c = profile.criteria[i];
      critMap[c.criterionId] = c;
    }
    var outcomeResults = {};
    var assessList = Array.isArray(assessments) ? assessments : [];
    for (var a = 0; a < assessList.length; a++) {
      var item = assessList[a] || {};
      var cid = item.CriterionID || item.criterionId;
      var crit = critMap[cid];
      if (!crit || !crit.outcome) continue;
      var code = crit.outcome.trim().toUpperCase();
      if (!outcomeResults[code]) {
        outcomeResults[code] = {
          outcome: code,
          criteriaCount: 0,
          totalMarksAvailable: 0,
          totalMarksEarned: 0,
          evidenceCount: 0
        };
      }
      outcomeResults[code].criteriaCount++;
      outcomeResults[code].totalMarksAvailable += toNumber(crit.maxMarks);
      outcomeResults[code].totalMarksEarned += toNumber(item.DerivedPoints || item.FinalApprovedPoints || item.points || 0);
      if (item.teacherWrittenNote || item.TeacherWrittenNote || item.AIEvidenceJSON || item.aiEvidence) {
        outcomeResults[code].evidenceCount++;
      }
    }
    return outcomeResults;
  }

  function generateFormativeRubricDraft(assignmentData, existingRubrics, options) {
    var data = assignmentData || {};
    var title = String(data.title || 'Formative Task').trim();
    var desc = String(data.description || '').trim();
    var maxMarks = toNumber(data.maxMarks) || 20;
    var course = String(data.courseCode || '7TECHI').trim();
    var outcomes = Array.isArray(data.outcomes) && data.outcomes.length ? data.outcomes : ['DT5-1', 'DT5-2', 'DT5-3'];

    var c1Marks = Math.max(1, Math.round(maxMarks * 0.3));
    var c2Marks = Math.max(1, Math.round(maxMarks * 0.4));
    var c3Marks = Math.max(1, maxMarks - c1Marks - c2Marks);

    var rows = [
      { criterion: 'Research and Ideation', part: 'Investigation', section: 'Formative', maxMarks: c1Marks, outcome: outcomes[0] || 'DT5-1', band: 'A', description: 'Extensive research exploring multiple innovative concepts with thorough documentation.' },
      { criterion: 'Research and Ideation', part: 'Investigation', section: 'Formative', maxMarks: c1Marks, outcome: outcomes[0] || 'DT5-1', band: 'B', description: 'Thorough research into design possibilities with clear annotations.' },
      { criterion: 'Research and Ideation', part: 'Investigation', section: 'Formative', maxMarks: c1Marks, outcome: outcomes[0] || 'DT5-1', band: 'C', description: 'Sound research into chosen concepts with basic documentation.' },
      { criterion: 'Research and Ideation', part: 'Investigation', section: 'Formative', maxMarks: c1Marks, outcome: outcomes[0] || 'DT5-1', band: 'D', description: 'Basic research with limited exploration of alternatives.' },
      { criterion: 'Research and Ideation', part: 'Investigation', section: 'Formative', maxMarks: c1Marks, outcome: outcomes[0] || 'DT5-1', band: 'E', description: 'Elementary research with minimal or missing documentation.' },

      { criterion: 'Development and Prototyping', part: 'Application', section: 'Formative', maxMarks: c2Marks, outcome: outcomes[1] || 'DT5-2', band: 'A', description: 'Exceptional prototype development demonstrating advanced technical skill and material testing.' },
      { criterion: 'Development and Prototyping', part: 'Application', section: 'Formative', maxMarks: c2Marks, outcome: outcomes[1] || 'DT5-2', band: 'B', description: 'Detailed prototype development with clear testing of technical parameters.' },
      { criterion: 'Development and Prototyping', part: 'Application', section: 'Formative', maxMarks: c2Marks, outcome: outcomes[1] || 'DT5-2', band: 'C', description: 'Sound development resulting in a functional prototype meeting core requirements.' },
      { criterion: 'Development and Prototyping', part: 'Application', section: 'Formative', maxMarks: c2Marks, outcome: outcomes[1] || 'DT5-2', band: 'D', description: 'Basic prototype with partial functionality and limited testing.' },
      { criterion: 'Development and Prototyping', part: 'Application', section: 'Formative', maxMarks: c2Marks, outcome: outcomes[1] || 'DT5-2', band: 'E', description: 'Incomplete prototype showing minimal development.' },

      { criterion: 'Evaluation and Reflection', part: 'Evaluation', section: 'Formative', maxMarks: c3Marks, outcome: outcomes[2] || 'DT5-3', band: 'A', description: 'Critical reflection evaluating performance against design criteria with clear next steps.' },
      { criterion: 'Evaluation and Reflection', part: 'Evaluation', section: 'Formative', maxMarks: c3Marks, outcome: outcomes[2] || 'DT5-3', band: 'B', description: 'Thorough evaluation identifying strengths and actionable improvements.' },
      { criterion: 'Evaluation and Reflection', part: 'Evaluation', section: 'Formative', maxMarks: c3Marks, outcome: outcomes[2] || 'DT5-3', band: 'C', description: 'Sound evaluation discussing outcome success and basic improvements.' },
      { criterion: 'Evaluation and Reflection', part: 'Evaluation', section: 'Formative', maxMarks: c3Marks, outcome: outcomes[2] || 'DT5-3', band: 'D', description: 'Basic reflection with superficial comments on outcome.' },
      { criterion: 'Evaluation and Reflection', part: 'Evaluation', section: 'Formative', maxMarks: c3Marks, outcome: outcomes[2] || 'DT5-3', band: 'E', description: 'Minimal or missing reflection on the task.' }
    ];

    var profile = buildProfileFromRubricRows(rows, {
      taskName: title,
      name: title + ' (Formative Rubric Draft)',
      source: 'formativeGeneratedDraft'
    });
    profile.reviewed = false;
    profile.assessmentType = 'Formative';
    profile.rationale = 'Generated formative rubric draft aligning with course ' + course + ' and syllabus outcomes ' + outcomes.join(', ') + '.';
    return profile;
  }

  function saveGradeScale(ss, gradeScale, user) {
    var scale = normalizeGradeScale(gradeScale && gradeScale.weights, gradeScale && gradeScale.bands);
    var validation = validateThresholds(scale.bands);
    if (!validation.ok) {
      return { success: false, message: 'Invalid thresholds: ' + validation.errors.join(' ') };
    }
    var sheet = ss.getSheetByName('Setup');
    if (!sheet) {
      sheet = ss.insertSheet('Setup');
    }
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var data = [
        ['GRADE SCALE', 'WEIGHT'],
        ['A', scale.weights.A != null ? scale.weights.A : 1.0],
        ['B', scale.weights.B != null ? scale.weights.B : 0.875],
        ['C', scale.weights.C != null ? scale.weights.C : 0.70],
        ['D', scale.weights.D != null ? scale.weights.D : 0.575],
        ['E', scale.weights.E != null ? scale.weights.E : 0.25],
        ['', ''],
        ['OVERALL GRADE BAND', 'MIN', 'MAX']
      ];
      for (var i = 0; i < scale.bands.length; i++) {
        var b = scale.bands[i];
        data.push([b.letter, b.min, b.max]);
      }
      data.push(['', '', '']);
      data.push(['LAST_UPDATED', new Date().toISOString(), user || '']);
      sheet.clearContents();
      sheet.getRange(1, 1, data.length, 3).setValues(data);
      return { success: true, gradeScale: scale, message: 'Grade scale thresholds and weights saved successfully.' };
    } finally {
      lock.releaseLock();
    }
  }

  return {
    SCHEMA_VERSION: SCHEMA_VERSION,
    CORE_BANDS: CORE_BANDS,
    DEFAULT_WEIGHTS: DEFAULT_WEIGHTS,
    DEFAULT_BANDS: DEFAULT_BANDS,
    PROFILES_HEADERS: PROFILES_HEADERS,
    escapeHtml: escapeHtml,
    normalizeGradeScale: normalizeGradeScale,
    overallLetter: overallLetter,
    bandWeightsFor: bandWeightsFor,
    orderBandKeys: orderBandKeys,
    buildProfileFromRubricRows: buildProfileFromRubricRows,
    profileFromLegacyDefinitions: profileFromLegacyDefinitions,
    profileFromCriteriaConfigRows: profileFromCriteriaConfigRows,
    validateProfile: validateProfile,
    profileSummary: profileSummary,
    scoreCriterion: scoreCriterion,
    scoreProfile: scoreProfile,
    serializeForAi: serializeForAi,
    knownObservableIds: knownObservableIds,
    renderReportHtml: renderReportHtml,
    getGradeScale: getGradeScale,
    ensureProfilesSheet: ensureProfilesSheet,
    readProfiles: readProfiles,
    getActiveProfile: getActiveProfile,
    getProfileById: getProfileById,
    saveProfile: saveProfile,
    readRubricRows: readRubricRows,
    buildProfileFromRubricFile: buildProfileFromRubricFile,
    buildProfileFromWorkbookCriteriaConfig: buildProfileFromWorkbookCriteriaConfig,
    validateThresholds: validateThresholds,
    previewThresholdRecalculation: previewThresholdRecalculation,
    calculateDeduction: calculateDeduction,
    renderFeedbackTemplate: renderFeedbackTemplate,
    validateOutcomeMappings: validateOutcomeMappings,
    calculateOutcomeCoverage: calculateOutcomeCoverage,
    generateFormativeRubricDraft: generateFormativeRubricDraft,
    saveGradeScale: saveGradeScale
  };
})();
