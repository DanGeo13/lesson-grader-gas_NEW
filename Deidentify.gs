/* ============================================================================
 * Deidentify.gs — best-effort removal of student identifiers from TEXT before
 * it leaves the school's Google account for an external evaluator (JEV) or is
 * kept as a refinement example. (SC43, SC45; policy in docs/PRIVACY.md)
 *
 * Removes: the student's full, preferred, first and family names (and any
 * other roster names supplied), email addresses, Classroom/Google user IDs,
 * long digit runs (student numbers), phone numbers, URLs and long opaque IDs.
 * Does NOT and cannot de-identify images, handwriting, PDFs or names that are
 * misspelt or unknown to the roster. It is a filter, not a guarantee.
 *
 * Load-time pure; no Apps Script services.
 * ============================================================================ */
var Deidentify = (function () {
  var STOP = { the: 1, and: 1, for: 1, with: 1, from: 1, this: 1, that: 1, year: 1, task: 1, class: 1, mr: 1, ms: 1, mrs: 1, miss: 1 };

  function escapeRe(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  /** Name variants worth removing for one person: full name, each part of 3+ letters. */
  function nameTerms(names) {
    var terms = [];
    (names || []).forEach(function (n) {
      var full = String(n || '').trim().replace(/\s+/g, ' ');
      if (!full) return;
      if (full.split(' ').length > 1) terms.push(full);
      full.split(/[\s'’\-]+/).forEach(function (part) {
        var p = part.replace(/[^A-Za-zÀ-ÿ]/g, '');
        if (p.length >= 3 && !STOP[p.toLowerCase()]) terms.push(p);
      });
    });
    var seen = {};
    return terms.filter(function (t) { var k = t.toLowerCase(); if (seen[k]) return false; seen[k] = true; return true; })
      .sort(function (a, b) { return b.length - a.length; }); // longest first: "Mia Taylor" before "Mia"
  }

  /* Non-name identifiers: links, emails, known IDs, opaque IDs, phones, student numbers. */
  function scrub(text, identity) {
    var id = identity || {};
    var out = String(text == null ? '' : text);
    var counts = { name: 0, email: 0, id: 0, phone: 0, link: 0 };
    function rep(re, token, key) {
      out = out.replace(re, function () { counts[key]++; return token; });
    }
    rep(/\bhttps?:\/\/[^\s)<>"']+/gi, '[LINK]', 'link');
    rep(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g, '[EMAIL]', 'email');
    (id.ids || []).filter(Boolean).forEach(function (v) {
      if (String(v).length >= 4) rep(new RegExp(escapeRe(v), 'g'), '[ID]', 'id');
    });
    rep(/\b[A-Za-z0-9_\-]{25,}\b/g, '[ID]', 'id');
    rep(/(\+?61[\s\-]?|\b0)4\d{2}[\s\-]?\d{3}[\s\-]?\d{3}\b/g, '[PHONE]', 'phone');
    rep(/(\+?61[\s\-]?|\b0)[2378][\s\-]?\d{4}[\s\-]?\d{4}\b/g, '[PHONE]', 'phone');
    rep(/\b\d{6,}\b/g, '[ID]', 'id');
    return { text: out, counts: counts };
  }

  function scrubNames(text, terms) {
    var count = 0;
    var out = text;
    terms.forEach(function (term) {
      out = out.replace(new RegExp('(^|[^A-Za-zÀ-ÿ])' + escapeRe(term) + '(?=$|[^A-Za-zÀ-ÿ])', 'gi'), function (m, pre) { count++; return pre + '[STUDENT]'; });
    });
    return { text: out, count: count };
  }

  /**
   * @param {string} text
   * @param {{names?:string[], ids?:string[], rosterNames?:string[]}} identity
   * @return {{text:string, counts:Object}}
   */
  function apply(text, identity) {
    var id = identity || {};
    var base = scrub(text, { ids: id.ids });
    var named = scrubNames(base.text, nameTerms((id.names || []).concat(id.rosterNames || [])));
    base.counts.name = named.count;
    return { text: named.text, counts: base.counts };
  }

  /** Throws if obvious identifiers survive; used as a last check before sending. */
  function assertClean(text, identity) {
    var id = identity || {};
    var leaks = [];
    if (/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/.test(text)) leaks.push('email');
    nameTerms(id.names || []).forEach(function (t) {
      if (new RegExp('(^|[^A-Za-zÀ-ÿ])' + escapeRe(t) + '(?=$|[^A-Za-zÀ-ÿ])', 'i').test(text)) leaks.push('name');
    });
    (id.ids || []).forEach(function (v) { if (v && String(v).length >= 4 && text.indexOf(String(v)) !== -1) leaks.push('id'); });
    if (leaks.length) throw new Error('De-identification check failed (' + leaks.join(', ') + '); nothing was sent.');
  }

  return { apply: apply, assertClean: assertClean, nameTerms: nameTerms };
})();
