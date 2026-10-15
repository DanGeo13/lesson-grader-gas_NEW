/* ============================================================================
 * EvidenceService.gs — reads a submission's Drive files for AI drafting,
 * records which files were readable, and keeps an append-only history of the
 * Markdown/text used for each assessment. (SC21, SC23, SC44)
 *
 * File statuses written to SubmissionFiles.AIReviewStatus:
 *   Reviewed     content was supplied to the AI provider
 *   Unsupported  file type cannot be read by the configured provider
 *   Unreadable   Drive refused access or conversion failed
 *   TooLarge     over the per-file or per-run size limit
 *   Skipped      over the per-run file limit
 * Evidence text is stored in the EvidenceVersions sheet keyed by content hash,
 * so every distinct version used for assessment is retained.
 * ============================================================================ */
var EvidenceService = (function () {
  var SHEET = 'EvidenceVersions';
  var HEADERS = ['VersionID', 'SubmissionRecordID', 'TaskName', 'DriveFileID', 'FileName', 'MimeType', 'ContentHash',
    'Markdown', 'Truncated', 'ExtractedAt', 'ExtractedBy', 'UsedForRun'];
  var CELL_LIMIT = 45000;           // Sheets cell limit is 50,000 characters
  var MAX_FILES = 5;
  var MAX_FILE_BYTES = 15 * 1024 * 1024;
  var MAX_RUN_BASE64 = 18 * 1024 * 1024;
  var TEXT_MIMES = { 'text/plain': true, 'text/markdown': true, 'text/csv': true };
  var GOOGLE_DOC = 'application/vnd.google-apps.document';
  var GOOGLE_SLIDES = 'application/vnd.google-apps.presentation';
  var IMAGE_RE = /^image\/(jpeg|jpg|png|webp|gif|heic)$/;

  function text(v) { return v == null ? '' : String(v); }

  /** Classifies how a file can reach each provider. Pure. */
  function classify(mime) {
    var m = text(mime).toLowerCase();
    if (m === GOOGLE_DOC) return { kind: 'doc', text: true, inline: true };
    if (m === GOOGLE_SLIDES) return { kind: 'slides', text: true, inline: true };
    if (TEXT_MIMES[m]) return { kind: 'text', text: true, inline: false };
    if (m === 'application/pdf') return { kind: 'pdf', text: false, inline: true };
    if (IMAGE_RE.test(m)) return { kind: 'image', text: false, inline: true };
    if (m === 'application/vnd.google-apps.drawing') return { kind: 'drawing', text: false, inline: true };
    return { kind: 'other', text: false, inline: false };
  }

  function contentHash(value) {
    try {
      var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
      return 'sha256:' + bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
    } catch (e) {
      return 'djb2:' + AssessmentFlags.hash(value);
    }
  }

  /* Google Docs → Markdown via the Drive export endpoint (keeps headings and
     lists); falls back to plain text. Slides export as plain text. */
  function exportText(fileId, kind) {
    if (kind === 'doc') {
      try {
        var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '/export?mimeType=text%2Fmarkdown', {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
        });
        if (res.getResponseCode() === 200) return res.getContentText();
      } catch (e) { /* fall through to plain text */ }
    }
    return DriveApp.getFileById(fileId).getAs('text/plain').getDataAsString();
  }

  /**
   * Reads up to MAX_FILES files for one submission.
   * @param {Array} fileRows SubmissionFiles rows (already filtered to the submission)
   * @param {{wantInline:boolean, wantText:boolean}} need what the provider can use
   * @return {{files:Array}} each {fileRecordId, rowNumber, fileName, driveFileId, mimeType, status, limitations, markdown, inlineData}
   */
  function collect(fileRows, need) {
    var out = [];
    var used = 0, base64Total = 0;
    (fileRows || []).forEach(function (r) {
      var item = { fileRecordId: text(r.FileRecordID), rowNumber: r._row, fileName: text(r.FileName) || 'Untitled file',
        driveFileId: text(r.DriveFileID), mimeType: text(r.MimeType), status: 'Unsupported', limitations: '', markdown: null, inlineData: null };
      out.push(item);
      if (!item.driveFileId) { item.status = 'Unreadable'; item.limitations = 'No Drive file ID recorded.'; return; }
      if (used >= MAX_FILES) { item.status = 'Skipped'; item.limitations = 'Only the first ' + MAX_FILES + ' files are read per run.'; return; }
      try {
        var file = DriveApp.getFileById(item.driveFileId);
        if (!item.mimeType) item.mimeType = file.getMimeType();
        var cls = classify(item.mimeType);
        var canText = need.wantText && cls.text;
        var canInline = need.wantInline && cls.inline;
        if (!canText && !canInline) {
          item.status = 'Unsupported';
          item.limitations = cls.kind === 'other' ? 'File type is not supported for automatic reading.' :
            'This provider cannot read ' + cls.kind + ' files' + (cls.kind === 'image' || cls.kind === 'pdf' ? ' (enable a vision provider or use Gemini).' : '.');
          return;
        }
        if (file.getSize && file.getSize() > MAX_FILE_BYTES) { item.status = 'TooLarge'; item.limitations = 'Over 15 MB.'; return; }
        if (canText) {
          item.markdown = exportText(item.driveFileId, cls.kind);
        }
        if (canInline && !canText) {
          var blob = /google-apps/.test(item.mimeType) ? file.getAs('application/pdf') : file.getBlob();
          var data = Utilities.base64Encode(blob.getBytes());
          if (base64Total + data.length > MAX_RUN_BASE64) { item.status = 'TooLarge'; item.limitations = 'Run size limit reached.'; return; }
          base64Total += data.length;
          item.inlineData = { mimeType: /google-apps/.test(item.mimeType) ? 'application/pdf' : (blob.getContentType() || item.mimeType), data: data };
        }
        used++;
        item.status = 'Reviewed';
      } catch (e) {
        item.status = 'Unreadable';
        item.limitations = 'Drive could not open this file for the signed-in teacher.';
      }
    });
    return { files: out };
  }

  function ensureSheet(ss) {
    var sheet = ss.getSheetByName(SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  /** Writes per-file statuses back to SubmissionFiles (non-destructive status columns only). */
  function recordStatuses(ss, files) {
    files.forEach(function (f) {
      if (!f.rowNumber) return;
      WebGrading.patch(ss, 'SubmissionFiles', f.rowNumber, {
        AIReviewStatus: f.status, Limitations: f.limitations, MimeType: f.mimeType, EvidenceStatusUpdatedAt: new Date().toISOString()
      });
    });
  }

  /** Appends each new distinct evidence text; never overwrites earlier versions. */
  function recordVersions(ss, submissionId, taskName, files, runId, actor) {
    var withText = files.filter(function (f) { return f.markdown != null && f.status === 'Reviewed'; });
    if (!withText.length) return [];
    var sheet = ensureSheet(ss);
    var existing = WebGrading.records(ss, SHEET).filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId); });
    var stored = [];
    withText.forEach(function (f) {
      var hash = contentHash(f.markdown);
      f.contentHash = hash;
      if (existing.some(function (r) { return text(r.DriveFileID) === f.driveFileId && text(r.ContentHash) === hash; })) return;
      var md = text(f.markdown);
      WebGrading.append(ss, SHEET, {
        VersionID: Utilities.getUuid(), SubmissionRecordID: submissionId, TaskName: taskName, DriveFileID: f.driveFileId,
        FileName: f.fileName, MimeType: f.mimeType, ContentHash: hash, Markdown: md.slice(0, CELL_LIMIT),
        Truncated: md.length > CELL_LIMIT, ExtractedAt: new Date().toISOString(), ExtractedBy: actor, UsedForRun: runId || ''
      });
      stored.push(hash);
    });
    return stored;
  }

  function history(ss, submissionId) {
    if (!ss.getSheetByName(SHEET)) return [];
    return WebGrading.records(ss, SHEET).filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId); }).map(function (r) {
      return { versionId: text(r.VersionID), fileName: text(r.FileName), hash: text(r.ContentHash), extractedAt: text(r.ExtractedAt),
        truncated: r.Truncated === true || text(r.Truncated).toLowerCase() === 'true', length: text(r.Markdown).length, runId: text(r.UsedForRun) };
    });
  }

  function versionText(ss, submissionId, versionId) {
    var row = WebGrading.records(ss, SHEET).filter(function (r) { return text(r.SubmissionRecordID) === text(submissionId) && text(r.VersionID) === text(versionId); })[0];
    if (!row) throw new Error('Evidence version not found.');
    return { fileName: text(row.FileName), extractedAt: text(row.ExtractedAt), markdown: text(row.Markdown) };
  }

  return { HEADERS: HEADERS, classify: classify, collect: collect, recordStatuses: recordStatuses,
    recordVersions: recordVersions, history: history, versionText: versionText, contentHash: contentHash, MAX_FILES: MAX_FILES };
})();
