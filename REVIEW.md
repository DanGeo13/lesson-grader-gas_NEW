# Lesson Grader — Full Project Review

Review date: 15 October 2026. Reviewer: Claude Code (automated engineering review, no live Google access).

This file is the working record for the review. Phase 1 (discovery) was written **before any application code
changed**. Later phases append to the sections marked *Phase 2+*. Acceptance status per success criterion is in
[ACCEPTANCE.md](ACCEPTANCE.md); the forward plan is in [ROADMAP.md](ROADMAP.md).

> **Branch note.** The brief asked for branch `review/full-project-improvement`. This cloud session is bound to the
> designated branch `claude/bold-lovelace-mg2bwl` and is not permitted to push elsewhere, so all work is on that
> branch, created from `main` at `4938680`. Rename or re-push it as `review/full-project-improvement` if you prefer
> that name (commands in the final report).

---

## 1. Project overview (Phase 1)

### 1.1 Purpose inferred from code

A Google Apps Script assessment system for a NSW Technology and Applied Studies teacher. It has **two generations
of code in one Apps Script project**:

| Generation | Files | What it does |
|---|---|---|
| **Legacy Sheets Studio** (container-bound) | `Code.gs`, `AssessmentStudio.html`, `ClassBatchGrading.html`, `CourseworkPicker.html`, `GradingWorkbookSetup.html`, `ManageRubrics.html`, `VoiceDictation.html` | Year 9 Jewellery only. Hardcoded 12-criterion rubric (`Code.gs:165-298`), Sheets menu, modal dialogs, Gemini drafting, approval into the `9DAT1 Marking` sheet, explicit Classroom grade sync. Uses `SpreadsheetApp.getActiveSpreadsheet()` throughout, so it only works when the script is bound to that workbook. |
| **Web app "Teacher's Studio"** (standalone `doGet`) | `WebAppIntegration.gs`, `WebGrading.gs`, `RubricEngine.gs`, `WebAccess.gs`, `WebApp.html`, `site/index.html` | Task-agnostic rubric profiles, one workbook per class task in a Drive library, workbook-scoped server functions, draft → approve → lock → reassess, Gemini evidence proposals, plus several prototype "audit" screens. |

### 1.2 Relationship to the target specification

The web app is the intended product. Its **server layer** (`WebGrading.gs`, `RubricEngine.gs`) is a sound
foundation: explicit workbook IDs, optimistic revisions, append-only history, pinned rubric versions and
deterministic scoring. Its **browser layer** (`WebApp.html`) is largely a visual prototype: many primary buttons
have no handler (Run AI, Lock, Reassess, Report, Import, Activate rubric, Create workbook, Send messages,
Upload rubric), the Classroom view throws `ReferenceError: loadClassroom is not defined`, and several server
functions report success for actions they never perform. There is **no Classroom publishing, no automatic sync,
no unlock, no AI queue, no Qwen/JEV integration, no evidence Markdown history and no prompt testing** in the web
app. See §3 and ACCEPTANCE.md.

### 1.3 Stack

- Google Apps Script, V8 runtime, `Australia/Sydney` time zone (`appsscript.json`).
- Advanced service: Classroom v1. Implicit services: DriveApp, SpreadsheetApp, PropertiesService, LockService,
  UrlFetchApp, HtmlService, ScriptApp.
- No explicit `oauthScopes` in the manifest (Apps Script infers scopes; see F33).
- Vanilla HTML/CSS/JS client; Google Fonts loaded from `fonts.googleapis.com`.
- Node 22 tooling only for checks/tests: `jsdom` (dev dependency). No build step, no clasp config committed.

### 1.4 File map

| Path | Lines | Role |
|---|---|---|
| `Code.gs` | 3131 | Legacy modules: `GradeScaleService`, `Config` (hardcoded rubric), `Utils`, `Logging`, `Sheets`, `RosterService`, `RubricsLibrary` (Drive folder/index), `AssessmentService`, `DriveService`, `ClassroomService`, `GeminiService`, `HtmlTemplates`, menus, ~45 `api*` RPCs. **Shared by the web app**: `Config` sheet names, `DriveService`, `ClassroomService.listTeacherCourses/listCourseWork`, `GeminiService.callGeminiWithFallback`, `RubricsLibrary`, `apiSubmitGradingWorkbookSetup`, `apiCreateRubricFromDocument`, `deriveStageFromCourseCode`. |
| `RubricEngine.gs` | 1225 | Pure rubric profile builders, validation, scoring, report HTML, thresholds, deductions, feedback templates; workbook-scoped profile store (`RubricProfiles`) and grade scale (`Setup`). |
| `WebGrading.gs` | 1350 | Web grading service: prepare tables, task dashboard, detail, save draft, approve, lock, reassess, rubric preview/activate, AI draft, report PDF, manual student, Classroom import, bulk AI, cohort check, samples, registration prefill, audit dashboard, incomplete flags, message drafts, class table; `apiWeb*` RPCs. |
| `WebAppIntegration.gs` | 1381 | `LessonGraderWeb`: Drive library walk, workbook validation/open, list, metadata, rubrics, tasks, create, Classroom link, read-only overview/detail (unused); `LessonGraderLauncherApi` JSON/JSONP probe; the only `doGet`. |
| `WebAccess.gs` | 239 | Teacher allowlist (`LESSON_GRADER_ALLOWED_EMAILS`), Gemini key set/test, settings in Script Properties, threshold preview, grade-scale save, formative draft, AI provider "test". |
| `WebApp.html` | 1358 | Single-page web UI (styles + markup + script). |
| `site/index.html` | 1352 | Static GitHub Pages launcher (stores deployment URL in localStorage, probes `?lgapi=status`, embeds or opens the app). |
| Legacy dialogs | 101–481 each | Sheets modal UIs listed in §1.1. |
| `scripts/check.cjs` | 64 | Syntax/structure checks (parses every `.gs` and inline script, one `doGet`, manifest, no hardcoded deployment in launcher). |
| `tests/*.test.cjs` | 1399 | 53 Node tests using `vm` sandboxes and an in-memory Sheet/Workbook fake (`tests/helpers.cjs`) plus jsdom UI tests. |
| `docs/` | — | `AUDIT.md` (earlier audit of an older revision), `IMPLEMENTATION.md`, `LAUNCHER.md`, local mock preview server. |
| `.github/workflows/` | — | `verify.yml` (check + test on push/PR), `pages.yml` (publishes `site/`). |

Third-party/generated: `node_modules/` (jsdom tree, not committed), `package-lock.json` (generated). Not read in full.

### 1.5 Entry points, deployment model, services and APIs

- **`doGet(e)`** (`WebAppIntegration.gs:1375`): `?lgapi=status` → read-only probe (JSON or JSONP with a validated
  callback); otherwise serves `WebApp.html` via `createHtmlOutputFromFile` **without** `addMetaTag('viewport')`
  (F05) and without `setXFrameOptionsMode` (F34).
- **No `doPost`.** All app calls go through `google.script.run` to global `apiWeb*` functions, each calling
  `WebAccess.requireTeacher()` first.
- **Legacy entry points:** `onOpen` menu, modal dialog openers, `checkAndAutoImportDueAssignments` (time trigger
  handler) and `installAutoImportTrigger`.
- **Intended deployment** (docs): Web app, *Execute as: user accessing*, *Who has access: domain*. With that model
  every Drive/Sheets/Classroom call runs as the signed-in teacher and the library root is that teacher's own
  `My Drive/Graded Assessments`.
- **External API:** Gemini `generativelanguage.googleapis.com/v1beta/...:generateContent`, key in
  `x-goog-api-key` header (`Code.gs:2016-2030`). No Qwen or JEV client exists (only a Script Property name and a
  settings field).

### 1.6 Data flow (current)

```
Teacher browser ──google.script.run──▶ apiWeb* (WebAccess.requireTeacher)
                                          │
              LessonGraderWeb.open(id) ───┤  walks My Drive/Graded Assessments/Stage N/<course>/<category>/*.gsheet
                                          ▼  (full Drive traversal on every call — F25)
                                  Grading workbook (one per class task)
             Submissions · SubmissionFiles · CriterionAssessments · AIAssessments · AssessmentHistory
             ApprovedGrades · RubricProfiles · Setup · ClassroomConfig · AuditLog · ClassTracking · ClassroomMessages
                                          │
     Classroom (list coursework, list studentSubmissions, user profiles) ◀── importTask (manual only)
     Drive (read attachments → base64)  ──▶ Gemini (inline files + rubric) ──▶ validated ticks ──▶ draft rows
```

### 1.7 Persistence structures and contracts (must be preserved)

Workbook sheets created/extended by `WebGrading.prepare` (`WebGrading.gs:6-15`) — all writes are by header name
(`append`/`patch`), and migrations only **append** missing columns:

| Sheet | Key columns | Notes |
|---|---|---|
| `Submissions` | `SubmissionRecordID`, Classroom IDs, `StudentUserID`, `Task`, `SubmissionVersion`, `Status`, `ParentSubmissionRecordID`, `CurrentOfficial`, approval/lock/sync fields, scores | One row per evidence version. Status vocabulary: `New`, `InReview`, `Approved`, `Locked`, `Superseded` (+ legacy `AIAssessed`, `TeacherReviewed`, `SyncFailed`). |
| `SubmissionFiles` | `SubmissionRecordID`, `DriveFileID`, `FileName`, `MimeType`, `EligibleForAI`, `AIReviewStatus`, `Limitations` | `AIExtractedText`/`Limitations` never populated by the web app. |
| `CriterionAssessments` | `AssessmentID`, `SubmissionRecordID`, `CriterionID`, `CheckboxesJSON`, `AIEvidenceJSON`, `TeacherWrittenNote`, `Status`, `RubricProfileID` | Current draft per submission; overwritten in place on each save; history is in `AssessmentHistory`. |
| `AIAssessments` | `RunType` (`TeacherDraft`, `EvidenceProposal`, legacy `FeedbackGeneration`) | Feedback drafts are rows with `RunType=TeacherDraft`. |
| `AssessmentHistory` | `HistoryID` (also the **optimistic revision token**), `Action`, `PreviousStateJSON`, `NewStateJSON` | Append-only. |
| `ApprovedGrades` | `GradeID`, `SubmissionRecordID`, `RubricProfileID`, `Points`, `Percent`, `Grade`, `FeedbackJSON`, scores | Immutable approval snapshot, one per submission version. |
| `RubricProfiles` | `ProfileID`, `TaskName`, `JSONDefinition`, `Active` | Versioned rubric snapshots; assessments pin `RubricProfileID`. |
| `Setup` | free-form `GRADE SCALE` and `OVERALL GRADE BAND` tables | Read by `RubricEngine.getGradeScale` (`RubricEngine.gs:658`). |
| `ClassroomConfig` | `CourseID`, `CourseWorkID`, `TaskName`, `DueDate`, `AutoImported`, `Active` | Task ↔ assignment links. |
| `AuditLog`, `ClassTracking`, `ClassroomMessages` | — | Prototype audit/tracking tables. |

Drive library (`Code.gs:718-760`): `Graded Assessments/{Settings/Rubrics Index, Shared Rubrics/, Stage 4|5|6/<code>/<category>/}`.
Workbook name: `<CODE> [<YEAR>] - <Task> - Grading` (`Code.gs:908`).

Script Properties: `LESSON_GRADER_ALLOWED_EMAILS`, `GEMINI_API_KEY`, `LESSON_GRADER_SETTINGS` (JSON),
`LESSON_GRADER_WEB_CATEGORY`, `QWEN_API_KEY`.

### 1.8 Authentication and authorisation

- `WebAccess.requireTeacher()` (`WebAccess.gs:21`) fails closed unless the active user's email is in the
  allowlist property. Every public `apiWeb*` and legacy `api*` calls it first (covered by `tests/security.test.cjs`).
- Workbook authorisation = membership of the caller's own Drive library (`WebAppIntegration.gs:129-163`).
- No role separation: every allowlisted teacher can approve, lock, change settings and grade scales. No
  "authorised unlock" role exists.
- `status()` returns the full email to the authorised app page (acceptable: same origin, signed-in user);
  the cross-origin probe masks it.

### 1.9 Assumptions and unresolved requirements (Phase 1)

| # | Assumption / open question | Basis |
|---|---|---|
| A1 | The web app is the product; the legacy Sheets Studio is kept working for the existing Jewellery workbook but not extended. | Code comments `Code.gs:1-6`, `WebGrading.gs:1-4`, README. |
| A2 | "JEV" is an external text-evaluation service (named in the brief alongside Qwen). The repo only contains a "JEV Subject Code" settings field defaulting to `9JEV`, which looks like a **class code** (Year 9 Jewellery). I assume JEV is a separate, owner-hosted evaluator reachable over HTTPS and that `9JEV` is unrelated. **Owner to confirm.** | Brief; `WebAccess.gs:120`, `WebApp.html:562`. |
| A3 | Qwen 3-VL runs on an owner-controlled always-on machine exposing an OpenAI-compatible `/v1/chat/completions` endpoint (the usual vLLM/Ollama/LM Studio interface). | Brief; no code in repo. |
| A4 | The Classroom API cannot create private comments and only lets a project patch grades on coursework **created by the same Google Cloud project**. Publishing to teacher-created assignments will therefore fail with `ProjectPermissionDenied` unless the assignment was created by this script's project. Must be verified live. | Classroom API documentation (known platform limit); legacy `Code.gs:1942` relies on the same call. |
| A5 | Time zone for all due-date logic is `Australia/Sydney` (manifest). Classroom due dates are UTC. | `appsscript.json`; Classroom API. |

---

## 2. Baseline (Phase 1)

| Command | Result |
|---|---|
| `node --version` | v22.22.0 |
| `npm ci` | OK (jsdom 26) |
| `npm run check` | ✅ all 12 source files and the launcher parse; one `doGet`; manifest has Classroom |
| `npm test` (`node --test tests/*.test.cjs`) | ✅ **53 pass, 0 fail** (5 suites: launcher 22, rubric-engine 12, security 3, ui-smoke 4, web-grading 12) |
| Lint / formatter | None configured. |
| Build | None (Apps Script files are deployed as-is). |
| `clasp` | Not configured (`.clasp.json` git-ignored, not present). |

**Checks that could not run:** anything against a live Apps Script deployment, Google Classroom, Drive, Sheets,
Gemini, Qwen or JEV. No credentials or test resources are available to this session, and the brief forbids
touching production records. Visual rendering checks are limited to jsdom (no layout engine) unless a headless
browser run is noted later.

**Is it genuinely persistent?** Partly. Grading drafts, approvals, locks, reassessment versions, rubric profiles
and Classroom imports really write to the workbook. **Prototype behaviour found:**

- `sendClassroomMessages` writes a sheet row and returns *"Successfully sent N classroom notification messages"*;
  nothing is sent (`WebGrading.gs:1117-1135`).
- `saveRegistrationPrefill` returns *"Registration & Evaluation document entry updated successfully"*; no
  document is touched (`WebGrading.gs:1010-1037`).
- `saveWorkSamples` logs `Status: 'Saved'` and reports success even when the Drive copy failed or no folder was
  given (`WebGrading.gs:942-960`); the UI always passes `{}` destinations, so **no copy is ever made**.
- `testAiProvider` returns *"configuration verified successfully with mock probe"* for any non-Gemini provider
  (`WebAccess.gs:219`).
- `previewRegistrationPrefill` fills invented evaluation text ("Strong engagement with CAD/modelling…",
  `WebGrading.gs:997-999`); the UI save sends hardcoded outcomes "DT5-1, DT5-2, DT5-3" (`WebApp.html:1203`).
- Assessment Studio shows fake "evidence anchors" `Page 1 … Iterative Test 1` (`WebApp.html:912`).
- Settings thresholds/deductions are stored in Script Properties but **never used for grading** (grading reads the
  workbook `Setup` sheet), so the Settings screen misrepresents the active policy.

---

## 3. Findings (Phase 1)

Severity: **Critical** = wrong grades/data loss/security exposure/false success on a consequential action;
**High** = required workflow broken or unsafe; **Medium** = correctness/robustness gap; **Low** = quality.
Status is updated in later phases.

| ID | Area | Severity | File:Line | Issue | Proposed fix | Status |
|---|---|---|---|---|---|---|
| F01 | Security / credentials | Critical | `WebApp.html:1291`, `WebAccess.gs:160,178,145` | "Save All Settings" sends `geminiApiKey` inside the settings patch; `saveSettings` merges arbitrary keys into `LESSON_GRADER_SETTINGS` and `getSettings` returns them to the browser — the Gemini key can be stored in plain settings JSON and echoed back. | Whitelist settings keys; strip any `*key*`/secret fields; never return secrets; separate key endpoints only. Migration: scrub existing property on next save. | Fixed |
| F02 | Honesty / persistence | Critical | `WebGrading.gs:1117-1135` | Classroom messages reported as "sent" but only logged. | Record as drafts with status `Drafted — not sent`; message states nothing was sent and why (no Classroom messaging API). | Fixed |
| F03 | Honesty / persistence | Critical | `WebGrading.gs:1010-1037`, `WebApp.html:1198-1214` | Register "updated successfully" but no register is written; hardcoded field values. | Real commit to a Google Doc (validated ID) or an `EvaluationRegister` sheet; pre-fill from real task data; preview + confirm. | Fixed |
| F04 | Honesty / persistence | High | `WebGrading.gs:942-960`, `WebApp.html:1220` | Work samples report success when no copy was made. | Fail per tier unless a copy succeeded; require destination folders; return partial result honestly. | Fixed |
| F05 | Mobile | High | `WebAppIntegration.gs:1378-1380` | HtmlService ignores `<meta name="viewport">` in the file; without `addMetaTag` phones render the desktop width. | `.addMetaTag('viewport','width=device-width, initial-scale=1')`. | Fixed |
| F06 | Workflow (UI) | Critical | `WebApp.html` (ids `run-ai`, `lock-grade`, `new-reassessment`, `download-report`, `task-import`, `task-add-rubric`, `activate-rubric`, `editor-add-criterion`, `upload-rubric`, `create-form`, `btn-send-classroom-messages`, `retry-auth`, `refresh-workbooks`, `student-search`) | Primary buttons have no event handler; the create form submits and reloads the page; `navigate('classroom')` calls undefined `loadClassroom` (`WebApp.html:699`). A fresh workbook cannot get a rubric or Classroom submissions through the UI. | Rewrite the client around a central API wrapper and wire every control to a real server action. | Fixed |
| F07 | Approval correctness | Critical | `WebApp.html:1048-1064`, `WebGrading.gs:341-410` | Approve uses the **last saved** draft; unsaved on-screen edits are silently ignored and the UI then shows "Approved". | Client: block approve while dirty (save first, then approve with the new revision). Server already checks revision. | Fixed |
| F08 | Legacy data integrity | Critical | `Code.gs:578-584` | `Utils.buildRichTextFromMarkdown` does not exist. Legacy approval writes letters to `9DAT1 Marking` then throws before marking rows Approved → partially applied approval. | Implement the helper (plain rich text) and write the markbook after status validation. | Fixed |
| F09 | Grading rules | High | `RubricEngine.gs:37`, `WebAccess.gs:100`, `Code.gs:22,158` | Band weights A 1.0 / B 0.875 / C 0.70 / D 0.575 / E 0.25 conflict with confirmed reference points A 100% / C 50% / E 0% (SC19). | Do not silently change. Make the active rule set explicit, report conflict in health check, provide a target rule set that cannot be activated until B/D are confirmed. See §4. | Mitigated — blocked on policy |
| F10 | Grading rules | Medium | `RubricEngine.gs:98-104,484-486` | Letter boundaries compare raw floating percentages; e.g. 17/20 can be computed as 84.999999… and fall to B. | Normalise percentages to 9 d.p. before comparison (removes float noise, not a policy change); test every boundary. | Fixed |
| F11 | Grading UI drift | High | `WebApp.html:988-1018` | Client live score re-implements maths with hardcoded boundaries 85/75/65/50 and a default weight of 1.0 when a weight is missing; diverges from server for custom `Setup` scales. | Server returns the profile's scale; client mirrors `RubricEngine` exactly (shared tested function) and server remains authoritative on save. | Fixed |
| F12 | Approval rules | High | `WebGrading.gs:369-371` | A criterion with no ticks blocks approval forever, so genuine "no creditable evidence" cannot be approved without ticking something false. | Explicit per-criterion "No creditable evidence (0 marks)" teacher confirmation that resolves the blocking flag; recorded in history. | Fixed |
| F13 | Flags | High | `WebGrading.gs` | No distinction between blocking flags and advisory suggestions; AI evidence quality, unreadable files and stale feedback are not flagged. | Pure `Flags` module: blocking (no selection, unresolved AI insufficient-evidence, unreadable evidence not acknowledged) vs advisory (capped bands, teacher note not in feedback, feedback older than grade). | Fixed |
| F14 | Publishing | High | — | Web app has no Classroom publishing at all (SC39/40). | Separate `Publishing` module: preview selected approved students → typed confirmation → `studentSubmissions.patch` per student with isolation, log + history; never automatic. | Fixed |
| F15 | Locking | High | `WebGrading.gs:412-432` | No unlock workflow; locked work can only be reassessed by version (reassess allows Locked). | Locked blocks reassess/publish changes; explicit unlock with reason, restricted to `LESSON_GRADER_UNLOCK_EMAILS`, recorded in history. | Fixed |
| F16 | Sync | High | `Code.gs:1544-1635,1965` | The only automatic import uses `getActiveSpreadsheet()`; in the standalone web project that is `null`, so no web workbook is ever auto-synced (SC08). | Library-wide hourly sync over every workbook link: import due-dated assignments after due date, keep syncing hourly, idempotent, never touches grades. Trigger install/remove from Settings. | Fixed |
| F17 | Sync | Medium | `WebGrading.gs:697-698` | Change detection only compares attachment IDs; an edited Google Doc (same file ID) shows as "unchanged", and Classroom state/late flags are never refreshed. | Refresh non-grading metadata; if evidence changed after approval, flag `EvidenceChangedAt` + history without touching the approved result. | Fixed |
| F18 | Sync | Medium | `WebGrading.gs:750`, `WebGrading.gs:332-336` | Import and drafts call `WebAccess.requireTeacher()` for the actor, which fails inside a time trigger. | Pass an explicit actor; trigger context resolves the trigger owner against the allowlist. | Fixed |
| F19 | AI reliability | High | `WebGrading.gs:756-791` | Bulk AI runs synchronously in one request: 6-minute limit, no resume, no queue, no outage choice (SC48/49). | `AiQueue` sheet with dedupe key, time-budgeted processor, retry/attempt counts, stale-revision skip, readiness counts; manual continue always available. | Fixed |
| F20 | AI evidence | High | `WebGrading.gs:574-595` | AI ticks without evidence notes are accepted; criteria with no supported evidence are not flagged (SC22). | Accept a tick only with an evidence note naming a supplied file; otherwise record as unsupported and flag. | Fixed |
| F21 | AI evidence | Medium | `WebGrading.gs:535-548` | Skipped/unreadable files are returned to the client but not persisted or shown in the Studio (SC23). | Persist `AIReviewStatus`/`Limitations` per file; show readable/unreadable badges; unreadable files create an acknowledgement flag. | Fixed |
| F22 | Privacy | High | `WebGrading.gs:550-557`, `Code.gs:2185` | No de-identification policy; no JEV path; legacy feedback prompt sends the full student name. | `Deidentify` module (names, emails, IDs, phone numbers) used for JEV and refinement examples; document limits (binary files cannot be de-identified). | Fixed |
| F23 | Evidence history | Medium | — | Evidence text used for assessment is never stored (SC44). | `EvidenceVersions` append-only sheet keyed by content hash. | Fixed |
| F24 | Async/UI | High | `WebApp.html:875-883,759-786` | No stale-response guard: opening student B while A is loading can render A's data under B's context; saves then target A's id with B's on-screen ticks. | Sequence token per context; ignore responses for a superseded context; every write carries workbook/task/submission captured at render. | Fixed |
| F25 | Performance | Medium | `WebAppIntegration.gs:129-163` | Every call walks the whole Drive library to validate membership. | Per-user cache of verified library IDs (10 min), invalidated on create; still validates MIME/trash on each open. | Fixed |
| F26 | Classroom | Medium | `Code.gs:1433` | Teacher courses not paginated (50 cap). | Paginate. | Fixed |
| F27 | Data integrity | Medium | `WebGrading.gs:1179-1201` | Class tracking hardcodes `7TECHI`, `Stage 5`, `2026`, `MaxMarks: 100`; writes the submission ID into `StudentUserID`. | Derive from workbook/task data; store the real student ID; keep idempotent key. | Fixed |
| F28 | XSS | Medium | `AssessmentStudio.html:80,172-187` | Legacy studio injects teacher notes, explanations and error text into `innerHTML` unescaped (`</textarea>` breakout). | Escape all interpolated values. | Fixed |
| F29 | XSS | Low | `WebApp.html:907` | Evidence `href` uses HTML escaping only; a `javascript:` link would execute. | Allow only `https:` Drive/Docs links. | Fixed |
| F30 | Dictation | High | `WebApp.html:1066-1099`, `VoiceDictation.html:135-140` | Web dictation is single-shot, has no permission/no-speech error messages, does not mark the field dirty for prefill fields, and is absent from feedback fields; legacy dictation posts to `window.opener`, which is null for GAS dialogs, so transcripts are lost. | Shared dictation controller: continuous with interim preview, appends to any textarea, editable, explicit states (listening/denied/unsupported/no speech), never saves or approves. Legacy: relay through the parent via `google.script.run` is out of scope — document as legacy limitation. | Fixed (web); legacy dialog deferred |
| F31 | Settings integrity | Medium | `WebAccess.gs:153-180`, `WebApp.html:1266-1297` | Settings accept any shape; thresholds saved there do not change grading. | Validate known keys; grading policy lives only in the workbook `Setup` sheet and is shown read-only with an explicit "Save grade scale to this workbook" action. | Fixed |
| F32 | Feedback | High | — | No way to rebuild feedback from revised selections/observations; no stale-feedback detection; first-name usage uncontrolled (SC31–34). | Deterministic, grounded feedback builder from ticked observables, next-band targets and teacher observations; first name once; stale detection by selection hash. | Fixed |
| F33 | Config / scopes | Medium | `appsscript.json` | No explicit OAuth scopes; adding DocumentApp/external requests changes consent silently. | Declare explicit minimal scopes. | Fixed |
| F34 | Launcher | Medium | `WebAppIntegration.gs:1378`, `docs/LAUNCHER.md` | Launcher embeds the app in an iframe, but HtmlService defaults to `X-Frame-Options: SAMEORIGIN`; embedding cannot work as documented. | Opt-in `LESSON_GRADER_ALLOW_EMBED=true` → `ALLOWALL` (clickjacking trade-off documented); default stays safe; launcher already offers "open in tab". | Fixed |
| F35 | Rubric persistence | High | `RubricEngine.gs:770-801` | Rubric profiles are JSON only; workbook edits cannot flow back into the app (SC14). | Mirror activated rows into a human-editable `RubricRows` sheet; detect drift; "Reload from workbook" previews and activates a new version (existing assessments stay pinned). | Fixed |
| F36 | Teacher level override | High (policy) | `RubricEngine.gs:438-499` | No criterion-level override exists (only the distinct full-marks override). The rule for how a teacher-selected level interacts with observables is not documented (SC18/SC25). | Do not invent. Keep observables as the only scoring input; expose the level read-only; record the open decision; health check reports it. | Blocked — policy decision |
| F37 | Workbook setup | Medium | `WebAppIntegration.gs:437-539` | No "create from Classroom assignment" flow; category must be preconfigured via Script Property or be the only folder. | One flow: course → assignment → workbook name defaults → create + task tab + link + enable grading. Category can be chosen. | Fixed |
| F38 | Resume | Medium | `WebApp.html` | No resume of last workbook/task/student (SC01). | Remember last context per browser (convenience only) and re-validate against server on load. | Fixed |
| F39 | Landing page | Medium | `WebAppIntegration.gs:235-280` | Workbook list returns name only; no class/task/updated info (SC02). | Parse name into course/year/task, add last-updated and task tabs. | Fixed |
| F40 | Error handling | Medium | `WebGrading.gs` (many) | Server throws raw errors; `google.script.run` failure messages reach the toast, but save status text can still read "Draft saved" from an earlier save. | Explicit save-state machine (`Unsaved` / `Saving…` / `Saved <time>` / `Save failed — not saved`). | Fixed |
| F41 | Dead code | Low | `WebAppIntegration.gs:675-1218` | `getTaskAssessmentOverview_` and `getSubmissionAssessmentDetail_` (~540 lines) are unused by any UI. | Keep the public `apiWeb*` names (compatibility) but mark deprecated; remove later. | Deferred |
| F42 | Invented content | Medium | `RubricEngine.gs:1106-1147`, `WebApp.html:1326-1334` | "Generate Formative Rubric" fabricates generic criteria and outcome codes; UI discards the result. | Remove the button from the UI; keep the server function for compatibility, labelled as a template only. | Fixed |
| F43 | Accessibility | Medium | `WebApp.html` | Body text 12–13px, touch targets 32–38px, icon-only mobile nav labels at 10px, no `aria-live` save status, observables are tiny checkboxes. | Mobile-first rewrite: 16px body, ≥44px targets, labelled controls, focus management, live regions. | Fixed |
| F44 | Report | Low | `WebGrading.gs:611-623` | Report renders the pinned profile but does not include deductions or approval snapshot values. | Render from the `ApprovedGrades` snapshot. | Deferred |
| F45 | Tests | Medium | `tests/` | Tests encode some fake behaviour as correct (messages "sent", prefill "updated"). No tests for boundaries, flags, publishing, unlock, sync triggers, queue, de-identification or UI failure states. | Replace with honest-behaviour tests and add the missing suites. | Fixed |
| F46 | Date/time | Medium | `WebAppIntegration.gs:590-596`, `Code.gs:1526` | Classroom due time is UTC; both paths treat a missing time as 23:59 **UTC** (≈ 09:59/10:59 next day Sydney) rather than end of day Sydney. | Missing due time → 23:59 Australia/Sydney; store ISO UTC; display in Sydney time. | Fixed |

---

## 4. Grading rules investigation (Phase 1)

Evidence is from the current repository. "Confirmed" means the code implements it consistently and tests cover
it; it does **not** mean the school policy confirms it.

| # | Question | Repository evidence | Status |
|---|---|---|---|
| 1 | **B and D criterion contributions** | Default band weights `B: 0.875`, `D: 0.575` (`RubricEngine.gs:37`, `WebAccess.gs:100`, `Code.gs:22,158`); overridable per workbook in the `Setup` sheet `GRADE SCALE` table (`RubricEngine.gs:686-695`). | **Implemented, policy-unconfirmed.** The brief confirms only A/C/E. |
| 2 | **Observable weighting** | Each observable in a band earns `maxMarks × bandWeight ÷ observablesInBand` (`RubricEngine.gs:474-477`). No per-observable weights exist. Ticks across several bands are summed and **capped at the criterion maximum** (`RubricEngine.gs:480-483`). | **Confirmed (equal within a band).** |
| 3 | **Teacher-selected criterion level** | No level selector exists. The letter shown per criterion is derived from the criterion percentage (`RubricEngine.gs:486`). The only teacher override is the "distinct" checkbox, which awards full marks (`RubricEngine.gs:448-460`). | **Unresolved.** Not implemented; must not be invented. |
| 4 | **"Bottom of the grade band" rule** | No occurrence in code, tests or docs. | **Unresolved / not implemented.** |
| 5 | **Overall A–E boundaries** | A ≥ 85, B ≥ 75, C ≥ 65, D ≥ 50, E ≥ 0 (`RubricEngine.gs:38-44`); overridable per workbook (`Setup` → `OVERALL GRADE BAND`). Same boundaries also convert each criterion's percentage to its letter. | **Implemented, policy-unconfirmed.** |
| 6 | **Rounding / exact boundaries** | Letter = first band with `percent >= min`, no rounding (`RubricEngine.gs:98-104`) → an exact 85.00 is A, 84.99 is B. Displays use `toFixed(2)`/`toFixed(1)`; explanation uses `Math.round`. Floating-point noise can push an exact boundary below `min` (F10). Legacy Classroom grade rounds to 0.1 (`Code.gs:1941`). | **Implemented; float defect.** |
| 7 | **Classroom representation** | Legacy only: numeric `assignedGrade = total ÷ rubricMax × coursework.maxPoints`, rounded to 0.1, written to `assignedGrade` and `draftGrade` (`Code.gs:1931-1942`). Letters are never sent. Web app: nothing. | **Partially implemented (legacy).** |

### Conflict with SC19

Under the repository defaults, ticking every observable in a single band gives:

| Band ticked in full | Repository result | SC19 target | Conflict |
|---|---|---|---|
| A | 100% | 100% | none |
| C | 70% | 50% | **yes** |
| E | 25% | 0% | **yes** |

Changing defaults would silently re-grade every un-approved draft and change future results for existing
workbooks. Policy for this review: **existing workbooks keep their configured scale**; the target rule set is
provided as an explicit, testable option whose B and D values must be confirmed before it can be activated.

---

## 5. Smallest working workflow and highest risks (Phase 1)

**Smallest end-to-end flow that works today:** only through server calls or a workbook that already has an
active rubric. In the browser a teacher can: open workbook → open task → add a manual student → tick observables
→ save draft → approve. A *fresh* workbook cannot get a rubric (activate button dead) or Classroom submissions
(import button dead) through the UI, and approval silently ignores unsaved edits.

**Highest-risk gaps:**

1. F01 — Gemini key can be persisted in settings JSON and returned to the browser.
2. F02/F03/F04 — consequential actions report success without doing anything.
3. F07 — approving silently discards on-screen edits.
4. F08 — legacy approval half-applies.
5. F09 — reference-point conflict with the confirmed SC19 rule.
6. F16/F14 — no automatic sync and no publishing in the web product.
7. F06/F05 — most primary controls dead; mobile layout not actually applied.

---

## 6. Phase 2+ work log

| Commit | Group | Findings / criteria |
|---|---|---|
| `51c3afd` | fix: secrets and false success | F01–F05, F08, F28, F34 |
| `c95948f` | fix: float-safe boundaries, SC19 reporting, reference rubric | F09 (mitigated), F10, SC15/16/19/20 |
| `19aed66` | feat: flags, unlock, comparison, AI layer, queue, prompts, refinement | F12, F13, F15, F19–F23, SC17, SC22–23, SC26, SC38, SC41–49 |
| `4ed4fef` | feat: hourly library sync, safe import | F16–F18, F26, F27, F39, F46, SC07–SC12 |
| `78a1980` | feat: publishing, feedback builder, rubric round-trip, health check, scopes | F14, F25, F32, F33, F35, F37, SC04, SC13/14, SC31–35, SC39/40 |
| `e964fca` | ui: mobile-first client rebuild | F06, F07, F11, F24, F29, F30, F38, F40, F42, F43, SC01/02/05/06/24/27/29/30/36/37/51 |

### 6.1 Behaviour changes (documented, with reasons)

1. **Sync no longer creates a new version for un-approved work** when attachments change (previously it did, leaving
   the teacher's draft on a non-current row). Drafts are now updated in place, kept, and flagged `EVIDENCE_CHANGED`.
   **Approved/locked work is never superseded by sync**: new evidence is held in `PendingEvidenceJSON` and used only
   when the teacher deliberately presses Reassess (SC12, SC26).
2. **Locked now means final**: reassessment of a locked version requires an authorised unlock first (SC41).
3. **Approval requires resolved blocking flags** (previously only "at least one tick"); a criterion can be approved at
   0 marks with an explicit "No creditable evidence" confirmation (F12, SC38).
4. **AI ticks need a cited file** or they are discarded and flagged (SC22). AI never overwrites saved teacher ticks.
5. **Settings no longer accept grade thresholds/weights** (they never affected grading); grading policy lives in each
   workbook's `Setup` sheet and is pinned into each rubric version at activation.
6. **Classroom message "send" records drafts only** and says so; register and sample actions report real outcomes.
7. Missing Classroom due time is 23:59 **Sydney**, not 23:59 UTC.

### 6.2 Architecture

Before: one 3,100-line legacy file, a 1,350-line web service mixing state, AI, Classroom import and audit
prototypes, and a single 1,350-line HTML file with styles, markup and script.

After (one responsibility per file; all Apps Script globals, load-order independent — enforced by `npm run check`):

| Concern | File(s) |
|---|---|
| Configuration, property names, settings validation | `AppConfig.gs` |
| Access control, credentials, provider config | `WebAccess.gs` |
| Drive library, workbook validation/creation, Classroom link, `doGet`, launcher probe | `WebAppIntegration.gs` |
| Data access helpers + assessment state/versioning (draft, approve, lock, unlock, reassess, compare, import) | `WebGrading.gs` |
| Pure rubric + deterministic grading | `RubricEngine.gs` |
| Pure approval flags | `AssessmentFlags.gs` |
| AI providers, validation, drafting | `AiGateway.gs`, `PromptRegistry.gs`, `AiQueue.gs` |
| Evidence reading + Markdown history | `EvidenceService.gs` |
| De-identification + refinement examples | `Deidentify.gs`, `Refinement.gs` |
| Feedback building / output format | `FeedbackBuilder.gs` |
| Classroom publishing | `Publishing.gs` |
| Automatic sync + triggers | `SyncService.gs` |
| Evaluation / register | `EvaluationRegister.gs` |
| Self-test | `HealthCheck.gs` |
| UI | `WebApp.html` (shell) + `WebAppStyles.html` + `WebAppClient.html` (Api wrapper, router, views) |
| Legacy Jewellery Sheets Studio (kept, shared helpers) | `Code.gs` + legacy dialogs |

Every externally called `apiWeb*`/`api*` name that existed before still exists with a compatible signature
(new optional trailing parameters only). New endpoints are listed in README.md.

### 6.3 Migration path (no data rewrite)

- Core sheets/columns are unchanged. New **columns** (`Submissions`: EvidenceChangedAt, PendingEvidenceJSON,
  UnlockedAt/By/Reason; `CriterionAssessments`: TeacherReviewJSON, AIProposedChecksJSON; `SubmissionFiles`:
  EvidenceStatusUpdatedAt; `ClassroomConfig`: LastSyncedAt, LastSyncResult) are appended on first write.
- New **sheets** are created when first needed: `RubricRows`, `AiQueue`, `EvidenceVersions`, `PublishLog`,
  `PromptTests`, `RefinementExamples`, `EvaluationRegister`.
- Existing drafts with feedback text will show `FEEDBACK_NOT_REVIEWED` until the teacher confirms or rebuilds the
  feedback once. Existing approvals and snapshots are untouched.
- Existing rubric profiles keep their pinned weights; `RubricRows` is filled the next time a rubric is activated.
- If `LESSON_GRADER_SETTINGS` ever contained a Gemini key (old "Save All Settings"), it is ignored and removed on the
  next settings save. **Rotate that Gemini key** (Google AI Studio → API keys) and save the new one via Settings.

### 6.4 Removed

| Item | Reason |
|---|---|
| Fake "evidence anchors" (`Page 1 … Iterative Test 1`) in the studio | Invented evidence locations |
| "Generate Formative Rubric" button | Fabricated generic criteria and outcome codes; result was discarded. Server function kept (labelled template-only) for compatibility |
| Settings thresholds/deductions UI | Never affected grading (misleading). Workbook `Setup` sheet is the policy source |
| Old single-file `WebApp.html` contents | Replaced by shell + partials; same file name and `doGet` contract |
| Old `docs/preview-mock.cjs` and UI tests contents | Rewritten for the new API and UI (files kept) |

No source file was deleted.

### 6.5 Final verification (Phase 7)

| Command | Baseline | Final |
|---|---|---|
| `npm run check` | pass (12 files) | pass (25 files, load-order independence, rendered page, client↔server API map, explicit scopes) |
| `npm test` | 53 pass / 0 fail | **111 pass / 0 fail** (14 suites) |
| Headless Chromium (Playwright 1.56, mock preview) at 360 / 768 / 1280 px | not run | no horizontal scroll, 16px body text, flagged-first order on 360px, no console errors |

Everything above is **mocked or emulated**. No live Apps Script deployment, Classroom, Drive, Sheets, Gemini, Qwen or
JEV was reachable. See ACCEPTANCE.md and docs/LIVE_TESTING.md.

### 6.6 Remaining risks

1. **Grading policy** (SC19, SC18, rule 4): B/D values, level-override rule, "bottom of band" rule and whether overall
   boundaries must change with the target reference points are unresolved. Current weights stay in force.
2. **Classroom grade write-back** only works for assignments created by this script's Cloud project; others need
   manual entry. Private comments cannot be written by the API.
3. **JEV/Qwen** contracts are assumed OpenAI-compatible (A2, A3); unverified.
4. **Voice input** depends on the browser's Web Speech API (Chrome/Edge/Android Chrome; Safari partial; Firefox none).
5. **Drive Markdown export** (`text/markdown`) is used for Docs with a plain-text fallback; PDFs are not converted to
   text (they go to Gemini directly, or are disclosed as unreadable for JEV).
6. Legacy Sheets Studio still has the `window.opener` dictation defect and its own Jewellery-only paths.
