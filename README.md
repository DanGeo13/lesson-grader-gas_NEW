# Lesson Grader

A Google Apps Script web app for rubric-based marking of Technology and Applied Studies projects. Google Classroom
supplies classes, assignments and submissions; Google Drive holds the evidence; **one Google Sheets workbook per class
task** is the assessment record. AI can draft rubric selections with cited evidence, but **marks are always calculated
by a deterministic engine** and **every result is approved by the teacher**. Approval and Classroom publishing are
separate actions.

> Status (15 Oct 2026): implemented and tested with fakes (111 automated tests). **Not yet verified against a live
> deployment, Classroom, Drive, Gemini, Qwen or JEV.** See [ACCEPTANCE.md](ACCEPTANCE.md) and
> [docs/LIVE_TESTING.md](docs/LIVE_TESTING.md). Review notes: [REVIEW.md](REVIEW.md). Plan: [ROADMAP.md](ROADMAP.md).

## Workflow

1. **Landing** — lists your grading workbooks (class, task, year, last update) and offers *Resume*.
2. **New workbook** — pick a Classroom class and any assignment (dated or not); the workbook, task tab, Classroom link
   and grading sheets are created together. A manual mode exists for projects outside Classroom.
3. **Rubric** — load a shared rubric, the active one, the workbook's `RubricRows` sheet, or extract one from a
   PDF/Markdown file; check every row; activate as a new version. Existing marking stays pinned to its version.
4. **Sync** — *Sync now* at any time; hourly automatic sync after the due date once triggers are installed.
5. **AI drafts** — for one student or a selected batch (queued, survives outages).
6. **Assessment Studio** — evidence beside the rubric; tick observables; flagged criteria first on phones; teacher
   observations and feedback by typing or dictation; *Update feedback from rubric & observations*.
7. **Review & approve** — shows marks, selected observables, feedback and any blocking flags. Saves an immutable
   snapshot. Nothing is published.
8. **Publish selected** — preview the students, grades and destination; type `PUBLISH`.
9. **Reassess** new evidence deliberately; **Final lock**; authorised **Unlock**.
10. **Evaluation / register** — facts pre-filled from the workbook; teacher evaluation typed or dictated; committed to
    a Google Doc and/or the workbook.

Screenshots (local mock preview, not real student data):
[landing](docs/screenshots/home-1280.png) · [class on a phone](docs/screenshots/task-360.png) ·
[studio on a phone](docs/screenshots/studio-360.png) · [feedback on a phone](docs/screenshots/studio-feedback-360.png) ·
[studio on a laptop](docs/screenshots/studio-1280.png)

## Grading rules (as implemented)

- Each observable earns `criterion marks × band weight ÷ observables in that band`; ticks in several bands are summed
  and **capped** at the criterion's marks. The distinct override awards full criterion marks.
- Default band weights A 100%, B 87.5%, C 70%, D 57.5%, E 25%; overall and criterion letters A ≥ 85, B ≥ 75,
  C ≥ 65, D ≥ 50, else E. Exact boundaries go to the higher band. Both are set per workbook in its `Setup` sheet and
  **pinned into each rubric version** when activated.
- **Conflict:** the agreed reference points are A 100%, C 50%, E 0%. The engine supports them
  (`RubricEngine.targetRuleSetWeights`) but will not invent B/D, and existing workbooks keep their configured weights.
  The health check reports the difference. Teacher-selected criterion levels are **not implemented** because the rule
  is unresolved. Details and hand-worked examples: [REVIEW.md §4](REVIEW.md), [docs/REFERENCE_RUBRIC.md](docs/REFERENCE_RUBRIC.md).
- Classroom stores numbers only: grade = final marks ÷ rubric total × assignment points, half-up to 1 d.p.
  (configurable). Letters and comments are produced as text to paste; the Classroom API cannot write private comments.

## Approval, publishing and locking

| State | Meaning | Allowed |
|---|---|---|
| Not started / Draft | `New` / `InReview` | edit, save, AI draft, approve when no blocking flags |
| Teacher-approved | `Approved` + `ApprovedGrades` snapshot | publish, lock, reassess |
| Final (locked) | `Locked` | publish; unlock only by `LESSON_GRADER_UNLOCK_EMAILS` with a reason |
| Reassessment vN | new version row; earlier approval kept | as Draft |

Blocking flags: no selection (or explicit "no creditable evidence"), AI found no evidence and the teacher has not
reviewed, unreadable files not acknowledged, feedback older than the latest change. Advisory: capped bands, uncited
AI ticks discarded, teacher changed the AI proposal, no feedback, evidence changed since the draft.

Sync never changes marks. New evidence for approved work is held until the teacher presses *Reassess*.

## Deploy

1. Back up existing grading workbooks (File → Make a copy) and export the current Apps Script project.
2. Push the files with clasp (or paste them into the editor):
   ```bash
   npm i -g @google/clasp && clasp login
   clasp clone <SCRIPT_ID>        # or: clasp create --type standalone
   clasp push                     # pushes *.gs, *.html, appsscript.json
   ```
   `.clasp.json` is git-ignored. All `.gs` files are load-order independent.
   **Run `clasp push` only from this repository's folder.** `clasp push` replaces *every* file in the Apps Script
   project the folder's `.clasp.json` points to. Copy `.clasp.json.example` to `.clasp.json` and paste the Lesson
   Grader script ID (Apps Script editor → Project Settings → Script ID). `.claspignore` limits the push to
   `appsscript.json` plus the 27 root `.gs`/`.html` files; check the list `clasp push` prints before deploying.
3. Project Settings → Script Properties:

   | Property | Required | Purpose |
   |---|---|---|
   | `LESSON_GRADER_ALLOWED_EMAILS` | yes | Comma-separated teacher emails (fails closed if empty) |
   | `LESSON_GRADER_UNLOCK_EMAILS` | recommended | Who may unlock final assessments (nobody if empty) |
   | `LESSON_GRADER_WEB_CATEGORY` | optional | Default Drive category folder for new workbooks |
   | `LESSON_GRADER_ALLOW_EMBED` | optional | `true` lets the GitHub launcher iframe the app (clickjacking trade-off) |
   | `GEMINI_API_KEY`, `QWEN_ENDPOINT_URL`, `QWEN_API_KEY`, `JEV_ENDPOINT_URL`, `JEV_API_KEY` | per provider | Set from the app's Settings; never returned to the browser |
   | `LESSON_GRADER_SETTINGS`, `LESSON_GRADER_PROMPTS`, `LESSON_GRADER_AI_HEALTH` | managed by the app | Validated settings, prompt versions, provider health (no secrets) |

4. Deploy → New deployment → Web app → **Execute as: user accessing**, **Who has access: your domain** (also declared in
   `appsscript.json`). Never "Anyone" with "Execute as me".
5. Open the app → Settings → *Run checks*, *Install hourly sync & AI queue*, AI provider keys/endpoints.

OAuth scopes are declared explicitly in `appsscript.json` (Sheets, Drive, Docs, external requests, triggers,
container UI for the legacy menu, Classroom courses/coursework/rosters/emails).

### AI hosts

- **Gemini**: Google AI Studio key. Receives the student files themselves.
- **Qwen 3-VL** (optional, owner-hosted, always on): an OpenAI-compatible `POST /v1/chat/completions` endpoint over
  HTTPS (e.g. vLLM or Ollama behind a TLS reverse proxy). Used only to describe images for JEV.
- **JEV** (owner-hosted evaluator): OpenAI-compatible chat endpoint that accepts `response_format: json_object`.
  Receives only de-identified Markdown and the rubric. Policy and limits: [docs/PRIVACY.md](docs/PRIVACY.md).
- This repository cannot provision or verify those machines. If a provider is down, teachers can queue or mark manually.

### Voice input

Dictation uses the browser's Web Speech API (`en-AU`, continuous, editable). Works in Chrome/Edge (desktop and
Android); Safari support is partial; Firefox has none. When unavailable, use the device keyboard's microphone key.
Dictation never saves, approves or publishes.

## Workbook structure

Core sheets (unchanged contract): `Submissions`, `SubmissionFiles`, `CriterionAssessments`, `AIAssessments`,
`AssessmentHistory` (append-only; its HistoryID is the save revision), `ApprovedGrades` (immutable snapshots),
`RubricProfiles` (versioned rubric JSON), `Setup`, `ClassroomConfig`, `AuditLog`, `ClassTracking`, `ClassroomMessages`.
Added on demand: `RubricRows`, `AiQueue`, `EvidenceVersions`, `PublishLog`, `PromptTests`, `RefinementExamples`,
`EvaluationRegister`. Migration notes: [REVIEW.md §6.3](REVIEW.md).

## Develop and test

```bash
npm ci
npm run check     # syntax, load-order independence, rendered page, client↔server API map, manifest scopes
npm test          # 111 tests: grading reference cases, flags, sync, AI, publishing, feedback, UI (jsdom)
npm run preview   # http://localhost:4173 mock UI · /launcher GitHub Pages launcher (mock data only)
```

## Troubleshooting

| Symptom | Fix |
|---|---|
| "Access denied" | Add your exact email to `LESSON_GRADER_ALLOWED_EMAILS` |
| Workbook not listed | It must be under *My Drive › Graded Assessments › Stage N › code › category* and have `Submissions`/`ClassroomConfig` |
| "Enable grading" banner | Press it; a Drive backup is made first |
| Due-dated work not arriving | Settings → install triggers (they run as the installing teacher) |
| Publish says the app cannot change grades | Classroom only allows grade changes on assignments created by this script's project; enter manually |
| "Feedback was written before…" | Update feedback from rubric, or tick "This feedback matches…", then save |
| Launcher shows a blank frame | Use "Open in tab", or set `LESSON_GRADER_ALLOW_EMBED=true` |

The original container-bound Sheets Studio (`Code.gs` menu and dialogs) remains for the legacy Jewellery markbook only.
