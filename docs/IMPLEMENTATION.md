# Teacher's Studio — Lesson Grader Implementation & Deployment Guide

**Status (26 September 2026):** Complete Teacher's Studio platform upgrade supporting universal project rubrics, deterministic scoring, Drive audit integration, and teacher-controlled moderation. Fully tested with 53 automated unit and integration tests.

---

## 1. System Architecture & Capabilities

### Phase 1: Grading and Rubrics
- **1.1 Editable A–E Thresholds:** Configurable percentage thresholds versioned in Settings with validation (blocks gaps, overlaps, inverted bounds) and recalculation preview.
- **1.2 Discrete Observables & Partial Credit:** Each criterion contains discrete observables per grade band with equal mark shares. Partial credit is deterministically summed without guessing.
- **1.3 Formative Rubric Generation:** AI generates draft formative rubrics tailored to interim milestones with syllabus outcome tags (`DT5-1`, etc.).
- **1.4 Reassessment & Version Pinning:** Submissions pin immutable `RubricProfileID` snapshots; updating a task rubric applies to future drafts only.
- **1.5 Curriculum Outcomes:** Syllabus outcomes mapped to criteria with coverage tally and validation.

### Phase 2: Grading Workflow & Interface
- **2.1 Actual-Page Evidence:** Deep-link evidence indicators (`#page=N`, anchor tags) connecting observables to student Drive folios.
- **2.2 Scrollable Split-Screen Studio:** Left pane holds student document viewer and file attachments; right pane displays rubric criteria, observables, notes, feedback, and score breakdown.
- **2.3 Bulk AI Grading Queue:** Batch proposes drafts for selected students with select-all, progress bar, and per-student error isolation (drafts require teacher approval).
- **2.4 Deterministic Deductions:** Late submission penalties (percentage per day, fixed marks, grace hours, weekend exclusion) calculated strictly from raw marks and shown distinctly.
- **2.5 Unreadable Submission Flagging:** Detects unreadable or missing student files, flagging them for resubmission.
- **2.6 Performance:** Batch range operations for fast Sheets I/O.
- **2.7 Readability & Back Navigation:** Accessible top-center back navigation, keyboard focus management, and dirty-state prompt for unsaved edits.

### Phase 3: Google Classroom Integration
- **3.1 Coursework Linkage:** Links course assignments to workbook tasks without automated background sync.
- **3.2 Incomplete Work Messaging:** Flags unsubmitted, incomplete, or unreadable work and prepares personalized message drafts for teacher preview and explicit confirmation.

### Phase 4: Settings, Tracking and Reporting
- **4.1 Feedback Templates:** Customizable feedback templates with token chips (`{{student}}`, `{{task}}`, `{{grade}}`, `{{final_score}}`, etc.), live sample preview, and token validation.
- **4.2 Formative/Summative Classification:** Distinguishes formative check-ins from summative assessments across rosters and reporting.
- **4.3 AI Provider Configuration:** Gemini key storage (Script Properties only), model picker (`gemini-2.5-flash`), Qwen 3VL configuration, and `9JEV` subject code tag.
- **4.4 Versioned Recalculation:** Threshold changes are previewed against existing cohorts before saving; historical approved grades remain unchanged.
- **4.5 Multi-Term Class Table:** Multi-term tracking matrix displaying student progression, raw scores, deductions, final scores, grades, and syllabus outcome coverage, with sync to `ClassTracking`.

### Phase 5: Teacher's Studio Platform
- **5.1 Platform & Module Branding:** Teacher's Studio navigation shell hosting Lesson Grader, Class Table, and Registration & Samples.
- **5.2 App Registry & Style Guide:** Modular app registry with consistent design tokens, semantic color palette, and full responsive support.

### Phase 6: Registration, Evaluation & Work Samples (School Audit)
- **6.1 Drive Destination Discovery:** Discovers and validates school Registration & Evaluation doc IDs and HIGH/MED/LOW sample folders.
- **6.2 Prefill Designated Document:** Automatically pre-fills task outcomes, dates, and approved grade distribution; generates safe export markdown.
- **6.3 Teacher Evaluation Dictation:** Voice dictation via Web Speech API with fallback to typed input for teacher variations and reflections.
- **6.4 Cohort Grading Lock:** Verifies all cohort marks are approved and sets an explicit audit lock before sample selection.
- **6.5 HIGH, MED, LOW Sample Placement:** Proposes representative student folios by percentile or grade band with teacher override and copies files into Drive audit folders upon confirmation.
- **6.6 Audit Dashboard & Audit Trail:** Compliance matrix tracking document status, sample counts, and append-only event logging in `AuditLog`.
- **6.7 Misplaced File Detection:** Identifies files outside expected folder structures.
- **6.8 Drive Safety Guards:** Verifies permissions and confirmation before creating copies or modifying shared documents.

---

## 2. Google Sheets Tab Schemas

| Sheet Name | Role & Schema |
|---|---|
| `RubricProfiles` | Immutable rubric snapshots: `RubricProfileID`, `Task`, `RubricName`, `Source`, `SchemaVersion`, `ProfileJSON`, `Active`, `CreatedAt`. |
| `Submissions` | Student roster: `SubmissionRecordId`, `CourseID`, `CourseWorkID`, `StudentUserID`, `StudentName`, `StudentEmail`, `Task`, `ClassName`, `Status`, `Late`, `Version`, `CurrentOfficial`, `AssessmentType`, `UpdatedAt`. |
| `CriterionAssessments` | Draft criteria states: `AssessmentID`, `SubmissionRecordId`, `CriterionID`, `RubricProfileID`, `ObservablesTicked`, `TeacherWrittenNote`, `CalculatedScore`, `AIProposedGrade`, `UpdatedAt`. |
| `ApprovedGrades` | Immutable approved grade register: `ApprovalID`, `SubmissionRecordId`, `StudentName`, `Task`, `RubricProfileID`, `AssessmentType`, `RawScore`, `DeductionPoints`, `DeductionReason`, `FinalScore`, `MaxMarks`, `Percentage`, `GradeLetter`, `TeacherFeedbackJSON`, `ApprovedBy`, `ApprovedAt`. |
| `AssessmentHistory` | Append-only audit log: `EventID`, `SubmissionRecordId`, `Action`, `PerformedBy`, `Timestamp`, `Revision`, `DetailsJSON`. |
| `ClassroomConfig` | Coursework mappings: `CourseID`, `CourseWorkID`, `CourseName`, `TaskName`, `LinkedAt`. |
| `ClassTracking` | Multi-term progression tracking matrix synced by `apiWebSyncClassTable`. |
| `AuditLog` | Audit actions register for cohort locks, document prefill updates, and sample copies. |

---

## 3. Settings Inventory

| Setting Key | Default / Format | Purpose |
|---|---|---|
| `thresholds` | A: 85–100%, B: 75–84.99%, C: 65–74.99%, D: 50–64.99%, E: 0–49.99% | Cut-off ranges for letter grade conversion. |
| `thresholdsVersion` | String (e.g. `1.0`) | Ensures historical assessments reference their creation scale. |
| `weights` | A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 | Deterministic multiplier for rubric performance levels. |
| `deductions` | `{ type: 'percent_per_day', rate: 10, maxDeductionPercent: 50, excludeWeekends: true, graceHours: 0, minFloor: 0 }` | Rules for late submission mark deductions. |
| `feedbackTemplate` | Multi-line string with `{{student}}`, `{{task}}`, `{{grade}}`, etc. | Template for drafting comprehensive student feedback. |
| `outcomes` | Array of `{ code, title }` (`DT5-1`, `DT5-2`, `DT5-3`, etc.) | Curriculum syllabus outcomes index. |
| `aiProviders` | `{ defaultProvider: 'gemini', geminiModel: 'gemini-2.5-flash', jevSubjectCode: '9JEV' }` | AI provider configuration and subject identifiers. |

---

## 4. Safe Deployment Steps

1. **Verify Google Apps Script Project:** Ensure V8 runtime and Google Classroom API Advanced Service are enabled in `appsscript.json`.
2. **Configure Script Properties:**
   - In Google Apps Script Editor → **Project Settings** → **Script Properties**:
     - `LESSON_GRADER_ALLOWED_EMAILS`: Comma-separated list of authorised teacher emails (e.g. `teacher@school.edu.au,headteacher@school.edu.au`).
     - `GEMINI_API_KEY`: Testing API key for AI draft proposals.
     - `LESSON_GRADER_SETTINGS`: (Optional) Custom JSON settings override.
3. **Deploy Web App:**
   - Click **Deploy** → **New deployment**.
   - Select type **Web app**.
   - **Execute as:** `User accessing the web app`.
   - **Who has access:** `Anyone within [Your School Domain]`.
   - *Never select "Anyone" + "Execute as me".*
4. **Connect Launcher (Optional):**
   - Open the static GitHub Pages launcher (`site/index.html`).
   - In the launcher settings dialog, enter the Apps Script `/exec` URL. The address is saved securely in the teacher's browser localStorage.

---

## 5. Verification & Test Suite

Run tests in your terminal:
```bash
npm ci
npm run check    # Validates syntax, manifests, and file integrity
npm test         # Runs all 53 unit and integration tests across 5 test suites
npm run preview  # Starts local preview on http://0.0.0.0:4173 (sample data only)
```
