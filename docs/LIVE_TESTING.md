# Authorised live verification procedure

Automated tests use fakes. These steps verify the "Not verified" criteria in ACCEPTANCE.md against real Google
services. **Use a test Classroom class with test student accounts and a copy of a workbook — never live student
records.** Record results in ACCEPTANCE.md (Evidence / Final status).

## 0. Preparation (owner)

1. Deploy a **test** web app (see README "Deploy"). Set Script Properties: `LESSON_GRADER_ALLOWED_EMAILS`
   (your account), `LESSON_GRADER_UNLOCK_EMAILS` (your account), Gemini key via Settings.
2. Create a test Classroom class with two test students whose names are similar (e.g. "Sam Lee", "Sam Leigh").
3. Create two assignments: one **with** a due date in the past (graded, e.g. 40 points), one **without** a due date.
   For SC40, also create one assignment **from this script's Cloud project** (Classroom API) if possible.
4. Students submit: a Google Doc with distinct text, a JPG photo, and one unsupported file (e.g. `.stl`).

## 1. Setup and sync (SC01–SC05, SC07–SC12)

| Step | Expect | Criteria |
|---|---|---|
| Landing → New workbook → From Classroom → dated assignment | Workbook created in Drive, linked; banner says it syncs hourly after the due date | SC04, SC03 |
| Repeat for the undated assignment | Second workbook; assignment listed with "no due date" | SC03, SC04, SC09 |
| Landing lists both with class/task/year | SC02 |
| Task → Rubric → load a shared rubric → activate | Active rubric shown; `RubricRows` sheet filled | SC13 |
| Edit a description in `RubricRows` in Sheets → reopen Rubric | Drift banner → review → activate | SC14 |
| Sync now ×3 | First imports 2; next two "unchanged"; row counts unchanged | SC10, SC11, SC07 |
| Settings → Install triggers; wait one hour | `ClassroomConfig.LastSyncedAt` set, AutoImported TRUE | SC08 |
| Open student A in workbook 1, switch to workbook 2, save there | Only workbook 2 changes | SC05 |
| Close the browser, reopen the app | Resume card returns to the last task; saved ticks reload | SC01, SC52 |

## 2. Marking (SC06, SC21–SC30, SC36–SC39, SC41–SC42, SC50–SC51)

Do every step once on a phone and once on a laptop (SC06).

1. Draft with AI → ticks appear only with cited file notes; the `.stl` shows "Unsupported"; acknowledgement is
   required before approval (SC21, SC22, SC23). Compare with the other student's evidence: no cross-contamination.
2. Browse Evidence while unsaved edits exist → edits remain (SC24).
3. Untick an AI suggestion, tick another, add an observation by typing and by dictation (Chrome/Edge/Android);
   deliberately dictate a wrong word and correct it (SC25, SC29, SC30).
4. Save, refresh, Sync now → refinements remain (SC26).
5. Update feedback from rubric & observations → check wording, first name once (SC31–SC34). Preview student comment
   with different output settings (SC35).
6. Change a tick → approval shows "Feedback was written before…" and is blocked until resolved (SC32, SC38).
7. Approve → `ApprovedGrades` row; Classroom grade unchanged (SC37, SC39).
8. Lock → controls read-only; Unlock with reason → Approved; Reassess → v2, comparison shown (SC36, SC41, SC42).
9. Make the save fail (open the same student in two tabs, save in one, then the other) → "Not saved —…" (SC51).
10. Remove the Gemini key → AI unavailable → queue / continue manually; manual save/approve work (SC48, SC50).
    Restore the key → Process AI queue (or wait for the trigger) → draft ready; job not duplicated (SC49).
11. Student resubmits after approval → Sync → approved result unchanged, banner offers Reassess (SC12).

## 3. Publishing (SC40) — grades only, test class only

Select one approved test student → Publish → check destination and number → type PUBLISH. Expect either the draft
grade in Classroom for that student only, or the explicit "only lets an app change grades on assignments that app
created" message. Leave "return to students" off unless testing it.

## 4. AI services (SC43, SC44, SC46, SC47)

1. Configure JEV (and optionally Qwen) in Settings; Test. Draft provider = JEV. Draft a student whose Doc contains
   their name, email and a classmate's name. Inspect `AIAssessments.ResponseJSON.payloadPreview` and the JEV server
   log: no identifiers (SC43).
2. Edit the student Doc and draft again → two rows in `EvidenceVersions` (SC44).
3. Settings → Drafting prompts → new version → Test on current task → PromptTests rows; approved results unchanged;
   Activate requires ACTIVATE (SC46, SC47).

## 5. Register (SC53, SC54)

Evaluation / register → check the pre-filled facts against the workbook → dictate an evaluation → enter a test Doc ID
→ Check destination → commit → the Doc has the new section and `EvaluationRegister` has a row.
