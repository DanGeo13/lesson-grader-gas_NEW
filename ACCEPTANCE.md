# Acceptance criteria — SC01–SC54

Statuses: **Pass** (required behaviour verified at the appropriate level), **Fail**, **Partial**, **Blocked**
(needs credentials, infrastructure or a policy decision), **Not verified** (implemented but not yet verifiable
here). A criterion that depends on a live Google or AI service cannot be **Pass** from mocked tests alone.

Baseline taken on 15 October 2026 at commit `4938680` (see REVIEW.md §2). "Final status" is filled in Phase 7.

| Criterion | Baseline status | Existing implementation | Planned change | Verification method | Evidence | Final status |
|---|---|---|---|---|---|---|
| SC01 Resume assessment | Fail | Workbook state is persistent in Sheets, but the UI always starts at Overview with no resume. | Remember last workbook/task/student per browser; re-validate on load; "Resume" card. | jsdom test + manual reopen. | `WebApp.html:1336-1352` | — |
| SC02 Workbook landing page | Partial | `LessonGraderWeb.list` lists validated library workbooks by name only. | Parse course/year/task, add updated time and task tabs; search. | Unit test of listing + manual two-workbook check. | `WebAppIntegration.gs:235-280` | — |
| SC03 All assignments selectable | Partial | `listCourseWork` paginates all states incl. undated; Classroom picker in web UI is broken (`loadClassroom` undefined). | Working picker; course list pagination; undated shown with "No due date". | Unit test with paged/undated mocks; live check. | `Code.gs:1446-1485`, `WebApp.html:699` | — |
| SC04 One workbook per class task | Partial | Name scheme `CODE [YEAR] - Task - Grading`; link stored in `ClassroomConfig`; create form not wired. | "Create from Classroom assignment" flow creating workbook + task + link. | Unit test; live two-task check. | `Code.gs:898-933`, `WebAppIntegration.gs:437-673` | — |
| SC05 Validated workbook context | Partial | Every server call takes an explicit workbook ID validated against the library. Client has no stale-response guard. | Client context tokens; writes carry captured context; isolation test across two workbooks. | Unit + jsdom tests. | `WebAppIntegration.gs:141-163` | — |
| SC06 Complete phone and laptop workflow | Fail | Responsive CSS exists but viewport not applied by HtmlService; many actions unwired. | Mobile-first rewrite, `addMetaTag`, all actions wired. | jsdom + headless browser at 360/768/1280; live device check required. | `WebAppIntegration.gs:1378` | — |
| SC07 Stable student/submission identity | Partial | Import keys by Classroom course + coursework + `userId`; UI uses `SubmissionRecordID`. Class tracking stores the wrong ID. | Fix tracking; similar-name test. | Unit test with similar names. | `WebGrading.gs:690-693,1199` | — |
| SC08 Automatic due-date import and hourly sync | Fail | Hourly trigger exists only for the bound legacy sheet (`getActiveSpreadsheet`). | Library-wide hourly sync service + trigger management. | Unit tests with clock/trigger mocks; live trigger check required. | `Code.gs:1544-1635,1965-1983` | — |
| SC09 Manual processing of undated assignments | Partial | `importTask` works for undated links (server); UI import button unwired. | Wire "Sync now". | Unit test + live. | `WebGrading.gs:664-754` | — |
| SC10 Manual sync always available | Partial | Server import exists; no UI. | "Sync now" on every task regardless of trigger state. | jsdom + live. | — | — |
| SC11 Idempotent synchronisation | Partial | Unchanged attachments are skipped (tested once). | Three-sync count test; metadata refresh without duplication. | Unit test. | `tests/web-grading.test.cjs` import test | — |
| SC12 Approved assessments protected from sync | Partial | New evidence creates a new version; approved rows untouched. Same-file edits undetected. | Evidence-changed flag on approved versions; deliberate reassess. | Unit test. | `WebGrading.gs:697-747` | — |
| SC13 Complete rubric loading | Partial | Profiles carry criteria, observables, bands, weights, marks; no human-readable mirror. | `RubricRows` mirror + full field comparison test. | Unit test. | `RubricEngine.gs:139-233` | — |
| SC14 Bidirectional rubric persistence | Fail | App → JSON profile only. | Workbook `RubricRows` round trip with explicit activation. | Unit test + live. | — | — |
| SC15 Correct weighted observable calculation | Pass (unit) | Per-band equal shares × criterion marks; unequal criterion marks supported. | Reference rubric with hand-worked results. | Unit tests. | `tests/rubric-engine.test.cjs:27-43` | — |
| SC16 Reproducible results | Pass (unit) | Scoring is a pure function of profile + ticks; feedback not an input. | Add explicit feedback-variation test. | Unit test. | `RubricEngine.gs:438-528` | — |
| SC17 AI cannot directly assign final marks | Pass (unit) | AI returns ticks; marks recomputed; AI-proposed grade informational only. | Test with arbitrary AI score field. | Unit test. | `tests/web-grading.test.cjs` AI test | — |
| SC18 Teacher level and observable refinement | Partial | Observables editable (save/reload works); no level selection; rule undocumented. | Keep observable refinement; level displayed read-only; open decision recorded. | Unit + jsdom. | `RubricEngine.gs:438-499` | — |
| SC19 Agreed criterion reference points | Fail | Repository weights give C = 70%, E = 25%. | Explicit rule-set config; target set offered, blocked until B/D confirmed; no historical recalculation. | Reference rubric tests. | `RubricEngine.gs:37` | — |
| SC20 Explicit grade boundaries | Partial | `>= min` boundaries, no rounding; float noise defect. | Normalised comparison; boundary tests below/at/above each band. | Unit tests. | `RubricEngine.gs:98-104` | — |
| SC21 Correct assessment evidence context | Partial | AI uses only the selected submission's files and the pinned task profile. | Cross-student contamination test; evidence Markdown per student. | Unit test. | `WebGrading.gs:524-609` | — |
| SC22 Evidence-linked suggestions | Fail | Ticks accepted without evidence notes. | Ticks require evidence notes; insufficient evidence flagged. | Unit test. | `WebGrading.gs:577-595` | — |
| SC23 Unreadable evidence disclosed | Partial | Skipped files listed in API response only. | Persist + display readable/unreadable status; acknowledgement flag. | Unit + jsdom. | `WebGrading.gs:535-549` | — |
| SC24 Evidence alongside assessment | Partial | Split layout exists; evidence only links out; anchors are fake. | Evidence panel with inline preview, edits preserved while browsing. | jsdom. | `WebApp.html:313-332,912` | — |
| SC25 Teacher override and practical observations | Partial | Teacher can untick/tick AI proposals and type notes per criterion. | Show AI vs teacher markers; general practical observation field. | Unit + jsdom. | `WebApp.html:922-954` | — |
| SC26 Teacher changes preserved | Partial | Saved drafts persist; sync never overwrites drafts (creates versions). | Test refine → save → reload → sync. | Unit test. | `WebGrading.gs:276-339` | — |
| SC27 Mobile flagged-first review | Fail | No flags; no ordering. | Flags; mobile ordering with unflagged criteria collapsed. | jsdom + headless browser. | — | — |
| SC28 Whole-class drafts, individual review | Partial | Bulk AI server function (synchronous); roster opens individual Studio. | Queue-based batch with readiness. | Unit + jsdom. | `WebGrading.gs:756-791` | — |
| SC29 Typed and dictated input | Partial | Dictation on criterion notes and two prefill fields only; single-shot. | Dictation on all assessment text fields incl. feedback and personal evaluation. | jsdom with mocked SpeechRecognition; live device check. | `WebApp.html:1066-1099` | — |
| SC30 Dictation editable and non-publishing | Partial | Dictation appends to an editable textarea; never saves. | Keep; explicit states; test. | jsdom. | — | — |
| SC31 Teacher observations incorporated | Fail | No feedback rebuild. | Grounded feedback builder includes observations. | Unit test. | — | — |
| SC32 Feedback consistent with revised assessment | Fail | No stale detection. | Selection hash + stale advisory flag + rebuild action. | Unit test. | — | — |
| SC33 Grounded final feedback | Fail | Feedback is free text; legacy AI prompt asks the model to cite materials it may not see. | Builder uses only ticked observables, next-band targets and teacher notes. | Unit test. | `Code.gs:2185-2198` | — |
| SC34 Natural first-name usage | Fail | Not controlled. | First name used once at the opening when enabled. | Unit test. | — | — |
| SC35 Configurable output format | Partial | Feedback template tokens exist; preview button unwired; Classroom mapping undocumented. | Output format settings + preview; documented numeric mapping; letter in comment text. | Unit + jsdom; live Classroom check. | `RubricEngine.gs:1006-1041` | — |
| SC36 Clear assessment states | Partial | Status pill shows raw status; no Reassessment/Locked labels in UI. | State labels: Draft, Approved, Reassessment vN, Final (locked). | jsdom. | — | — |
| SC37 Complete pre-approval review | Fail | Confirm dialog only. | Approval review panel: marks, feedback, selected observables, flags. | jsdom. | — | — |
| SC38 Required flags resolved before approval | Partial | Server blocks criteria with zero ticks only. | Blocking vs advisory flags enforced server-side. | Unit + jsdom. | `WebGrading.gs:369-371` | — |
| SC39 Approval saves without publishing | Pass (unit, server) | Approve writes `ApprovedGrades`, never calls Classroom. UI approves stale drafts (F07). | Keep; fix UI. | Unit + jsdom. | `WebGrading.gs:379-408` | — |
| SC40 Selective explicit publishing | Fail | Not in web app. | Publishing module with selection + confirmation. | Unit (mock Classroom); live check blocked (A4). | — | — |
| SC41 Explicit final-lock protection | Partial | Lock blocks edits; no unlock. | Authorised unlock with reason. | Unit + jsdom. | `WebGrading.gs:412-432` | — |
| SC42 Reassessment preserves history | Partial | New version row; old approval retained; no change comparison. | Version comparison view/data. | Unit test. | `WebGrading.gs:434-476` | — |
| SC43 De-identified JEV payload | Fail | No JEV. | Deidentify module + JEV payload builder; policy documented. | Unit test; live JEV blocked. | — | — |
| SC44 Markdown history retained | Fail | Not stored. | `EvidenceVersions` append-only by hash. | Unit test. | — | — |
| SC45 De-identified refinement examples | Fail | Not stored. | `RefinementExamples` on approval + uploaded examples, de-identified. | Unit test. | — | — |
| SC46 Prompt changes tested before activation | Fail | Prompts hardcoded. | Prompt versions, test runs vs approved portfolios, activation gate. | Unit test; live AI blocked. | — | — |
| SC47 Approved results protected from prompt changes | Partial | AI cannot change approved work (editable() guard). No prompt comparison. | Test-run report lists potential changes only. | Unit test. | `WebGrading.gs:146-149` | — |
| SC48 AI outage choices | Fail | Errors only. | Queue or continue manually. | Unit + jsdom. | — | — |
| SC49 Queue resumes and reports readiness | Fail | No queue. | Time-triggered processor, dedupe, readiness counts. | Unit test. | — | — |
| SC50 Manual marking during AI outage | Partial | Manual save/approve do not depend on AI (server). | Test with AI unavailable; UI does not gate on AI. | Unit + jsdom. | — | — |
| SC51 Honest save status | Fail | Toast on failure, but several server actions return false success (F02–F04) and save label can stay stale. | Save-state machine; remove false success. | Unit + jsdom. | `WebGrading.gs:1134,1027` | — |
| SC52 Durable saved data | Partial | Sheets persistence real; not verified live. | Persistence round-trip tests; live check. | Unit + live. | — | — |
| SC53 Correct evaluation/register context | Fail | Prefill invents text; destination not written. | Real context pre-fill; validated destination. | Unit test; live Doc check. | `WebGrading.gs:970-1037` | — |
| SC54 Reviewable evaluation entry | Partial | Typed/dictated fields exist; commit is fake. | Preview → edit/dictate → commit to Doc/sheet. | Unit + jsdom; live. | — | — |

**Baseline counts:** Pass 4 (SC15, SC16, SC17, SC39 — unit level only) · Fail 21 · Partial 29 · Blocked 0 ·
Not verified 0 · **Total 54**.
