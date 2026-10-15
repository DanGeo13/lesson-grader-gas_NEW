# Privacy and de-identification policy

Applies to AI drafting, the JEV evaluator, refinement examples and prompt testing. Implemented in
`Deidentify.gs`, `AiGateway.gs`, `Refinement.gs`; tested in `tests/ai.test.cjs`.

## What goes where

| Destination | What it receives | De-identified? |
|---|---|---|
| Grading workbook (school Google account) | Full student records, evidence text history (`EvidenceVersions`), AI run logs | No — this is the school's own record |
| **Gemini** (Google AI, `draftProvider = gemini`) | The task rubric and the student's files themselves (Docs as text, PDFs/images inline) | **No.** Files can contain names and faces. Use only if the school permits Gemini for student work. |
| **Qwen 3-VL** (owner-hosted, `visionProvider = qwen`) | Individual **images** for description | **No** (images cannot be de-identified). The host must be under the school's/owner's control. |
| **JEV** (owner-hosted, `draftProvider = jev`) | De-identified Markdown evidence + rubric + task name | **Yes**, see below |
| `RefinementExamples` sheet | AI-vs-teacher tick differences, de-identified notes and feedback, pasted examples | **Yes** |

## What is removed before JEV or refinement storage

- the student's full name, preferred name and every name part of 3+ letters;
- every other student name on the workbook roster (classmates named in group work);
- email addresses, the Classroom user ID, submission IDs, long digit runs (student numbers);
- Australian mobile and landline numbers, URLs, and long opaque IDs;
- original **file names** (often contain student names) — replaced by labels "File 1", "File 2"; evidence notes
  are mapped back to real names only inside the workbook.

A final check (`Deidentify.assertClean`) refuses to send if the student's name, email or IDs still appear.

## Known limitations — not anonymity

- Names that are misspelt, nicknames, or people not on the roster are **not** detected.
- Content inside images, handwriting and scanned PDFs is not text and is never sent to JEV unless a vision
  provider describes it first; that description is then de-identified like any other text.
- Context can still identify a student (a unique project, a sibling's name, a local business).
- De-identification is best-effort filtering, not a guarantee. The teacher remains responsible for reviewing
  examples before sharing them outside the workbook.

## Verification

`tests/ai.test.cjs` builds a JEV payload from a document containing the student's name, a classmate's name, an
email address, a phone number, a student number and an identifying file name, and asserts none of them reach the
JEV request body. This is **mock-level** verification; inspect a real payload in `AIAssessments.ResponseJSON`
(`payloadPreview`, first 2,000 characters) after a live run.
