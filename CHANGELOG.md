# Changelog

## 2026-10-15 — Full project review (branch `claude/bold-lovelace-mg2bwl`)

### Security and honesty
- API keys can no longer be stored in or returned from settings; legacy settings JSON is scrubbed (rotate any key
  saved through the old "Save All Settings").
- Removed false success: Classroom messages are recorded as drafts ("nothing was sent"); register entries are really
  written to a Doc and/or sheet; work-sample copies report per-file results; provider tests are real.
- Explicit OAuth scopes and user-accessing web app model in the manifest; framing blocked unless opted in.
- Legacy studio: escaped all interpolated values; fixed the missing rich-text helper that half-applied approvals.

### Grading
- Float-safe boundaries; reference-point report and target rule set (B/D never invented); Classroom number mapping
  with half-up rounding; hand-worked reference rubric (R1–R10).
- Blocking vs advisory flags enforced at approval; explicit "no creditable evidence"; feedback must match the latest
  selections.

### Workflow
- Create a workbook straight from any Classroom assignment; library-wide hourly sync after due dates; idempotent
  sync that keeps drafts and never displaces approved results.
- Authorised unlock; reassessment comparison; separate selective publishing with preview-bound confirmation.
- AI: evidence-cited ticks only, readable/unreadable file disclosure, Markdown evidence history, de-identified JEV
  path, optional Qwen 3-VL vision, outage queue, prompt test-before-activate, de-identified refinement examples.
- Grounded feedback builder; configurable output; evaluation/register entries from workbook facts.
- Read-only health check.

### UI
- Rebuilt mobile-first client (shell + partials), central API wrapper, stale-response guards, save states, dictation
  on every text box, flagged-first phone layout, dark mode and print styles.

### Tooling
- `npm run check` enforces load-order independence, rendered page validity and client↔server API coverage.
- Tests grew from 53 to 111.
