# Roadmap

Required criteria stay labelled **Required** until verified Pass in ACCEPTANCE.md.

| Item | Required criterion or optional enhancement | Effort | Value | Dependency / blocker | Implementation note |
|---|---|---|---|---|---|
| Confirm B/D contributions, overall boundaries and "bottom of band" rule | **Required** SC19, SC20 | S | H | School policy decision | Then set `Setup` weights via `targetRuleSetWeights`; activate new rubric versions; never recalculate approvals |
| Decide teacher-selected criterion level rule | **Required** SC18, SC25 | M | H | Policy decision | Options: adjust selections / separate moderated result; implement as pinned rule with tests |
| Live verification run | **Required** all "Not verified" | M | H | Test class, test accounts, deployment | Follow docs/LIVE_TESTING.md and update ACCEPTANCE.md |
| Classroom grade publishing for teacher-created assignments | **Required** SC40 | M | H | Classroom API only patches project-created coursework | Option: create assignments from the app (so the project owns them) |
| JEV endpoint contract | **Required** SC43, SC48 | S | H | Owner-hosted service details | Confirm OpenAI-compatible JSON mode or add an adapter in `AiGateway.callOpenAiCompatible` |
| Device testing of dictation and layout | **Required** SC06, SC29, SC30 | S | H | Phone + laptop | Chrome/Android and iPad Safari |
| Report PDF from the approval snapshot (deductions, version) | Optional | S | M | — | F44 |
| Remove deprecated read-only overview endpoints | Optional | S | L | Confirm no external callers | F41 |
| Legacy studio dictation relay | Optional | M | L | Only if the Jewellery Sheets Studio stays in use | F30 legacy part |
| PDF text extraction for JEV (Drive OCR advanced service) | Optional | M | M | Extra scope | Lets scanned folios reach JEV as text |
| Per-teacher roles (marker vs approver) | Optional | M | M | School process | Extend `WebAccess` with a second allowlist |
| Batch approval review screen | Optional | M | M | — | Approve several reviewed students with the same blocking-flag rules |
