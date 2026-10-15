# Reference rubric and hand-worked results

This small rubric exercises every grading rule the system implements. Each case is calculated by hand below and
checked by `tests/reference-rubric.test.cjs` (cases R1–R9) and `tests/web-grading.test.cjs` (R10). The fixture is
`tests/fixtures/reference-rubric.cjs`.

**Nothing here is school policy.** Cases marked *legacy* use the weights currently in the repository; cases
marked *target* use the confirmed SC19 reference points with **illustrative, unconfirmed** B/D values.

## Rubric

| ID | Criterion | Marks | Band A | Band B | Band C | Band D | Band E |
|---|---|---|---|---|---|---|---|
| C01 | Research (Folio) | 10 | 2 observables | 2 | 2 | 1 | 1 |
| C02 | Making (Practical) | 30 | 3 observables | — | 3 | — | 2 |
| | **Total** | **40** | | | | | |

## Rules applied (as implemented)

1. An observable earns `criterion marks × band weight ÷ observables in that band` (`RubricEngine.gs` `scoreCriterion`).
2. A criterion's points are the sum over ticked observables, **capped** at the criterion's marks.
3. The distinct override awards full criterion marks.
4. Criterion letter and overall letter use the same boundaries: A ≥ 85, B ≥ 75, C ≥ 65, D ≥ 50, else E.
   An exact boundary belongs to the higher band. Percentages are compared unrounded (after removing
   floating-point noise below 10⁻⁹).
5. Classroom grade = final points ÷ rubric total × assignment max points, rounded half-up to 1 d.p.

Legacy weights: A 1.0, B 0.875, C 0.70, D 0.575, E 0.25.

## Cases

**R1 Unequal weights, partial achievement (legacy).** C01 ticks A-1: 10 × 1.0 ÷ 2 = **5.00** (50% → D).
C02 ticks A-1, A-2: 2 × 30 × 1.0 ÷ 3 = **20.00** (66.67% → C). Total 25 / 40 = **62.5% → D**.
The 30-mark criterion moves the total three times as much per percentage point as the 10-mark criterion.

**R2 A/C/E reference points (legacy).**
- All A: C01 = 10 (100%), C02 = 30 (100%) — matches the target.
- All C: C01 = 2 × 10 × 0.7 ÷ 2 = 7 (70%); C02 = 3 × 30 × 0.7 ÷ 3 = 21 (70%) — **target is 50%: conflict**.
- All E: C01 = 10 × 0.25 = 2.5 (25%); C02 = 2 × 30 × 0.25 ÷ 2 = 7.5 (25%) — **target is 0%: conflict**.

**R3 Target rule set (A 1.0, C 0.5, E 0.0; B 0.75 and D 0.25 illustrative only).**
All A = 40; all C = 5 + 15 = **20 (50%)**; all E = **0**. With the current overall boundaries a full-C portfolio
(50%) is graded **D**, so adopting the target reference points also requires confirming the overall boundaries
(see REVIEW.md §4, unresolved rule 4 "bottom of the grade band"). The engine refuses to build target weights
until B and D are supplied explicitly.

**R4 B/D behaviour (legacy, policy unconfirmed).** C01 both B observables: 10 × 0.875 = **8.75 (87.5%) → A**.
A criterion evidenced entirely at B therefore reads as an A letter. C01 D-1: 10 × 0.575 = **5.75 (57.5%) → D**.

**R5 Teacher overrides.** AI proposal: C01 B-1 = 10 × 0.875 ÷ 2 = 4.375; C02 all A = 30; total 34.375.
Teacher removes C02 A-3 and ticks C02 C-1 (seen in class): C02 = 20 + 30 × 0.7 ÷ 3 = **27**; total **31.375**.
Distinct override on C01 alone: **10** (full marks).

**R6 Cap.** C01 all A and all B: raw 10 + 8.75 = 18.75, capped to **10**; the raw value is kept for review.

**R7 Boundaries.** 100 A · 85.01 A · **85 A** · 84.99 B · **75 B** · 74.99 C · **65 C** · 64.99 D · **50 D** ·
49.99 E · 0 E. Floating-point sums such as ten C ticks of 0.7 (6.999999999999999) are treated as exactly 70%.

**R8 Rounding.** 25/40 on a 100-point assignment → **62.5**. 31.375/40 on a 20-point assignment →
15.6875 → **15.7**. Raw-points mode requires assignment total = rubric total. Display rounding never changes a
letter: 84.996% shows as "85.0%" but is a **B**.

**R9 Reproducibility.** Identical ticks with different feedback/notes give byte-identical score objects.

**R10 Reassessment without overwrite.** Approve v1 (snapshot row in `ApprovedGrades`), create reassessment v2,
grade and approve v2 differently: `ApprovedGrades` holds **both** rows; v1's row, ticks and history are unchanged
and v1 is no longer the current official version.
