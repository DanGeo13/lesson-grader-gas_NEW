/* Reference rubric used by tests/reference-rubric.test.cjs and documented,
   with every expected value worked by hand, in docs/REFERENCE_RUBRIC.md. */
const rows = [
  // C01 Research — 10 marks (low weight)
  { criterion: 'Research', part: 'Folio', maxMarks: 10, band: 'A', description: 'Evaluates several sources and justifies choices' },
  { criterion: 'Research', part: 'Folio', maxMarks: 10, band: 'A', description: 'Links research to design decisions' },
  { criterion: 'Research', part: 'Folio', maxMarks: 10, band: 'B', description: 'Explains relevant sources' },
  { criterion: 'Research', part: 'Folio', maxMarks: 10, band: 'B', description: 'Some links to design decisions' },
  { criterion: 'Research', part: 'Folio', maxMarks: 10, band: 'C', description: 'Describes relevant sources' },
  { criterion: 'Research', part: 'Folio', maxMarks: 10, band: 'C', description: 'Records initial ideas' },
  { criterion: 'Research', part: 'Folio', maxMarks: 10, band: 'D', description: 'Lists a source' },
  { criterion: 'Research', part: 'Folio', maxMarks: 10, band: 'E', description: 'Attempts research' },
  // C02 Making — 30 marks (high weight)
  { criterion: 'Making', part: 'Practical', maxMarks: 30, band: 'A', description: 'Accurate measuring and marking out' },
  { criterion: 'Making', part: 'Practical', maxMarks: 30, band: 'A', description: 'Secure, clean joints' },
  { criterion: 'Making', part: 'Practical', maxMarks: 30, band: 'A', description: 'High-quality finish' },
  { criterion: 'Making', part: 'Practical', maxMarks: 30, band: 'C', description: 'Mostly accurate marking out' },
  { criterion: 'Making', part: 'Practical', maxMarks: 30, band: 'C', description: 'Functional joints' },
  { criterion: 'Making', part: 'Practical', maxMarks: 30, band: 'C', description: 'Acceptable finish' },
  { criterion: 'Making', part: 'Practical', maxMarks: 30, band: 'E', description: 'Attempts construction' },
  { criterion: 'Making', part: 'Practical', maxMarks: 30, band: 'E', description: 'Uses tools with assistance' }
];
const LEGACY_WEIGHTS = { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 };
const BANDS = [{ letter: 'A', min: 85, max: 100 }, { letter: 'B', min: 75, max: 84.99 }, { letter: 'C', min: 65, max: 74.99 }, { letter: 'D', min: 50, max: 64.99 }, { letter: 'E', min: 0, max: 49.99 }];
module.exports = { rows, LEGACY_WEIGHTS, BANDS };
