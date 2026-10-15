/**
 * Legacy Sheets dialog workflows, including the original Year 9 Jewellery
 * criteria. Universal project/rubric grading is implemented separately in
 * RubricEngine.gs + WebGrading.gs; never use these hardcoded grading paths for
 * another task. See docs/IMPLEMENTATION.md before deploying changes.
 */

/* ----------------------------------------------------------------------------
 * BLOCK A — ADD new module: GradeScaleService.
 * Place this as a new top-level module, anywhere after Config and before
 * AssessmentService (e.g. directly above "var AssessmentService = ...").
 *
 * ASSUMED SETUP SHEET LAYOUT (please confirm/correct):
 *   A row containing the text "GRADE SCALE" marks the start of that table.
 *   Each following row has a letter (A-E) in one cell and a percentage
 *   number in another cell on the same row (e.g. "A" | 100%, or "A" | 100).
 *   The table ends at the first fully blank row after the header.
 *   The "OVERALL GRADE BANDS" table works the same way, but with a letter
 *   plus TWO numbers per row (min and max), e.g. "A" | 85 | 100.
 * ---------------------------------------------------------------------------- */
var GradeScaleService = (function() {
  var DEFAULT_WEIGHTS = { A: 1.0, B: 0.875, C: 0.70, D: 0.575, E: 0.25 };
  var DEFAULT_BANDS = [
    { letter: 'A', min: 85, max: 100 },
    { letter: 'B', min: 75, max: 84.99 },
    { letter: 'C', min: 65, max: 74.99 },
    { letter: 'D', min: 50, max: 64.99 },
    { letter: 'E', min: 0, max: 49.99 }
  ];

  function findHeaderRow_(data, keyword) {
    for (var r = 0; r < data.length; r++) {
      var rowStr = data[r].join(' ').toUpperCase();
      if (rowStr.indexOf(keyword) !== -1) return r;
    }
    return -1;
  }

  function extractLetterAndNumbers_(row) {
    var letter = null;
    var nums = [];
    for (var c = 0; c < row.length; c++) {
      var cell = row[c];
      if (letter === null && typeof cell === 'string' && /^[A-E]$/i.test(cell.trim())) {
        letter = cell.trim().toUpperCase();
        continue;
      }
      if (typeof cell === 'number') {
        nums.push(cell);
      } else if (typeof cell === 'string' && /^-?\d+(\.\d+)?%?$/.test(cell.trim())) {
        nums.push(parseFloat(cell.replace('%', '')));
      }
    }
    return { letter: letter, nums: nums };
  }

  function getGradeWeights() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_SETUP);
      if (!sheet) return DEFAULT_WEIGHTS;
      var data = sheet.getDataRange().getValues();
      var headerRow = findHeaderRow_(data, 'GRADE SCALE');
      if (headerRow === -1) return DEFAULT_WEIGHTS;

      var weights = {};
      for (var r = headerRow + 1; r < data.length; r++) {
        var rowJoined = data[r].join('').trim();
        if (!rowJoined) break;
        var parsed = extractLetterAndNumbers_(data[r]);
        if (parsed.letter && parsed.nums.length > 0) {
          var val = parsed.nums[0];
          weights[parsed.letter] = val > 1 ? val / 100 : val;
        }
        if (Object.keys(weights).length >= 5) break;
      }
      var complete = ['A', 'B', 'C', 'D', 'E'].every(function(l) { return weights[l] !== undefined; });
      return complete ? weights : DEFAULT_WEIGHTS;
    } catch (err) {
      Logging.logError('GradeScaleService.getGradeWeights', err);
      return DEFAULT_WEIGHTS;
    }
  }

  function getOverallGradeBands() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_SETUP);
      if (!sheet) return DEFAULT_BANDS;
      var data = sheet.getDataRange().getValues();
      var headerRow = findHeaderRow_(data, 'OVERALL GRADE BAND');
      if (headerRow === -1) return DEFAULT_BANDS;

      var bands = [];
      for (var r = headerRow + 1; r < data.length; r++) {
        var rowJoined = data[r].join('').trim();
        if (!rowJoined) break;
        var parsed = extractLetterAndNumbers_(data[r]);
        if (parsed.letter && parsed.nums.length >= 2) {
          bands.push({ letter: parsed.letter, min: parsed.nums[0], max: parsed.nums[1] });
        } else if (parsed.letter && parsed.nums.length === 1) {
          bands.push({ letter: parsed.letter, min: parsed.nums[0], max: 100 });
        }
        if (bands.length >= 5) break;
      }
      if (bands.length === 0) return DEFAULT_BANDS;
      bands.sort(function(a, b) { return b.min - a.min; });
      return bands;
    } catch (err) {
      Logging.logError('GradeScaleService.getOverallGradeBands', err);
      return DEFAULT_BANDS;
    }
  }

  /* Round-half-up letter selection: picks the nearest weight; an exact
     midpoint tie rounds UP to the higher letter. */
  function getOverallGradeLetter(totalPercent, bands) {
    bands = bands || getOverallGradeBands();
    for (var i = 0; i < bands.length; i++) {
      if (totalPercent >= bands[i].min) return bands[i].letter;
    }
    return bands.length ? bands[bands.length - 1].letter : 'E';
  }

  return {
    getGradeWeights: getGradeWeights,
    getOverallGradeBands: getOverallGradeBands,

    getOverallGradeLetter: getOverallGradeLetter
  };
})();



/* ============================================================================
 * SECTION 1: CONFIGURATION & CRITERIA DEFINITIONS
 * ============================================================================ */
var Config = (function() {
  var SHEET_SETUP = 'Setup';
  var SHEET_MARKING = '9DAT1 Marking';
  var SHEET_RUBRIC = 'Rubric';
  var SHEET_SUMMARY = 'Summary';

  var SHEET_CLASSROOM_CONFIG = 'ClassroomConfig';
  var SHEET_SUBMISSIONS = 'Submissions';
  var SHEET_SUBMISSION_FILES = 'SubmissionFiles';
  var SHEET_CRITERIA_CONFIG = 'CriteriaConfig';
  var SHEET_CRITERION_ASSESSMENTS = 'CriterionAssessments';
  var SHEET_AI_ASSESSMENTS = 'AIAssessments';
  var SHEET_ASSESSMENT_HISTORY = 'AssessmentHistory';
  var SHEET_ERROR_LOG = 'ErrorLog';

  var SCRIPT_PROP_GEMINI_KEY = 'GEMINI_API_KEY';
  var PRIMARY_MODEL = 'gemini-2.5-flash';
  var FALLBACK_MODELS = ['gemini-2.0-flash', 'gemini-1.5-flash'];
  var GEMINI_API_VERSION = 'v1beta';

  var GRADE_WEIGHTS = { 'A': 1.000, 'B': 0.875, 'C': 0.700, 'D': 0.575, 'E': 0.250 };
  var STATUS = {
    NEW: 'New', IN_REVIEW: 'InReview', AI_ASSESSED: 'AIAssessed',
    TEACHER_REVIEWED: 'TeacherReviewed', APPROVED: 'Approved',
    LOCKED: 'Locked', SUPERSEDED: 'Superseded', SYNC_FAILED: 'SyncFailed'
  };

  var CRITERIA_DEFINITIONS = [
    {
      criterionId: "C01", title: "Identification & Exploration of Need", part: "Part A", section: "Section 1: Project Proposal & Management", maxMarks: 5, outcome: "DT5-2",
      mdOverride: { id: "C01-MD", text: "Missing / Distinct: clear evidence beyond listed observables at an A standard." },
      bands: {
        E: [["C01-E-1", "Design brief present."], ["C01-E-2", "Identifies design opportunity/need."], ["C01-E-3", "At least two success criteria listed."], ["C01-E-4", "Understandable and on-task."]],
        D: [["C01-D-1", "At least three success criteria outlined."], ["C01-D-2", "Most criteria relevant to need."], ["C01-D-3", "Criteria specific enough to check."], ["C01-D-4", "Basic consideration of user/context."]],
        C: [["C01-C-1", "At least four sound success criteria outlined."], ["C01-C-2", "Specific enough to guide decisions."], ["C01-C-3", "Describes design opportunity with some detail."], ["C01-C-4", "Considers user and situation."]],
        B: [["C01-B-1", "At least five mostly relevant success criteria outlined."], ["C01-B-2", "Clear explanation of design opportunity."], ["C01-B-3", "Considers constraints (time/cost/safety)."], ["C01-B-4", "Proposed solution characteristics well-aligned."]],
        A: [["C01-A-1", "SIX highly relevant and accurate success criteria."], ["C01-A-2", "Comprehensive design brief explaining solution clearly."], ["C01-A-3", "Reflects deep understanding of designer considerations."], ["C01-A-4", "Addresses constraints and real wearer needs thoughtfully."]]
      }
    },
    {
      criterionId: "C02", title: "Inspiration (Annotated Mood Board)", part: "Part A", section: "Section 1: Project Proposal & Management", maxMarks: 10, outcome: "DT5-2",
      mdOverride: { id: "C02-MD", text: "Missing / Distinct: exceptionally curated thematic visual inquiry." },
      bands: {
        E: [["C02-E-1", "Mood board present."], ["C02-E-2", "3-4 jewellery images."], ["C02-E-3", "Basic relevance identifiable."], ["C02-E-4", "Minimal labelling."]],
        D: [["C02-D-1", "Basic inspiration collection (5+ sources)."], ["C02-D-2", "Exploration of jewellery styles/materials."], ["C02-D-3", "Brief annotations identify images."], ["C02-D-4", "Basic connection to initial ideas."]],
        C: [["C02-C-1", "Sound range of 6-8 diverse inspirations."], ["C02-C-2", "Annotations explain basic relevance & design influence."], ["C02-C-3", "Identifies specific design elements (colour palette, motifs)."], ["C02-C-4", "Relevant to theme and target market."]],
        B: [["C02-B-1", "Thorough range of 8+ well-curated inspirations."], ["C02-B-2", "Clear annotations explaining relevance and influence."], ["C02-B-3", "Analyses links to materials and wearer style."], ["C02-B-4", "Structured, cohesive emerging aesthetic direction."]],
        A: [["C02-A-1", "Extensive, highly diverse inspiration sources."], ["C02-A-2", "Insightful, detailed annotations explaining aesthetic influence."], ["C02-A-3", "Explicit links to planned components/techniques."], ["C02-A-4", "Exceptional visual presentation and creative vision."]]
      }
    },
    {
      criterionId: "C03", title: "Action Plan / Gantt Chart", part: "Part A", section: "Section 1: Project Proposal & Management", maxMarks: 5, outcome: "DT5-10",
      mdOverride: { id: "C03-MD", text: "Missing / Distinct: dynamic digital tracking with milestone adaptation." },
      bands: {
        E: [["C03-E-1", "Action plan/timeline present."], ["C03-E-2", "Lists 3-4 basic tasks."], ["C03-E-3", "Indication of dates or order."], ["C03-E-4", "Loosely related to jewellery project."]],
        D: [["C03-D-1", "Outlines 5 key tasks for Part A & B."], ["C03-D-2", "Estimated timeframes or weeks assigned."], ["C03-D-3", "Sequential progression shown."], ["C03-D-4", "Due dates identified."]],
        C: [["C03-C-1", "Action plan completed soundly with key details."], ["C03-C-2", "Differentiates research, design, and making."], ["C03-C-3", "Logical and realistic sequencing."], ["C03-C-4", "Notes resources and checkpoints."]],
        B: [["C03-B-1", "Action plan completed thoroughly across stages."], ["C03-B-2", "Clear durations and estimated vs actual tracking."], ["C03-B-3", "Reflects workshop and curing constraints."], ["C03-B-4", "Evidence of ongoing monitoring and feedback."]],
        A: [["C03-A-1", "Action plan completed extensively with wide details."], ["C03-A-2", "Precise milestones, dependencies, and WHS checkpoints."], ["C03-A-3", "Integrated evaluation checkpoints and contingency time."], ["C03-A-4", "Rigorous management of time and equipment."]]
      }
    },
    {
      criterionId: "C04", title: "Application of Experiments of Materials (2 Techniques)", part: "Part A", section: "Section 2: Project Development & Realisation", maxMarks: 10, outcome: "DT5-9",
      mdOverride: { id: "C04-MD", text: "Missing / Distinct: unusually sophisticated experimentation or innovative testing." },
      bands: {
        E: [["C04-E-1", "Evidence of 1 material/technique test piece."], ["C04-E-2", "Identifiable material or technique."], ["C04-E-3", "Basic photo/note documentation."], ["C04-E-4", "Loosely connected to jewellery task."]],
        D: [["C04-D-1", "One OR two techniques researched with basic samples."], ["C04-D-2", "Basic engagement rather than token attempt."], ["C04-D-3", "Simple reference samples included."], ["C04-D-4", "Basic reflection on what worked/failed."]],
        C: [["C04-C-1", "Two jewellery techniques soundly researched."], ["C04-C-2", "Reasonable quality reference samples included."], ["C04-C-3", "Original evidence of development of both techniques."], ["C04-C-4", "Comparison of technique suitability."]],
        B: [["C04-B-1", "Two techniques thoroughly researched with high quality samples."], ["C04-B-2", "Original evidence of development of both techniques."], ["C04-B-3", "Thoughtful parameter testing (tension, curing, joins)."], ["C04-B-4", "Explains influence on final design decisions."]],
        A: [["C04-A-1", "Two techniques extensively researched with outstanding samples."], ["C04-A-2", "Original multi-stage development evidence beyond basic class templates."], ["C04-A-3", "Detailed technical reasoning linking tests to final choices."], ["C04-A-4", "High-level understanding of material properties."]]
      }
    },
    {
      criterionId: "C05", title: "Description of Inspirational Australian Jeweller", part: "Part A", section: "Section 2: Project Development & Realisation", maxMarks: 5, outcome: "DT5-4",
      mdOverride: { id: "C05-MD", text: "Missing / Distinct: primary correspondence or critical contemporary market analysis." },
      bands: {
        E: [["C05-E-1", "Australian jeweller identified."], ["C05-E-2", "At least 1 image of work/biography."], ["C05-E-3", "Minimal biographical details."], ["C05-E-4", "Information related to jewellery."]],
        D: [["C05-D-1", "Basic outline of biographical information."], ["C05-D-2", "Covers a few details about their background."], ["C05-D-3", "Shows examples of distinctive work."], ["C05-D-4", "Basic interest/connection stated."]],
        C: [["C05-C-1", "Sound description covering key biographical details."], ["C05-C-2", "Describes style, philosophy, and materials."], ["C05-C-3", "Multiple images with descriptive captions."], ["C05-C-4", "Explains what inspires the student."]],
        B: [["C05-B-1", "Thorough description covering a range of details."], ["C05-B-2", "Detailed examination of signature techniques and aesthetic."], ["C05-B-3", "Analyses market trends and customer focus."], ["C05-B-4", "Links designer practice to student's work."]],
        A: [["C05-A-1", "Extensive biographical case study covering wide details."], ["C05-A-2", "Critical analysis of technical mastery and design identity."], ["C05-A-3", "Insightful parallels drawn to student's collection brief."], ["C05-A-4", "Synthesized in own words with proper attribution."]]
      }
    },
    {
      criterionId: "C06", title: "Explanations of Design Elements & Principles", part: "Part A", section: "Section 2: Project Development & Realisation", maxMarks: 10, outcome: "DT5-4",
      mdOverride: { id: "C06-MD", text: "Missing / Distinct: visual semiotic analysis or advanced design theory critique." },
      bands: {
        E: [["C06-E-1", "1-2 elements/principles named."], ["C06-E-2", "Minimal descriptive definitions."], ["C06-E-3", "At least one illustrative image."], ["C06-E-4", "Recognisable design terminology."]],
        D: [["C06-D-1", "Outlines and basic information regarding elements/principles."], ["C06-D-2", "Identifies 3 distinct elements/principles."], ["C06-D-3", "Connects to simple jewellery examples."], ["C06-D-4", "Reasonable accuracy of basic terms."]],
        C: [["C06-C-1", "Mostly descriptive explanations of design elements & principles."], ["C06-C-2", "Covers 4-5 relevant elements/principles."], ["C06-C-3", "Applies terminology accurately to jewellery designs."], ["C06-C-4", "Explains visual interest enhancements."]],
        B: [["C06-B-1", "Explanations and detailed information regarding elements/principles."], ["C06-B-2", "Analyses interaction of balance, contrast, and form."], ["C06-B-3", "Strong application to planned jewellery collection."], ["C06-B-4", "Consistently high standard of design language."]],
        A: [["C06-A-1", "Extensive explanations and detailed theoretical analysis."], ["C06-A-2", "Deep insight into balance, scale, and proportion."], ["C06-A-3", "Synthesizes theory with practical fabrication."], ["C06-A-4", "Professional annotations displaying mastery."]]
      }
    },
    {
      criterionId: "C07", title: "Analysis of Factor Affecting Design", part: "Part A", section: "Section 2: Project Development & Realisation", maxMarks: 10, outcome: "DT5-4",
      mdOverride: { id: "C07-MD", text: "Missing / Distinct: life-cycle assessment or comprehensive ergonomics inquiry." },
      bands: {
        E: [["C07-E-1", "Piece identified for analysis."], ["C07-E-2", "Mentions at least 1 factor (cost/function)."], ["C07-E-3", "Minimal comments on factor."], ["C07-E-4", "Loosely relevant design analysis."]],
        D: [["C07-D-1", "Basic outline of factors affecting one chosen piece."], ["C07-D-2", "Basic description of factor impacts."], ["C07-D-3", "Outlines user/situational requirements."], ["C07-D-4", "Simple cause-and-effect reasoning."]],
        C: [["C07-C-1", "Sound description of factors affecting one chosen piece."], ["C07-C-2", "Analyses 2-3 factors (function/ethics/cost/form)."], ["C07-C-3", "Considers wearability and durability."], ["C07-C-4", "Draws relevant conclusions for own project."]],
        B: [["C07-B-1", "Thorough explanation of factors affecting one chosen piece."], ["C07-B-2", "Detailed interplay of cost, ethics, manufacturing, and form."], ["C07-B-3", "Analyses how constraints were resolved."], ["C07-B-4", "Explicitly links findings to own parameters."]],
        A: [["C07-A-1", "Extensive analysis of chosen piece by designer of choice."], ["C07-A-2", "Critical evaluation of cultural, ergonomic, and environmental factors."], ["C07-A-3", "Sophisticated vocabulary regarding sustainability."], ["C07-A-4", "Deep synthesis guiding structural choices."]]
      }
    },
    {
      criterionId: "C08", title: "Evidence of Creativity (3 Sketches + Final Designs)", part: "Part B", section: "Section 2: Evidence of Creativity", maxMarks: 5, outcome: "DT5-2",
      mdOverride: { id: "C08-MD", text: "Missing / Distinct: highly original design development or innovative visual presentation." },
      bands: {
        E: [["C08-E-1", "Very few or incomplete initial sketches."], ["C08-E-2", "Final designs underdeveloped and unclear."], ["C08-E-3", "Minimal annotations present."], ["C08-E-4", "Loosely related to jewellery task."]],
        D: [["C08-D-1", "Limited initial sketches with minimal variation."], ["C08-D-2", "Final designs show limited refinement."], ["C08-D-3", "Basic presentation with minimal colour/notes."], ["C08-D-4", "Clear connection to jewellery brief."]],
        C: [["C08-C-1", "Clear initial sketches that meet the brief with some variation."], ["C08-C-2", "Final designs show basic development."], ["C08-C-3", "Neat presentation with some colour."], ["C08-C-4", "Simple annotations of materials/dimensions."]],
        B: [["C08-B-1", "Three relevant initial sketches showing clear ideas."], ["C08-B-2", "Final designs well-presented with appropriate rendering/colour."], ["C08-B-3", "Annotations explain key features and construction."], ["C08-B-4", "Visible development and purposeful creativity."]],
        A: [["C08-A-1", "Three highly creative, detailed initial sketches with strong variation."], ["C08-A-2", "Refined into three final designs to professional standard."], ["C08-A-3", "Accurate rendering of proportion, texture, and colour."], ["C08-A-4", "Detailed annotations explaining design decisions."]]
      }
    },
    {
      criterionId: "C09", title: "Order of Construction (Steps + Materials & Tools)", part: "Part B", section: "Section 2: Evidence of Creativity", maxMarks: 10, outcome: "DT5-10",
      mdOverride: { id: "C09-MD", text: "Missing / Distinct: production flowchart with risk matrix and quality checks." },
      bands: {
        E: [["C09-E-1", "Attempts an Order of Construction but unclear or incomplete."], ["C09-E-2", "Resource list minimal or missing essential items."], ["C09-E-3", "Mentions some tools/materials."], ["C09-E-4", "Broadly relates to making jewellery."]],
        D: [["C09-D-1", "Simple Order of Construction with limited sequencing."], ["C09-D-2", "Basic list of materials and tools with some relevant items."], ["C09-D-3", "Missing key resources or detail."], ["C09-D-4", "Basic safety notes included."]],
        C: [["C09-C-1", "Generally clear Order of Construction with logical sequencing."], ["C09-C-2", "Suitable list of materials and tools."], ["C09-C-3", "Steps describe specific actions and techniques."], ["C09-C-4", "Relevant workshop WHS precautions noted."]],
        B: [["C09-B-1", "Clear and logical Order of Construction with appropriate sequencing."], ["C09-B-2", "Detailed list of materials and tools mostly comprehensive."], ["C09-B-3", "Technical notes explaining exact joining/shaping steps."], ["C09-B-4", "WHS and PPE integrated into procedure."]],
        A: [["C09-A-1", "Extensively constructs logical, clear Order of Construction."], ["C09-A-2", "Well-sequenced method to complete jewellery piece."], ["C09-A-3", "Extensive, clearly documented list of materials and tools."], ["C09-A-4", "Comprehensive safety management embedded."]]
      }
    },
    {
      criterionId: "C10", title: "Final Evaluation", part: "Part B", section: "Section 3: Project Evaluation", maxMarks: 10, outcome: "DT5-2",
      mdOverride: { id: "C10-MD", text: "Missing / Distinct: authentic consumer market testing data or lifecycle analysis." },
      bands: {
        E: [["C10-E-1", "Limited information on final product range/process."], ["C10-E-2", "Minimal relation to design brief."], ["C10-E-3", "States basic personal satisfaction."], ["C10-E-4", "Relevant to finished project."]],
        D: [["C10-D-1", "Outlines final product range and process."], ["C10-D-2", "Relates to design brief and limitations at a basic level."], ["C10-D-3", "Identifies a challenge and simple modification."], ["C10-D-4", "Suggests basic future improvement."]],
        C: [["C10-C-1", "Describes final product range and process soundly."], ["C10-C-2", "Relates to design brief and limitations addressing key components."], ["C10-C-3", "Reflects on material/time difficulties."], ["C10-C-4", "Meaningful suggestions for improvement."]],
        B: [["C10-B-1", "Explains final product range and process thoroughly."], ["C10-B-2", "Relates to design brief and limitations addressing many components."], ["C10-B-3", "Critical reflection on aesthetic cohesion and function."], ["C10-B-4", "Constructive proposals for refinement."]],
        A: [["C10-A-1", "Critically evaluates final product range and process."], ["C10-A-2", "Relates comprehensively to design brief and limitations."], ["C10-A-3", "Deep reflection on material limitations and personal skills."], ["C10-A-4", "Insightful recommendations for scalable production."]]
      }
    },
    {
      criterionId: "C11", title: "Construction of Jewellery (Practical Skills)", part: "Part B", section: "Section 2: Practical Project", maxMarks: 10, outcome: "DT5-9",
      mdOverride: { id: "C11-MD", text: "Missing / Distinct: exceptionally refined finishing, complex joins, or innovative construction." },
      bands: {
        E: [["C11-E-1", "Limited evidence of uniformity and quality in techniques."], ["C11-E-2", "Piece basically intact."], ["C11-E-3", "Wearable/usable rather than falling apart."], ["C11-E-4", "Photos show basic construction."]],
        D: [["C11-D-1", "Basic evidence of uniformity and quality in required techniques."], ["C11-D-2", "Demonstrates basic skill development."], ["C11-D-3", "Sharp ends managed; item functions."], ["C11-D-4", "Intentionally made rather than accidental."]],
        C: [["C11-C-1", "Sound evidence of uniformity and quality in some techniques."], ["C11-C-2", "Demonstrates reasonable skill development."], ["C11-C-3", "Joins secure, components aligned, and sturdy."], ["C11-C-4", "Reasonable consistency in repeated links/beads."]],
        B: [["C11-B-1", "Thorough evidence of uniformity and quality in most techniques."], ["C11-B-2", "Demonstrates high standard skill development."], ["C11-B-3", "High-quality finish with clean joins and neat edges."], ["C11-B-4", "Careful material handling and minimal visible flaws."]],
        A: [["C11-A-1", "Extensive evidence of uniformity and quality in all required techniques."], ["C11-A-2", "Demonstrating outstanding skill development."], ["C11-A-3", "Consistently neat, secure, flawless joins."], ["C11-A-4", "Polished, professional surface finishing."]]
      }
    },
    {
      criterionId: "C12", title: "Jewellery Aesthetics (Design Quality & Range Cohesion)", part: "Part B", section: "Section 2: Practical Project", maxMarks: 10, outcome: "DT5-9",
      mdOverride: { id: "C12-MD", text: "Missing / Distinct: unified couture collection exhibiting gallery presentation." },
      bands: {
        E: [["C12-E-1", "Limited displays of jewellery design displaying minimal creativity."], ["C12-E-2", "At least 2 pieces share basic theme/colour."], ["C12-E-3", "Aesthetics discernible."], ["C12-E-4", "Identifiable personal adornment set."]],
        D: [["C12-D-1", "Basic displays of jewellery design displaying creativity to audience."], ["C12-D-2", "Some cohesion across pieces (similar beads/colours)."], ["C12-D-3", "Basic consideration of target tastes."], ["C12-D-4", "Neat presentation communicating style."]],
        C: [["C12-C-1", "Sound displays of jewellery design matching reasonably."], ["C12-C-2", "Displays reasonable level of creativity and appropriateness."], ["C12-C-3", "Materials and equipment used are suitable."], ["C12-C-4", "Clear thematic links across 3 pieces."]],
        B: [["C12-B-1", "High level displays where pieces mostly match perfectly."], ["C12-B-2", "Displays thorough levels of creativity and appropriateness."], ["C12-B-3", "Very unique pieces suitable for purpose."], ["C12-B-4", "High presentation quality and styled photography."]],
        A: [["C12-A-1", "Outstanding displays where pieces match perfectly and are very unique."], ["C12-A-2", "Extensive level of creativity and appropriateness to target audience."], ["C12-A-3", "Materials and equipment very suitable for design purpose."], ["C12-A-4", "Exhibition quality presentation and brand identity."]]
      }
    }
  ];

  function getSerializedRubricPrompt() {
    var out = [];
    for (var i = 0; i < CRITERIA_DEFINITIONS.length; i++) {
      var c = CRITERIA_DEFINITIONS[i];
      var s = "CRITERION " + c.criterionId + ": " + c.title + " (" + c.part + ", /" + c.maxMarks + " Marks)\n";
      var bands = ['A', 'B', 'C', 'D', 'E'];
      for (var b = 0; b < bands.length; b++) {
        var bn = bands[b];
        var items = c.bands[bn] || [];
        var itemArr = [];
        for (var k = 0; k < items.length; k++) {
          itemArr.push("[" + items[k][0] + "] " + items[k][1]);
        }
        s += "  - Band " + bn + ": " + itemArr.join("; ") + "\n";
      }
      out.push(s);
    }
    return out.join("\n");
  }

  return {
    /* Modern (underscore) names — used throughout this codebase */
    SHEET_SETUP: SHEET_SETUP, SHEET_MARKING: SHEET_MARKING, SHEET_RUBRIC: SHEET_RUBRIC, SHEET_SUMMARY: SHEET_SUMMARY,
    SHEET_CLASSROOM_CONFIG: SHEET_CLASSROOM_CONFIG, SHEET_SUBMISSIONS: SHEET_SUBMISSIONS, SHEET_SUBMISSION_FILES: SHEET_SUBMISSION_FILES,
    SHEET_CRITERIA_CONFIG: SHEET_CRITERIA_CONFIG, SHEET_CRITERION_ASSESSMENTS: SHEET_CRITERION_ASSESSMENTS,
    SHEET_AI_ASSESSMENTS: SHEET_AI_ASSESSMENTS, SHEET_ASSESSMENT_HISTORY: SHEET_ASSESSMENT_HISTORY, SHEET_ERROR_LOG: SHEET_ERROR_LOG,
    SCRIPT_PROP_GEMINI_KEY: SCRIPT_PROP_GEMINI_KEY, PRIMARY_MODEL: PRIMARY_MODEL, FALLBACK_MODELS: FALLBACK_MODELS,
    GEMINI_API_VERSION: GEMINI_API_VERSION, GRADE_WEIGHTS: GRADE_WEIGHTS, STATUS: STATUS, CRITERIA_DEFINITIONS: CRITERIA_DEFINITIONS,

    /* Legacy (no-underscore) aliases — kept so any old references never see "undefined" again */
    SHEETSETUP: SHEET_SETUP, SHEETMARKING: SHEET_MARKING, SHEETRUBRIC: SHEET_RUBRIC, SHEETSUMMARY: SHEET_SUMMARY,
    SHEETCLASSROOMCONFIG: SHEET_CLASSROOM_CONFIG, SHEETSUBMISSIONS: SHEET_SUBMISSIONS, SHEETSUBMISSIONFILES: SHEET_SUBMISSION_FILES,
    SHEETCRITERIACONFIG: SHEET_CRITERIA_CONFIG, SHEETCRITERIONASSESSMENTS: SHEET_CRITERION_ASSESSMENTS,
    SHEETAIASSESSMENTS: SHEET_AI_ASSESSMENTS, SHEETASSESSMENTHISTORY: SHEET_ASSESSMENT_HISTORY, SHEETERRORLOG: SHEET_ERROR_LOG,
    SCRIPTPROPGEMINIKEY: SCRIPT_PROP_GEMINI_KEY, PRIMARYMODEL: PRIMARY_MODEL, FALLBACKMODELS: FALLBACK_MODELS,
    GEMINIAPIVERSION: GEMINI_API_VERSION, GRADEWEIGHTS: GRADE_WEIGHTS, CRITERIADEFINITIONS: CRITERIA_DEFINITIONS,

    getSerializedRubricPrompt: getSerializedRubricPrompt
  };
})();

/* ============================================================================
 * SECTION 2: UTILITIES & AUDIT LOGGING (SAFE USER RESOLUTION)
 * ============================================================================ */
var Utils = (function() {
  function generateUuid() { return Utilities.getUuid(); }
  function formatDate(d) {
    if (!d) return '';
    if (typeof d === 'string') return d;
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Australia/Sydney', "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  function safeJsonParse(str, fallback) {
    if (!str) return fallback;
    if (typeof str === 'object') return str;
    try { return JSON.parse(str); } catch (e) { return fallback; }
  }
  function safeJsonStringify(obj) {
    try { return JSON.stringify(obj); } catch (e) { return '{}'; }
  }
  function sanitizeError(err) {
    if (!err) return 'Unknown error';
    var msg = err.message || String(err);
    return msg.replace(/key=[a-zA-Z0-9_-]+/gi, 'key=REDACTED');
  }
  function getSafeUserEmail() {
    try {
      return Session.getActiveUser().getEmail() || 'Teacher';
    } catch (e) {
      return 'Teacher';
    }
  }
  /* Feedback text -> Sheets rich text. Only **bold** is interpreted; every
     other character is written literally. This helper was referenced by the
     markbook sync but missing, which half-applied legacy approvals (F08). */
  function buildRichTextFromMarkdown(markdown) {
    var src = String(markdown == null ? '' : markdown);
    var plain = '', spans = [], re = /\*\*([^*]+)\*\*/g, last = 0, m;
    while ((m = re.exec(src))) {
      plain += src.slice(last, m.index);
      var start = plain.length;
      plain += m[1];
      spans.push([start, plain.length]);
      last = re.lastIndex;
    }
    plain += src.slice(last);
    var builder = SpreadsheetApp.newRichTextValue().setText(plain);
    if (spans.length) {
      var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
      spans.forEach(function (r) { if (r[1] > r[0]) builder.setTextStyle(r[0], r[1], bold); });
    }
    return builder.build();
  }
  return {
    buildRichTextFromMarkdown: buildRichTextFromMarkdown,
    generateUuid: generateUuid,
    formatDate: formatDate,
    safeJsonParse: safeJsonParse,
    safeJsonStringify: safeJsonStringify,
    sanitizeError: sanitizeError,
    getSafeUserEmail: getSafeUserEmail
  };
})();

var Logging = (function() {
  function logHistory(submissionRecordId, assessmentId, action, criterionId, previousState, newState, notes) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_ASSESSMENT_HISTORY);
      if (!sheet) return;
      sheet.appendRow([
        Utils.generateUuid(), submissionRecordId || '', assessmentId || '', action || '', criterionId || '',
        Utils.safeJsonStringify(previousState || {}), Utils.safeJsonStringify(newState || {}),
        Utils.getSafeUserEmail(), Utils.formatDate(new Date()), notes || ''
      ]);
    } catch (e) { console.error('History log error: ' + e.message); }
  }

  function logError(context, errorObj, metadata) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_ERROR_LOG);
      if (!sheet) return;
      sheet.appendRow([
        Utils.generateUuid(), Utils.formatDate(new Date()), context || 'General',
        Utils.sanitizeError(errorObj), (errorObj && errorObj.stack) ? Utils.sanitizeError(errorObj.stack) : '',
        Utils.safeJsonStringify(metadata || {}), Utils.getSafeUserEmail(), 'Logged'
      ]);
    } catch (e) { console.error('ErrorLog write failed: ' + e.message); }
  }
  return { logHistory: logHistory, logError: logError };
})();

/* ============================================================================
 * SECTION 3: SHEET MANAGEMENT & FORMULAS
 * ============================================================================ */
var Sheets = (function() {
  function setupOrMigrateAssessmentSystem() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Backup step
    try {
      var driveFile = DriveApp.getFileById(ss.getId());
      var timeString = Utils.formatDate(new Date()).replace(/[:\/\\]/g, '-');
      driveFile.makeCopy("[Backup " + timeString + "] " + ss.getName());
    } catch(e) {
      Logging.logError('setupOrMigrateAssessmentSystem', e);
      return { success: false, message: 'Backup failed. No migration was attempted. Check Drive permissions before retrying.' };
    }

    var schemas = [
      { name: 'RubricProfiles', headers: ['ProfileID', 'ProfileName', 'JSONDefinition', 'CreatedAt'] },
      { name: 'ClassLists', headers: ['ClassID', 'OfficialName', 'PreferredName', 'SchoolEmail', 'Active'] },
      { name: Config.SHEET_CLASSROOM_CONFIG, headers: ['ConfigID', 'CourseID', 'CourseName', 'CourseSection', 'CourseWorkID', 'AssignmentTitle', 'MaxPoints', 'SavedAt', 'SavedBy', 'Active', 'TaskName', 'DueDate', 'AutoImported'] },
      { name: Config.SHEET_SUBMISSIONS, headers: ['SubmissionRecordID', 'ClassroomCourseID', 'ClassroomCourseWorkID', 'ClassroomSubmissionID', 'StudentUserID', 'StudentName', 'StudentEmail', 'Class', 'Task', 'SubmissionVersion', 'SourceType', 'ClassroomState', 'TurnedInTime', 'UpdateTime', 'Late', 'AttachmentSummary', 'AttachmentFileIDsJSON', 'AttachmentMetadataJSON', 'DriveFolderID', 'Status', 'ParentSubmissionRecordID', 'CurrentOfficial', 'LockedAt', 'LockedBy', 'ApprovedAt', 'ApprovedBy', 'ClassroomAssignedGrade', 'LastClassroomSyncAt', 'LastSyncResult', 'Notes'] },
      { name: Config.SHEET_SUBMISSION_FILES, headers: ['SubmissionRecordID', 'FileRecordID', 'SourceType', 'DriveFileID', 'FileName', 'MimeType', 'AlternateLink', 'ThumbnailUrl', 'FileSize', 'EligibleForAI', 'AIReviewStatus', 'AIExtractedText', 'Limitations', 'CreatedAt'] },
      { name: Config.SHEET_CRITERIA_CONFIG, headers: ['CriterionID', 'CriterionTitle', 'Part', 'Section', 'MaxMarks', 'Outcome', 'Band', 'ThresholdType', 'RequiredCount', 'CheckboxID', 'CheckboxLabel', 'EvidenceFocus', 'IsMissingDistinctOverride'] },
      { name: Config.SHEET_CRITERION_ASSESSMENTS, headers: ['AssessmentID', 'SubmissionRecordID', 'CriterionID', 'AIProposedGrade', 'DerivedGrade', 'FinalApprovedGrade', 'DistinctOverrideSelected', 'CheckboxesJSON', 'AIEvidenceJSON', 'TeacherAdjusted', 'TeacherAudioTranscript', 'TeacherWrittenNote', 'Confidence', 'MinimumEvidenceIncomplete', 'Status', 'CreatedAt', 'UpdatedAt', 'ApprovedAt', 'ApprovedBy'] },
      { name: Config.SHEET_AI_ASSESSMENTS, headers: ['AssessmentID', 'SubmissionRecordID', 'RunType', 'RunTimestamp', 'ModelUsed', 'PromptVersion', 'InputFilesJSON', 'PromptSummary', 'ResponseJSON', 'ExecutionStatus', 'ErrorMessage', 'LatencySeconds', 'TriggeredBy', 'TeacherOutcome'] },
      { name: Config.SHEET_ASSESSMENT_HISTORY, headers: ['HistoryID', 'SubmissionRecordID', 'AssessmentID', 'Action', 'CriterionID', 'PreviousStateJSON', 'NewStateJSON', 'ActorEmail', 'Timestamp', 'Notes'] },
      { name: Config.SHEET_ERROR_LOG, headers: ['ErrorID', 'Timestamp', 'Context', 'ErrorMessage', 'StackTrace', 'MetadataJSON', 'UserEmail', 'ResolutionStatus'] }
    ];

    for (var i = 0; i < schemas.length; i++) {
      var sDef = schemas[i];
      var sheet = ss.getSheetByName(sDef.name);
      if (!sheet) sheet = ss.insertSheet(sDef.name);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(sDef.headers);
        var headerRange = sheet.getRange(1, 1, 1, sDef.headers.length);
        headerRange.setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }
    }

    var classroomConfigSheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG);
    if (classroomConfigSheet) {
      var columns = ['TaskName', 'DueDate', 'AutoImported'];
      var lastColumn = classroomConfigSheet.getLastColumn();
      var headings = lastColumn ? classroomConfigSheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
      var missing = columns.filter(function(name) { return headings.indexOf(name) === -1; });
      if (missing.length) {
        var firstNewColumn = headings.length + 1;
        var targetLastColumn = firstNewColumn + missing.length - 1;
        if (classroomConfigSheet.getMaxColumns() < targetLastColumn) {
          classroomConfigSheet.insertColumnsAfter(classroomConfigSheet.getMaxColumns(), targetLastColumn - classroomConfigSheet.getMaxColumns());
        }
        classroomConfigSheet.getRange(1, firstNewColumn, 1, missing.length).setValues([missing]);
      }
    }

    populateCriteriaConfig(ss);
    verifyMarkingSheetFormulas(ss);
    return { success: true, message: 'All tables and configurations verified.' };
  }

  function populateCriteriaConfig(ss) {
    var cSheet = ss.getSheetByName(Config.SHEET_CRITERIA_CONFIG);
    if (!cSheet || cSheet.getLastRow() > 1) return;
    var defs = Config.CRITERIA_DEFINITIONS;
    var rows = [];

    for (var i = 0; i < defs.length; i++) {
      var c = defs[i];
      rows.push([c.criterionId, c.title, c.part, c.section, c.maxMarks, c.outcome, 'A', 'Override', 1, c.mdOverride.id, c.mdOverride.text, 'Distinct Evidence', true]);
      var bands = ['A', 'B', 'C', 'D', 'E'];
      for (var b = 0; b < bands.length; b++) {
        var bName = bands[b];
        var items = c.bands[bName] || [];
        var req = (bName === 'A') ? items.length : Math.ceil(items.length * 2 / 3);
        for (var k = 0; k < items.length; k++) {
          rows.push([c.criterionId, c.title, c.part, c.section, c.maxMarks, c.outcome, bName, (bName === 'A' ? 'All' : 'TwoThirds'), req, items[k][0], items[k][1], 'Observable', false]);
        }
      }
    }
    if (rows.length > 0) cSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  function verifyMarkingSheetFormulas(ss) {
    var sheet = ss.getSheetByName(Config.SHEET_MARKING);
    if (!sheet) return;
    var numRows = sheet.getLastRow();
    if (numRows < 2) return;

    var W = GradeScaleService.getGradeWeights();
    var bands = GradeScaleService.getOverallGradeBands();
    var defs = Config.CRITERIA_DEFINITIONS;
    var colLetters = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];

    function ifChain_(col, maxMarks, r) {
      var wA = Math.round(maxMarks * W.A * 1000) / 1000;
      var wB = Math.round(maxMarks * W.B * 1000) / 1000;
      var wC = Math.round(maxMarks * W.C * 1000) / 1000;
      var wD = Math.round(maxMarks * W.D * 1000) / 1000;
      var wE = Math.round(maxMarks * W.E * 1000) / 1000;
      return 'IF(' + col + r + '="A",' + wA +
        ',IF(' + col + r + '="B",' + wB +
        ',IF(' + col + r + '="C",' + wC +
        ',IF(' + col + r + '="D",' + wD + ',' + wE + '))))';
    }

    var ifsArgs = [];
    for (var bi = 0; bi < bands.length; bi++) {
      if (bi === bands.length - 1) {
        ifsArgs.push('TRUE,"' + bands[bi].letter + '"');
      } else {
        ifsArgs.push('{R}>=' + bands[bi].min + ',"' + bands[bi].letter + '"');
      }
    }

    for (var r = 2; r <= numRows; r++) {
      if (!sheet.getRange(r, 2).getValue()) continue;

      var partATerms = [];
      var partBTerms = [];
      for (var c = 0; c < colLetters.length; c++) {
        var def = defs[c];
        if (!def) continue;
        var term = ifChain_(colLetters[c], def.maxMarks, r);
        if (def.part === 'Part A') partATerms.push(term);
        else partBTerms.push(term);
      }

      sheet.getRange(r, 15).setFormula('=SUM(' + partATerms.join(',') + ')');
      sheet.getRange(r, 16).setFormula('=SUM(' + partBTerms.join(',') + ')');
      sheet.getRange(r, 17).setFormula('=O' + r + '+P' + r);

      var rowIfs = 'IFS(' + ifsArgs.join(',').replace(/\{R\}/g, 'Q' + r) + ')';
      sheet.getRange(r, 18).setFormula('=' + rowIfs);
    }
  }

  function syncApprovedGradesToMarkingSheet(studentName, gradesObj, feedbackObj) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(Config.SHEET_MARKING);
    if (!sheet) throw new Error('Sheet ' + Config.SHEET_MARKING + ' not found.');

    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (data[r][1] && String(data[r][1]).trim().toLowerCase() === String(studentName).trim().toLowerCase()) {
        targetRow = r + 1;
        break;
      }
    }
    if (targetRow === -1) {
      targetRow = sheet.getLastRow() + 1;
      sheet.getRange(targetRow, 2).setValue(studentName);
    }

    var crits = ['C01', 'C02', 'C03', 'C04', 'C05', 'C06', 'C07', 'C08', 'C09', 'C10', 'C11', 'C12'];
    var rowGrades = [];
    for (var i = 0; i < crits.length; i++) rowGrades.push(gradesObj[crits[i]] || 'E');
    /* Columns C-N (3-14): the 12 letter grades. Unchanged, already correct. */
    sheet.getRange(targetRow, 3, 1, 12).setValues([rowGrades]);

    if (feedbackObj) {
      /* Columns AE, AF, AG = 31, 32, 33. This was previously 19, 20, 21
         (S, T, U), which collided with pre-existing numeric-mark formulas
         for criteria C05-C07 and corrupted them. */
      if (feedbackObj.whatWentWell) {
        sheet.getRange(targetRow, 31).setRichTextValue(Utils.buildRichTextFromMarkdown(feedbackObj.whatWentWell));
      }
      if (feedbackObj.areasForImprovement) {
        sheet.getRange(targetRow, 32).setRichTextValue(Utils.buildRichTextFromMarkdown(feedbackObj.areasForImprovement));
      }
      if (feedbackObj.goalsForNextAssessment) {
        sheet.getRange(targetRow, 33).setRichTextValue(Utils.buildRichTextFromMarkdown(feedbackObj.goalsForNextAssessment));
      }
    }
    /* verifyMarkingSheetFormulas only touches columns O, P, Q, R (15-18),
       which are already correct on your sheet — safe to leave as-is. */
    verifyMarkingSheetFormulas(ss);
  }


  function getSheetDataAsObjects(sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var results = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var obj = {};
      for (var c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
      results.push(obj);
    }
    return results;
  }

  /* Header-name -> 0-based column index map. All lifecycle writes must go
     through these helpers so a column can never be written positionally
     into the wrong field again (see docs/AUDIT.md R7). */
  function getHeaderIndex(sheet) {
    var map = {};
    if (!sheet || sheet.getLastRow() < 1) return map;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var c = 0; c < headers.length; c++) {
      var name = String(headers[c] || '').trim();
      if (name && map[name] === undefined) map[name] = c;
    }
    return map;
  }

  function setCellByHeader(sheet, rowNumber, headerName, value) {
    var map = getHeaderIndex(sheet);
    var idx = map[headerName];
    if (idx === undefined) throw new Error('Missing required column: ' + headerName);
    sheet.getRange(rowNumber, idx + 1).setValue(value);
  }

  function findRowByKey(sheet, headerName, keyValue) {
    var map = getHeaderIndex(sheet);
    var idx = map[headerName];
    if (idx === undefined || sheet.getLastRow() < 2) return -1;
    var values = sheet.getRange(2, idx + 1, sheet.getLastRow() - 1, 1).getValues();
    for (var r = 0; r < values.length; r++) {
      if (String(values[r][0]) === String(keyValue)) return r + 2;
    }
    return -1;
  }

  return {
    setupOrMigrateAssessmentSystem: setupOrMigrateAssessmentSystem,
    syncApprovedGradesToMarkingSheet: syncApprovedGradesToMarkingSheet,
    getSheetDataAsObjects: getSheetDataAsObjects,
    verifyMarkingSheetFormulas: verifyMarkingSheetFormulas,
    getHeaderIndex: getHeaderIndex,
    setCellByHeader: setCellByHeader,
    findRowByKey: findRowByKey
  };
})();

/* ============================================================================
 * SECTION 4: ASSESSMENT ENGINE & DETERMINISTIC INDEPENDENT GRADES
 * ============================================================================ */
/* ----------------------------------------------------------------------------
 * RosterService
 * ---------------------------------------------------------------------------- */
var RosterService = (function() {
  function getClassRosterMap(targetClassId) {
    var map = {};
    var classLists = Sheets.getSheetDataAsObjects("ClassLists");
    for (var i = 0; i < classLists.length; i++) {
      var row = classLists[i];
      if (row.Active !== true && row.Active !== "TRUE" && row.Active !== "true") continue;
      if (targetClassId && String(row.ClassID).toLowerCase().trim() !== String(targetClassId).toLowerCase().trim()) continue;
      if (!row.SchoolEmail) continue;
      var email = String(row.SchoolEmail).toLowerCase().trim();
      map[email] = {
        officialName: row.OfficialName || "",
        preferredName: row.PreferredName || "",
        classId: row.ClassID || ""
      };
    }
    return map;
  }

  function resolveStudentIdentity(email, officialNameFromClassroom, rosterMap) {
    if (!email) email = "";
    var normEmail = String(email).toLowerCase().trim();
    if (rosterMap && rosterMap[normEmail]) {
      var entry = rosterMap[normEmail];
      var disp = entry.preferredName || entry.officialName || officialNameFromClassroom;
      return {
        displayName: disp,
        officialName: entry.officialName,
        preferredName: entry.preferredName,
        matched: true
      };
    } else {
      return {
        displayName: officialNameFromClassroom,
        officialName: officialNameFromClassroom,
        preferredName: (officialNameFromClassroom || "").split(" ")[0],
        matched: false
      };
    }
  }

  return {
    getClassRosterMap: getClassRosterMap,
    resolveStudentIdentity: resolveStudentIdentity
  };
})();

/* ----------------------------------------------------------------------------
 * RubricsLibrary (Phase 3 Foundation)
 * ---------------------------------------------------------------------------- */
var RubricsLibrary = (function() {
  var ROOT_FOLDER_NAME = 'Graded Assessments';

  function getOrCreateFolder(parent, name) {
    var folders = parent.getFoldersByName(name);
    if (folders.hasNext()) return folders.next();
    return parent.createFolder(name);
  }

  function ensureRubricsLibraryStructure() {
    var root = DriveApp.getRootFolder();
    var libraryFolder = getOrCreateFolder(root, ROOT_FOLDER_NAME);
    var settingsFolder = getOrCreateFolder(libraryFolder, "Settings");
    getOrCreateFolder(libraryFolder, "Shared Rubrics");
    getOrCreateFolder(libraryFolder, "Stage 4");
    getOrCreateFolder(libraryFolder, "Stage 5");
    getOrCreateFolder(libraryFolder, "Stage 6");

    var files = settingsFolder.getFilesByName("Rubrics Index");
    var ss;
    if (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        ss = SpreadsheetApp.openById(file.getId());
      }
    }
    if (!ss) {
      ss = SpreadsheetApp.create("Rubrics Index");
      var file = DriveApp.getFileById(ss.getId());
      file.moveTo(settingsFolder);
    }

    var schemas = [
      { name: "Courses", headers: ["Stage", "CourseCode", "Year", "Active"] },
      { name: "Categories", headers: ["CourseCode", "CategoryName"] },
      { name: "Tasks", headers: ["CourseCode", "CategoryName", "TaskName", "RubricFileId"] }
    ];

    for (var i = 0; i < schemas.length; i++) {
      var sDef = schemas[i];
      var sheet = ss.getSheetByName(sDef.name);
      if (!sheet) {
        sheet = ss.insertSheet(sDef.name);
        sheet.appendRow(sDef.headers);
        sheet.getRange(1, 1, 1, sDef.headers.length).setFontWeight("bold");
      }
    }
    // Clean up default "Sheet1" if it exists and is not needed
    var sheet1 = ss.getSheetByName("Sheet1");
    if (sheet1 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);
    return { success: true, message: "Rubrics Library structure verified." };
  }

  function getIndexSpreadsheet() {
    var root = DriveApp.getRootFolder();
    var libraryFolder = getOrCreateFolder(root, ROOT_FOLDER_NAME);
    var settingsFolder = getOrCreateFolder(libraryFolder, "Settings");
    var files = settingsFolder.getFilesByName("Rubrics Index");
    if (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) return SpreadsheetApp.openById(file.getId());
    }
    return null;
  }

  function addCourseToIndex(stage, courseCode, year) {
    var ss = getIndexSpreadsheet();
    if (!ss) return false;
    var sheet = ss.getSheetByName("Courses");
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(stage) && String(data[r][1]) === String(courseCode) && String(data[r][2]) === String(year)) {
        return false; // Already exists
      }
    }
    sheet.appendRow([stage, courseCode, year, true]);
    return true;
  }

  function addCategoryToIndex(courseCode, categoryName) {
    var ss = getIndexSpreadsheet();
    if (!ss) return false;
    var sheet = ss.getSheetByName("Categories");
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(courseCode) && String(data[r][1]) === String(categoryName)) {
        return false;
      }
    }
    sheet.appendRow([courseCode, categoryName]);
    return true;
  }

  function addTaskToIndex(courseCode, categoryName, taskName, rubricFileId) {
    var ss = getIndexSpreadsheet();
    if (!ss) return false;
    var sheet = ss.getSheetByName("Tasks");
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(courseCode) && String(data[r][1]) === String(categoryName) && String(data[r][2]) === String(taskName)) {
        // Task already exists. If RubricFileId is empty/blank, update it and return true.
        // Otherwise, it is immutable, do nothing and return false.
        var existingRubricId = data[r][3] || "";
        if (existingRubricId === "" && rubricFileId) {
          sheet.getRange(r + 1, 4).setValue(rubricFileId);
          return true;
        }
        return false;
      }
    }
    sheet.appendRow([courseCode, categoryName, taskName, rubricFileId || ""]);
    return true;
  }

  function getActiveCourses(stage) {
    var ss = getIndexSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Courses");
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var r = 1; r < data.length; r++) {
      var isActive = data[r][3] === true || data[r][3] === "TRUE" || data[r][3] === "true";
      if (String(data[r][0]) === String(stage) && isActive) {
        results.push({ stage: data[r][0], courseCode: data[r][1], year: data[r][2], active: data[r][3] });
      }
    }
    return results;
  }

  function getCategoriesForCourse(courseCode) {
    var ss = getIndexSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Categories");
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(courseCode)) {
        results.push({ courseCode: data[r][0], categoryName: data[r][1] });
      }
    }
    return results;
  }

  function getTasksForCourseCategory(courseCode, categoryName) {
    var ss = getIndexSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Tasks");
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(courseCode) && String(data[r][1]) === String(categoryName)) {
        var rId = data[r][3] || "";
        var rName = "";
        if (rId) {
          try {
            rName = DriveApp.getFileById(rId).getName();
          } catch(e) {
            rName = "Unknown/Inaccessible File";
          }
        }
        results.push({ courseCode: data[r][0], categoryName: data[r][1], taskName: data[r][2], rubricFileId: rId, rubricName: rName });
      }
    }
    return results;
  }

  function ensureCourseCategoryFolderPath(stage, courseCode, categoryName, year) {
    if (stage !== 'Stage 4' && stage !== 'Stage 5' && stage !== 'Stage 6') {
      return { success: false, message: 'Invalid stage.' };
    }

    ensureRubricsLibraryStructure();

    var root = DriveApp.getRootFolder();
    var libraryFolder = getOrCreateFolder(root, ROOT_FOLDER_NAME);
    var stageFolder = getOrCreateFolder(libraryFolder, stage);
    var courseFolder = getOrCreateFolder(stageFolder, courseCode);
    var categoryFolder = getOrCreateFolder(courseFolder, categoryName);

    addCourseToIndex(stage, courseCode, year);
    addCategoryToIndex(courseCode, categoryName);

    return {
      success: true,
      categoryFolder: { id: categoryFolder.getId(), name: categoryFolder.getName() },
      message: 'Folder structure and index updated successfully.'
    };
  }

  function ensureGradingWorkbook(stage, courseCode, categoryName, year, taskName) {
    var folderResult = ensureCourseCategoryFolderPath(stage, courseCode, categoryName, year);
    if (!folderResult.success) {
      return folderResult;
    }

    var categoryFolder = DriveApp.getFolderById(folderResult.categoryFolder.id);

    // Required filename format:
    // Example: 7TECHI [2026] - Rube Goldberg Task - Grading
    var workbookName = String(courseCode).trim() +
      ' [' + String(year).trim() + '] - ' +
      String(taskName).trim() +
      ' - Grading';

    // Reuse the already-created workbook for this exact course/year/task.
    var files = categoryFolder.getFilesByName(workbookName);
    if (files.hasNext()) {
      var existingFile = files.next();
      return {
        success: true,
        workbook: { id: existingFile.getId(), name: existingFile.getName() },
        message: 'Found existing grading workbook: ' + workbookName
      };
    }

    var ss = SpreadsheetApp.create(workbookName);
    var newFile = DriveApp.getFileById(ss.getId());
    newFile.moveTo(categoryFolder);

    return {
      success: true,
      workbook: { id: newFile.getId(), name: newFile.getName() },
      message: 'Created grading workbook: ' + workbookName
    };
  }

  return {
    ROOT_FOLDER_NAME: ROOT_FOLDER_NAME,
    ensureRubricsLibraryStructure: ensureRubricsLibraryStructure,
    ensureCourseCategoryFolderPath: ensureCourseCategoryFolderPath,
    ensureGradingWorkbook: ensureGradingWorkbook,
    addCourseToIndex: addCourseToIndex,
    addCategoryToIndex: addCategoryToIndex,
    addTaskToIndex: addTaskToIndex,
    getActiveCourses: getActiveCourses,
    getCategoriesForCourse: getCategoriesForCourse,
    getTasksForCourseCategory: getTasksForCourseCategory
  };
})();

var AssessmentService = (function() {

  function deriveGradeFromChecks(criterionId, checksMap) {
    var defs = Config.CRITERIA_DEFINITIONS;
    var crit = null;
    for (var i = 0; i < defs.length; i++) {
      if (defs[i].criterionId === criterionId) { crit = defs[i]; break; }
    }
    var W = GradeScaleService.getGradeWeights();
    var bands = GradeScaleService.getOverallGradeBands();

    if (!crit) {
      return { grade: 'E', score: W.E || 0.25, explanation: 'Criterion not configured.', incompleteMinimumEvidence: true, bandBreakdown: {} };
    }

    checksMap = checksMap || {};

    if (checksMap[crit.mdOverride.id] === true || checksMap[crit.mdOverride.id] === 'true') {
      return { grade: 'A', score: W.A || 1.0, explanation: 'Distinct A evidence selected \u2014 full marks awarded.', incompleteMinimumEvidence: false, distinctOverride: true, bandBreakdown: {} };
    }

    var totalPoints = 0;
    var bandCalcs = {};
    var bandNames = ['A', 'B', 'C', 'D', 'E'];
    var totalTicked = 0;

    for (var b = 0; b < bandNames.length; b++) {
      var bandName = bandNames[b];
      var items = crit.bands[bandName] || [];
      var ticked = 0;
      for (var k = 0; k < items.length; k++) {
        if (checksMap[items[k][0]] === true || checksMap[items[k][0]] === 'true') ticked++;
      }

      bandCalcs[bandName] = { ticked: ticked, total: items.length, fraction: items.length > 0 ? (ticked / items.length) : 0 };
      totalTicked += ticked;

      if (items.length > 0) {
        var bandWeight = W[bandName] !== undefined ? W[bandName] : 0.25;
        var pointPerObs = (crit.maxMarks * bandWeight) / items.length;
        totalPoints += ticked * pointPerObs;
      }
    }

    var percentage = crit.maxMarks > 0 ? (totalPoints / crit.maxMarks) * 100 : 0;
    var grade = GradeScaleService.getOverallGradeLetter(percentage, bands);
    var score = crit.maxMarks > 0 ? totalPoints / crit.maxMarks : 0;

    var incompleteMinimumEvidence = (totalTicked === 0);

    var explanation = 'Total points: ' + totalPoints.toFixed(2) + ' / ' + crit.maxMarks + ' (' + Math.round(percentage) + '%) \u2192 nominal grade ' + grade + '.';

    return { grade: grade, score: score, explanation: explanation, incompleteMinimumEvidence: incompleteMinimumEvidence, distinctOverride: false, bandBreakdown: bandCalcs };
  }

  function getAssessmentStudioBootstrapData(targetSubId) {
    try {
      var legacyWorkbook = !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_MARKING);
      var submissions = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS).filter(function (s) {
        return legacyWorkbook && /jewell/i.test(String(s.Task || ''));
      });
      var configs = Sheets.getSheetDataAsObjects(Config.SHEET_CLASSROOM_CONFIG);
      var activeConfig = null;
      for (var c = 0; c < configs.length; c++) {
        if (configs[c].Active === true || configs[c].Active === 'true' || configs[c].Active === 'TRUE') {
          activeConfig = configs[c];
          break;
        }
      }
      var selectedId = targetSubId && submissions.some(function (s) { return s.SubmissionRecordID === targetSubId; }) ?
        targetSubId : (submissions.length > 0 ? submissions[0].SubmissionRecordID : null);
      var subDetail = selectedId ? getSubmissionForMarking(selectedId) : null;
      return {
        criteria: Config.CRITERIA_DEFINITIONS,
        submissions: submissions,
        activeConfig: activeConfig,
        selectedSubmissionId: selectedId,
        submissionDetail: subDetail,
        gradeWeights: GradeScaleService.getGradeWeights(),
        overallGradeBands: GradeScaleService.getOverallGradeBands()
      };
    } catch (err) {
      Logging.logError('getAssessmentStudioBootstrapData', err);
      return {
        criteria: Config.CRITERIA_DEFINITIONS, submissions: [], activeConfig: null, selectedSubmissionId: null, submissionDetail: null, error: err.message,
        gradeWeights: GradeScaleService.getGradeWeights(),
        overallGradeBands: GradeScaleService.getOverallGradeBands()
      };
    }
  }


  function getSubmissionForMarking(submissionRecordId) {
    var subRows = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
    var submission = null;
    for (var s = 0; s < subRows.length; s++) {
      if (subRows[s].SubmissionRecordID === submissionRecordId) {
        submission = subRows[s];
        break;
      }
    }
    if (!submission) throw new Error('Submission record ' + submissionRecordId + ' not found.');

    var allFiles = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSION_FILES);
    var files = [];
    for (var f = 0; f < allFiles.length; f++) {
      if (allFiles[f].SubmissionRecordID === submissionRecordId) files.push(allFiles[f]);
    }

    var allAssess = Sheets.getSheetDataAsObjects(Config.SHEET_CRITERION_ASSESSMENTS);
    var criteriaMap = {}, assessmentId = null;
    for (var a = 0; a < allAssess.length; a++) {
      if (allAssess[a].SubmissionRecordID === submissionRecordId) {
        var rec = allAssess[a];
        assessmentId = rec.AssessmentID;
        criteriaMap[rec.CriterionID] = {
          assessmentId: rec.AssessmentID, criterionId: rec.CriterionID,
          aiProposedGrade: rec.AIProposedGrade || '', derivedGrade: rec.DerivedGrade || 'E',
          finalApprovedGrade: rec.FinalApprovedGrade || '',
          distinctOverrideSelected: (rec.DistinctOverrideSelected === true || rec.DistinctOverrideSelected === 'TRUE'),
          checkboxes: Utils.safeJsonParse(rec.CheckboxesJSON, {}),
          aiEvidence: Utils.safeJsonParse(rec.AIEvidenceJSON, []),
          teacherAdjusted: (rec.TeacherAdjusted === true || rec.TeacherAdjusted === 'TRUE'),
          teacherAudioTranscript: rec.TeacherAudioTranscript || '', teacherWrittenNote: rec.TeacherWrittenNote || '',
          confidence: rec.Confidence || 0, minimumEvidenceIncomplete: (rec.MinimumEvidenceIncomplete === true || rec.MinimumEvidenceIncomplete === 'TRUE'),
          status: rec.Status || Config.STATUS.NEW
        };
      }
    }
    if (!assessmentId) assessmentId = Utils.generateUuid();

    var defs = Config.CRITERIA_DEFINITIONS;
    for (var c = 0; c < defs.length; c++) {
      var cid = defs[c].criterionId;
      if (!criteriaMap[cid]) {
        criteriaMap[cid] = { assessmentId: assessmentId, criterionId: cid, aiProposedGrade: '', derivedGrade: 'E', finalApprovedGrade: '', distinctOverrideSelected: false, checkboxes: {}, aiEvidence: [], teacherAdjusted: false, teacherAudioTranscript: '', teacherWrittenNote: '', confidence: 0, minimumEvidenceIncomplete: true, status: Config.STATUS.NEW };
      }
      var calc = deriveGradeFromChecks(cid, criteriaMap[cid].checkboxes);
      criteriaMap[cid].explanation = calc.explanation;
      criteriaMap[cid].bandCalculations = calc.bandBreakdown;
      criteriaMap[cid].derivedGrade = ((submission.Status === Config.STATUS.APPROVED || submission.Status === Config.STATUS.LOCKED) && criteriaMap[cid].finalApprovedGrade) ? criteriaMap[cid].finalApprovedGrade : calc.grade;
      criteriaMap[cid].minimumEvidenceIncomplete = calc.incompleteMinimumEvidence;
    }

    var allAiRuns = Sheets.getSheetDataAsObjects(Config.SHEET_AI_ASSESSMENTS);
    var latestFeedback = null;
    for (var r = allAiRuns.length - 1; r >= 0; r--) {
      if (allAiRuns[r].SubmissionRecordID === submissionRecordId && allAiRuns[r].RunType === 'FeedbackGeneration') {
        var parsed = Utils.safeJsonParse(allAiRuns[r].ResponseJSON, null);
        if (parsed) { latestFeedback = parsed; break; }
      }
    }
    return { submission: submission, files: files, assessmentId: assessmentId, criteriaMap: criteriaMap, feedback: latestFeedback || { whatWentWell: '', areasForImprovement: '', goalsForNextAssessment: '' } };
  }

  function rejectFinalSubmission(submissionRecordId) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_SUBMISSIONS);
    if (!sheet) return { success: false, message: 'Submissions sheet is missing.' };
    var row = Sheets.findRowByKey(sheet, 'SubmissionRecordID', submissionRecordId);
    if (row === -1) return { success: false, message: 'Submission not found.' };
    var headers = Sheets.getHeaderIndex(sheet);
    var status = sheet.getRange(row, headers.Status + 1).getValue();
    if (status === Config.STATUS.APPROVED || status === Config.STATUS.LOCKED || status === Config.STATUS.SUPERSEDED) {
      return { success: false, message: 'This assessment is final. Create a new reassessment version first.' };
    }
    if (headers.CurrentOfficial !== undefined && sheet.getRange(row, headers.CurrentOfficial + 1).getValue() === false) {
      return { success: false, message: 'This is a historical submission version.' };
    }
    return null;
  }

  function saveCriterionDraft(assessmentId, submissionRecordId, criterionId, state) {
    var blocked = rejectFinalSubmission(submissionRecordId);
    if (blocked) return blocked;
    if (!Config.CRITERIA_DEFINITIONS.some(function (c) { return c.criterionId === criterionId; })) {
      return { success: false, message: 'Criterion is not part of the legacy Jewellery rubric.' };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(Config.SHEET_CRITERION_ASSESSMENTS);
    var checksMap = state.checkboxes || {};
    var calc = deriveGradeFromChecks(criterionId, checksMap);

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] === assessmentId && data[r][2] === criterionId) {
        var oldStatus = data[r][Sheets.getHeaderIndex(sheet).Status];
        if (data[r][1] !== submissionRecordId || oldStatus === Config.STATUS.APPROVED || oldStatus === Config.STATUS.LOCKED) {
          return { success: false, message: 'Cannot alter another submission or an approved criterion.' };
        }
        rowIndex = r + 1; break;
      }
    }
    var now = Utils.formatDate(new Date());
    var rowValues = [
      assessmentId, submissionRecordId, criterionId, state.aiProposedGrade || '', calc.grade, '',
      !!checksMap[criterionId + '-MD'], Utils.safeJsonStringify(checksMap), Utils.safeJsonStringify(state.aiEvidence || []),
      !!state.teacherAdjusted, state.teacherAudioTranscript || '', state.teacherWrittenNote || '', state.confidence || 1.0,
      calc.incompleteMinimumEvidence, Config.STATUS.IN_REVIEW, now, now, '', ''
    ];

    if (rowIndex !== -1) sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    else sheet.appendRow(rowValues);

    Logging.logHistory(submissionRecordId, assessmentId, 'DraftSaved', criterionId, {}, state, 'Updated draft for ' + criterionId);
    return { success: true, derivedGrade: calc.grade, explanation: calc.explanation, minimumEvidenceIncomplete: calc.incompleteMinimumEvidence, bandCalculations: calc.bandBreakdown };
  }

  function saveAssessmentDraft(assessmentId, submissionRecordId, criteriaMap, feedbackObj) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: 'Could not obtain lock.' }; }
    try {
      var blocked = rejectFinalSubmission(submissionRecordId);
      if (blocked) return blocked;
      for (var cid in criteriaMap) {
        if (!Object.prototype.hasOwnProperty.call(criteriaMap, cid)) continue;
        var saved = saveCriterionDraft(assessmentId, submissionRecordId, cid, criteriaMap[cid]);
        if (!saved.success) return saved;
      }
      if (feedbackObj) {
        var aiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_AI_ASSESSMENTS);
        if (aiSheet) {
          aiSheet.appendRow([assessmentId, submissionRecordId, 'FeedbackGeneration', Utils.formatDate(new Date()), Config.PRIMARY_MODEL, 'TeacherDraft', '[]', 'Teacher edited feedback', Utils.safeJsonStringify(feedbackObj), 'Success', '', 1.0, Utils.getSafeUserEmail(), 'Accepted']);
        }
      }
      updateSubmissionStatus(submissionRecordId, Config.STATUS.TEACHER_REVIEWED);
      return { success: true, message: 'Draft assessment saved successfully.' };
    } catch (err) {
      Logging.logError('saveAssessmentDraft', err);
      return { success: false, message: 'Failed to save: ' + Utils.sanitizeError(err) };
    } finally { lock.releaseLock(); }
  }

  function approveAssessment(assessmentId, submissionRecordId, feedbackObj) {
    var blocked = rejectFinalSubmission(submissionRecordId);
    if (blocked) return blocked;
    var detail = getSubmissionForMarking(submissionRecordId);
    if (String(detail.assessmentId) !== String(assessmentId)) return { success: false, message: 'Assessment ID has changed; reload first.' };
    var defs = Config.CRITERIA_DEFINITIONS;
    var stored = Sheets.getSheetDataAsObjects(Config.SHEET_CRITERION_ASSESSMENTS).filter(function (r) {
      return r.AssessmentID === assessmentId && r.SubmissionRecordID === submissionRecordId;
    });
    var grades = {};
    for (var i = 0; i < defs.length; i++) {
      var cid = defs[i].criterionId;
      var cd = detail.criteriaMap[cid];
      var row = stored.filter(function (r) { return r.CriterionID === cid; })[0];
      if (!row || !cd || cd.minimumEvidenceIncomplete) {
        return { success: false, message: 'Review and save evidence for ' + cid + ' before approval.' };
      }
      grades[cid] = cd.derivedGrade;
    }

    var now = Utils.formatDate(new Date());
    var userEmail = Utils.getSafeUserEmail();
    Sheets.syncApprovedGradesToMarkingSheet(detail.submission.StudentName, grades, feedbackObj);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cSheet = ss.getSheetByName(Config.SHEET_CRITERION_ASSESSMENTS);
    if (cSheet) {
      var cData = cSheet.getDataRange().getValues();
      for (var r = 1; r < cData.length; r++) {
        if (cData[r][0] === assessmentId) {
          Sheets.setCellByHeader(cSheet, r + 1, 'FinalApprovedGrade', cData[r][4]);
          Sheets.setCellByHeader(cSheet, r + 1, 'Status', Config.STATUS.APPROVED);
          Sheets.setCellByHeader(cSheet, r + 1, 'ApprovedAt', now);
          Sheets.setCellByHeader(cSheet, r + 1, 'ApprovedBy', userEmail);
        }
      }
    }

    var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
    if (sSheet) {
      var sRow = Sheets.findRowByKey(sSheet, 'SubmissionRecordID', submissionRecordId);
      if (sRow !== -1) {
        Sheets.setCellByHeader(sSheet, sRow, 'Status', Config.STATUS.APPROVED);
        Sheets.setCellByHeader(sSheet, sRow, 'ApprovedAt', now);
        Sheets.setCellByHeader(sSheet, sRow, 'ApprovedBy', userEmail);
      }
    }
    Logging.logHistory(submissionRecordId, assessmentId, 'Approved', '', {}, grades, 'Approved for ' + detail.submission.StudentName);
    return { success: true, message: 'Assessment approved. Letter grades (fluidly derived from Setup grade scale) synced to ' + Config.SHEET_MARKING + '.' };
  }

  function lockAssessment(assessmentId, submissionRecordId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
    var now = Utils.formatDate(new Date());
    var user = Utils.getSafeUserEmail();
    var row = Sheets.findRowByKey(sSheet, 'SubmissionRecordID', submissionRecordId);
    if (row === -1) return { success: false, message: 'Submission not found.' };
    if (String(sSheet.getRange(row, Sheets.getHeaderIndex(sSheet)['Status'] + 1).getValue()) !== Config.STATUS.APPROVED) {
      return { success: false, message: 'Assessment must be approved before locking.' };
    }

    Sheets.setCellByHeader(sSheet, row, 'Status', Config.STATUS.LOCKED);
    Sheets.setCellByHeader(sSheet, row, 'LockedAt', now);
    Sheets.setCellByHeader(sSheet, row, 'LockedBy', user);

    var cSheet = ss.getSheetByName(Config.SHEET_CRITERION_ASSESSMENTS);
    if (cSheet) {
      var cData = cSheet.getDataRange().getValues();
      for (var cr = 1; cr < cData.length; cr++) {
        if (cData[cr][0] === assessmentId) Sheets.setCellByHeader(cSheet, cr + 1, 'Status', Config.STATUS.LOCKED);
      }
    }
    Logging.logHistory(submissionRecordId, assessmentId, 'Locked', '', {}, {}, 'Locked by ' + user);
    return { success: true, message: 'Assessment permanently locked.' };
  }

  function createReassessmentFromSubmission(submissionRecordId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
    var rows = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
    var parent = null;
    for (var p = 0; p < rows.length; p++) {
      if (rows[p].SubmissionRecordID === submissionRecordId) { parent = rows[p]; break; }
    }
    if (!parent) return { success: false, message: 'Parent submission not found.' };
    if (parent.Status !== Config.STATUS.APPROVED && parent.Status !== Config.STATUS.LOCKED) {
      return { success: false, message: 'Approve the assessment before creating a reassessment.' };
    }
    if (parent.CurrentOfficial === false || parent.CurrentOfficial === 'FALSE') {
      return { success: false, message: 'A newer version is already official.' };
    }
    if (rows.some(function (r) { return r.ParentSubmissionRecordID === submissionRecordId; })) {
      return { success: false, message: 'A reassessment already exists for this version.' };
    }

    var newSubId = Utils.generateUuid();
    var newVer = (parseInt(parent.SubmissionVersion, 10) || 1) + 1;
    var now = Utils.formatDate(new Date());

    var parentRow = Sheets.findRowByKey(sSheet, 'SubmissionRecordID', submissionRecordId);
    sSheet.appendRow([
      newSubId, parent.ClassroomCourseID, parent.ClassroomCourseWorkID, parent.ClassroomSubmissionID,
      parent.StudentUserID, parent.StudentName, parent.StudentEmail, parent.Class, parent.Task,
      newVer, 'Reassessment', parent.ClassroomState, parent.TurnedInTime, now, parent.Late,
      parent.AttachmentSummary, parent.AttachmentFileIDsJSON, parent.AttachmentMetadataJSON,
      parent.DriveFolderID, Config.STATUS.NEW, submissionRecordId, false, '', '', '', '', '', '', '', 'Reassessment created'
    ]);

    var allFiles = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSION_FILES);
    var fSheet = ss.getSheetByName(Config.SHEET_SUBMISSION_FILES);
    if (fSheet) {
      for (var f = 0; f < allFiles.length; f++) {
        if (allFiles[f].SubmissionRecordID === submissionRecordId) {
          var fo = allFiles[f];
          fSheet.appendRow([newSubId, Utils.generateUuid(), fo.SourceType, fo.DriveFileID, fo.FileName, fo.MimeType, fo.AlternateLink, fo.ThumbnailUrl, fo.FileSize, fo.EligibleForAI, 'Pending', '', '', now]);
        }
      }
    }
    /* Only switch official status once the new row and its files exist. */
    if (parentRow !== -1) Sheets.setCellByHeader(sSheet, parentRow, 'CurrentOfficial', false);
    var childRow = Sheets.findRowByKey(sSheet, 'SubmissionRecordID', newSubId);
    if (childRow !== -1) Sheets.setCellByHeader(sSheet, childRow, 'CurrentOfficial', true);
    Logging.logHistory(newSubId, '', 'ReassessmentCreated', '', {}, { parentId: submissionRecordId, version: newVer }, 'Reassessment v' + newVer);
    return { success: true, message: 'Reassessment version ' + newVer + ' created successfully.', newSubmissionRecordId: newSubId };
  }

  function updateSubmissionStatus(submissionRecordId, status) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
    if (!sSheet) return;
    var data = sSheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] === submissionRecordId) { sSheet.getRange(r + 1, 20).setValue(status); break; }
    }
  }

  return {
    rejectFinalSubmission: rejectFinalSubmission,
    deriveGradeFromChecks: deriveGradeFromChecks, getAssessmentStudioBootstrapData: getAssessmentStudioBootstrapData,
    getSubmissionForMarking: getSubmissionForMarking, saveCriterionDraft: saveCriterionDraft,
    saveAssessmentDraft: saveAssessmentDraft, approveAssessment: approveAssessment,
    lockAssessment: lockAssessment, createReassessmentFromSubmission: createReassessmentFromSubmission,
    updateSubmissionStatus: updateSubmissionStatus
  };
})();

/* ============================================================================
 * SECTION 5: GOOGLE DRIVE, ATTACHMENTS & PDF EXPORTS
 * ============================================================================ */
var DriveService = (function() {
  var SUPPORTED_IMAGE_MIMES = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp', 'image/gif', 'image/heic'];
  var SUPPORTED_DOC_MIMES = ['application/vnd.google-apps.document', 'application/vnd.google-apps.presentation', 'application/vnd.google-apps.drawing', 'application/pdf'];

  function isEligibleForAi(mimeType) {
    if (!mimeType) return false;
    mimeType = mimeType.toLowerCase();
    return SUPPORTED_IMAGE_MIMES.indexOf(mimeType) !== -1 || SUPPORTED_DOC_MIMES.indexOf(mimeType) !== -1;
  }

  /* Reads ONE file from Drive and returns it prepared for Gemini, or a
     status explaining why it could not be used. This is the function that
     genuinely touches Drive — call it once per file, one at a time. */
  function prepareSingleFileForAi(fileObj) {
    var fileId = fileObj.DriveFileID;
    var fileName = fileObj.FileName || 'Attachment';
    var mime = (fileObj.MimeType || '').toLowerCase();

    if (!isEligibleForAi(mime)) {
      return { fileId: fileId, fileName: fileName, status: 'unsupported' };
    }
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = (mime.indexOf('google-apps') !== -1) ? file.getAs('application/pdf') : file.getBlob();
      var bytes = blob.getBytes();
      if (bytes.length / (1024 * 1024) > 15) {
        return { fileId: fileId, fileName: fileName, status: 'partial', limitations: ['Exceeds 15MB limit'] };
      }
      var effectiveMime = (mime.indexOf('google-apps') !== -1) ? 'application/pdf' : (blob.getContentType() || mime);
      return {
        fileId: fileId,
        fileName: fileName,
        status: 'reviewed',
        mimeType: effectiveMime,
        inlineData: { mimeType: effectiveMime, data: Utilities.base64Encode(bytes) }
      };
    } catch (err) {
      return { fileId: fileId, fileName: fileName, status: 'unreadable' };
    }
  }

  /* Bulk variant retained for the separate Class-Wide AI Grading Runner,
     which still processes one whole student at a time (already genuinely
     sequential at the student level) and does not need per-file granularity. */
  function prepareSubmissionFilesForAi(submissionRecordId) {
    var allFiles = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSION_FILES);
    var targetFiles = [];
    for (var a = 0; a < allFiles.length; a++) {
      if (allFiles[a].SubmissionRecordID === submissionRecordId) targetFiles.push(allFiles[a]);
    }
    var prepared = [], metadata = [];
    for (var j = 0; j < targetFiles.length; j++) {
      var result = prepareSingleFileForAi(targetFiles[j]);
      metadata.push({ fileId: result.fileId, fileName: result.fileName, status: result.status, limitations: result.limitations });
      if (result.status === 'reviewed') {
        prepared.push({ fileId: result.fileId, fileName: result.fileName, mimeType: result.mimeType, inlineData: result.inlineData });
      }
    }
    return { files: prepared, metadata: metadata };
  }

  function exportDriveFileAsPdf(driveFileId) {
    try {
      var file = DriveApp.getFileById(driveFileId);
      var mime = file.getMimeType();
      var blob;
      if (mime.indexOf('google-apps') !== -1) {
        blob = file.getAs('application/pdf');
      } else if (mime === 'application/pdf') {
        blob = file.getBlob();
      } else {
        blob = file.getBlob();
      }
      var cleanName = file.getName().replace(/\.[^/.]+$/, '') + '.pdf';
      var base64 = Utilities.base64Encode(blob.getBytes());
      return {
        success: true,
        fileName: cleanName,
        dataUrl: 'data:application/pdf;base64,' + base64
      };
    } catch (e) {
      return { success: false, message: 'Could not export Drive file: ' + Utils.sanitizeError(e) };
    }
  }

  return {
    isEligibleForAi: isEligibleForAi,
    prepareSingleFileForAi: prepareSingleFileForAi,
    prepareSubmissionFilesForAi: prepareSubmissionFilesForAi,
    exportDriveFileAsPdf: exportDriveFileAsPdf
  };
})();

/* ============================================================================
 * SECTION 6: GOOGLE CLASSROOM INTEGRATION
 * FIXED: importOrRefreshClassroomSubmissions no longer passes courseWorkStates
 * (invalid for studentSubmissions.list) and now paginates through all results.
 * ============================================================================ */
var ClassroomService = (function() {
  function listTeacherCourses() {
    try {
      var res = Classroom.Courses.list({ teacherId: 'me', courseStates: ['ACTIVE'], pageSize: 50 });
      var courses = [];
      var raw = res.courses || [];
      for (var i = 0; i < raw.length; i++) {
        courses.push({ id: raw[i].id, name: raw[i].name, section: raw[i].section || '', room: raw[i].room || '' });
      }
      return { success: true, courses: courses };
    } catch (err) {
      Logging.logError('listTeacherCourses', err);
      return { success: false, message: 'Classroom courses error: ' + Utils.sanitizeError(err) };
    }
  }

  function listCourseWork(courseId) {
    try {
      var coursework = [];
      var pageToken = null;
      do {
        var params = { pageSize: 100 };
        if (pageToken) params.pageToken = pageToken;
        var res = Classroom.Courses.CourseWork.list(String(courseId), params);
        var raw = (res && res.courseWork) ? res.courseWork : [];
        for (var i = 0; i < raw.length; i++) {
          var cw = raw[i];
          var dueDate = '';
          if (cw.dueDate) {
            var month = String(cw.dueDate.month);
            var day = String(cw.dueDate.day);
            dueDate = String(cw.dueDate.year) + '-' +
              (month.length === 1 ? '0' : '') + month + '-' +
              (day.length === 1 ? '0' : '') + day;
          }
          coursework.push({
            id: cw.id,
            title: cw.title || '(Untitled assignment)',
            maxPoints: cw.maxPoints == null ? 0 : cw.maxPoints,
            dueDate: dueDate,
            state: cw.state || ''
          });
        }
        pageToken = (res && res.nextPageToken) ? res.nextPageToken : null;
      } while (pageToken);
      coursework.sort(function(a, b) {
        var at = String(a.title).toLowerCase();
        var bt = String(b.title).toLowerCase();
        return at < bt ? -1 : (at > bt ? 1 : 0);
      });
      return { success: true, coursework: coursework };
    } catch (err) {
      Logging.logError('listCourseWork', err, { courseId: courseId });
      return { success: false, message: 'Coursework error: ' + Utils.sanitizeError(err) };
    }
  }

  function saveClassroomSelection(courseId, courseWorkId) {
    try {
      var course = Classroom.Courses.get(courseId);
      var cw = Classroom.Courses.CourseWork.get(courseId, courseWorkId);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG);
      if (!sheet) { Sheets.setupOrMigrateAssessmentSystem(); sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG); }

      var data = sheet.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        if (!String(data[r][10] || '').trim()) sheet.getRange(r + 1, 10).setValue(false);
      }

      sheet.appendRow([Utils.generateUuid(), courseId, course.name || '', course.section || '', courseWorkId, cw.title || '', cw.maxPoints || 100, Utils.formatDate(new Date()), Utils.getSafeUserEmail(), true, '', '', '']);
      return { success: true, message: 'Saved assignment: ' + cw.title };
    } catch (err) {
      Logging.logError('saveClassroomSelection', err);
      return { success: false, message: 'Save selection error: ' + Utils.sanitizeError(err) };
    }
  }

  function saveClassroomSelectionForTask(courseId, courseWorkId, taskName) {
    try {
      var course = Classroom.Courses.get(courseId);
      var cw = Classroom.Courses.CourseWork.get(courseId, courseWorkId);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG);
      if (!sheet) { Sheets.setupOrMigrateAssessmentSystem(); sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG); }
      ensureClassroomConfigColumns_(sheet);
      if (!String(taskName || '').trim()) return { success: false, message: 'Task name is required.' };

      // We do NOT set other rows to Active: false here.
      // Multiple rows may have Active: true simultaneously as long as they have a distinct TaskName.

      var dueStr = '';
      if (cw.dueDate) {
        var yy = cw.dueDate.year;
        var mm = cw.dueDate.month < 10 ? '0' + cw.dueDate.month : cw.dueDate.month;
        var dd = cw.dueDate.day < 10 ? '0' + cw.dueDate.day : cw.dueDate.day;
        var hh = (cw.dueTime && cw.dueTime.hours != null) ? ('0' + cw.dueTime.hours).slice(-2) : '23';
        var min = (cw.dueTime && cw.dueTime.minutes != null) ? ('0' + cw.dueTime.minutes).slice(-2) : '59';
        var sec = (cw.dueTime && cw.dueTime.seconds != null) ? ('0' + cw.dueTime.seconds).slice(-2) : '00';
        dueStr = yy + '-' + mm + '-' + dd + 'T' + hh + ':' + min + ':' + sec + 'Z';
      }

      sheet.appendRow([
        Utils.generateUuid(), courseId, course.name || '', course.section || '', courseWorkId,
        cw.title || '', cw.maxPoints || 100, Utils.formatDate(new Date()), Utils.getSafeUserEmail(),
        true, taskName, dueStr, false
      ]);
      return { success: true, message: 'Saved assignment for task: ' + taskName };
    } catch (err) {
      Logging.logError('saveClassroomSelectionForTask', err);
      return { success: false, message: 'Save selection error: ' + Utils.sanitizeError(err) };
    }
  }

  function checkAndAutoImportDueAssignments() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG);
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // Header indices
    var hConfigId = headers.indexOf('ConfigID');
    var hCourseId = headers.indexOf('CourseID');
    var hCourseWorkId = headers.indexOf('CourseWorkID');
    var hActive = headers.indexOf('Active');
    var hDueDate = headers.indexOf('DueDate');
    var hAutoImported = headers.indexOf('AutoImported');

    if (hDueDate === -1 || hAutoImported === -1 || hActive === -1) {
      Logging.logError('checkAndAutoImportDueAssignments', new Error('ClassroomConfig is missing scheduling columns. Run Setup / Migrate Assessment System.'));
      return;
    }

    var now = new Date();

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var isActive = row[hActive] === true || row[hActive] === 'TRUE' || row[hActive] === 'true';
      var autoImported = row[hAutoImported] === true || row[hAutoImported] === 'TRUE' || row[hAutoImported] === 'true';
      var dueDateStr = row[hDueDate];

      if (isActive && !autoImported && dueDateStr) {
        var dueDateObj = new Date(dueDateStr);
        if (!isNaN(dueDateObj.getTime()) && dueDateObj < now) {
          // Time to auto-import this specific assignment
          try {
            // Re-use logic to mock getActiveClassroomConfig just for this call context
            // by injecting the CourseID and CourseWorkID explicitly into the import function.
            // Since importOrRefreshClassroomSubmissions currently relies on getActiveClassroomConfig(),
            // we will pass the row config directly.

            var cfg = {
              CourseID: row[hCourseId],
              CourseWorkID: row[hCourseWorkId],
              CourseName: row[headers.indexOf('CourseName')],
              AssignmentTitle: row[headers.indexOf('AssignmentTitle')],
              TaskName: row[headers.indexOf('TaskName')]
            };

            var importResult = ClassroomService.importOrRefreshClassroomSubmissions(cfg);
            var gradingFailed = false;

            /* Task-scoped assignments use their own reviewed rubrics in the
               web app. Import here, but NEVER send them to the legacy hardcoded
               Jewellery AI prompt; Dan can run task-aware AI in the web UI. */
            if (importResult.success && importResult.stats && !cfg.TaskName) {
              // Legacy Jewellery-only import: trigger the original grader
              var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
              var sData = sSheet.getDataRange().getValues();

              for (var s = 1; s < sData.length; s++) {
                // If it's a NEW status and matches our assignment, auto-grade
                if (sData[s][1] === cfg.CourseID && sData[s][2] === cfg.CourseWorkID && sData[s][19] === Config.STATUS.NEW) {
                  try {
                    var subId = sData[s][0];
                    var gradingResult = GeminiService.runInitialAiAssessment(subId);
                    if (!gradingResult || !gradingResult.success) {
                      gradingFailed = true;
                      Logging.logError('checkAndAutoImportDueAssignments (Grading)',
                        new Error(gradingResult && gradingResult.message || 'AI grading failed.'));
                    }
                  } catch (e) {
                    gradingFailed = true;
                    Logging.logError('checkAndAutoImportDueAssignments (Grading)', e);
                  }
                }
              }
            }

            // Mark complete only after the import and every attempted assessment succeeds.
            if (importResult.success && importResult.stats && !gradingFailed) {
              sheet.getRange(r + 1, hAutoImported + 1).setValue(true);
            } else {
              Logging.logError('checkAndAutoImportDueAssignments (Row ' + r + ')',
                new Error(importResult.message || 'Import or grading failed; will retry next hour.'));
            }

          } catch (e) {
            Logging.logError('checkAndAutoImportDueAssignments (Row ' + r + ')', e);
          }
        }
      }
    }
  }

  function ensureClassroomConfigColumns_(sheet) {
    var required = ['TaskName', 'DueDate', 'AutoImported'];
    var lastCol = sheet.getLastColumn();
    var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var missing = [];
    for (var i = 0; i < required.length; i++) {
      if (headers.indexOf(required[i]) === -1) missing.push(required[i]);
    }
    if (missing.length) {
      var start = headers.length + 1;
      if (sheet.getMaxColumns() < start + missing.length - 1) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), start + missing.length - 1 - sheet.getMaxColumns());
      }
      sheet.getRange(1, start, 1, missing.length).setValues([missing]);
    }
  }

  function getActiveClassroomConfig() {
    var configs = Sheets.getSheetDataAsObjects(Config.SHEET_CLASSROOM_CONFIG);
    for (var i = configs.length - 1; i >= 0; i--) {
      if (!String(configs[i].TaskName || '').trim() &&
          (configs[i].Active === true || configs[i].Active === 'TRUE' || configs[i].Active === 'true')) {
        return configs[i];
      }
    }
    return null;
  }

  /* Paginates through ALL student submissions instead of a single 100-item page. */
  function listAllStudentSubmissions(courseId, courseWorkId) {
    var all = [];
    var pageToken = null;
    do {
      var params = { pageSize: 100 };
      if (pageToken) params.pageToken = pageToken;
      /* IMPORTANT: courseWorkStates is NOT a valid parameter for studentSubmissions.list */
      var response = Classroom.Courses.CourseWork.StudentSubmissions.list(String(courseId), String(courseWorkId), params);
      if (response && response.studentSubmissions) all = all.concat(response.studentSubmissions);
      pageToken = (response && response.nextPageToken) ? response.nextPageToken : null;
    } while (pageToken);
    return all;
  }

  function importOrRefreshClassroomSubmissions(overrideCfg) {
    var cfg = overrideCfg || getActiveClassroomConfig();
    if (!cfg) return { success: false, message: 'No active Classroom assignment configured. Please select one via Classroom menu.' };

    var rosterMap = RosterService.getClassRosterMap(cfg.CourseName);

    try {
      var submissions = listAllStudentSubmissions(cfg.CourseID, cfg.CourseWorkID);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
      var fSheet = ss.getSheetByName(Config.SHEET_SUBMISSION_FILES);

      if (!sSheet || !fSheet) {
        Sheets.setupOrMigrateAssessmentSystem();
        sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
        fSheet = ss.getSheetByName(Config.SHEET_SUBMISSION_FILES);
      }
      if (!sSheet) return { success: false, message: 'Required sheet "' + Config.SHEET_SUBMISSIONS + '" could not be created or found.' };

      var sData = sSheet.getDataRange().getValues();
      var existing = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
      var subMap = {};
      for (var e = 0; e < existing.length; e++) {
        var row = existing[e];
        if (!row.StudentUserID) continue;
        /* Scope prior records to THIS Classroom assignment so importing a
           second task can never overwrite another task's rows (AUDIT R5). */
        var mapKey = (row.ClassroomCourseID || '') + '|' + (row.ClassroomCourseWorkID || '') + '|' + row.StudentUserID;
        if (!subMap[mapKey]) subMap[mapKey] = [];
        subMap[mapKey].push(row);
      }

      var stats = { total: submissions.length, newCount: 0, updated: 0, newVersion: 0, skipped: 0 };
      var now = Utils.formatDate(new Date());
      var taskLabel = String(cfg.TaskName || '').trim() || cfg.AssignmentTitle || 'General';

      for (var i = 0; i < submissions.length; i++) {
        var cSub = submissions[i];
        var uid = cSub.userId;
        var sName = 'Student ' + (i + 1);
        var sEmail = '';

        try {
          var p = Classroom.UserProfiles.get(uid);
          if (p && p.name && p.name.fullName) sName = p.name.fullName;
          if (p && p.emailAddress) sEmail = p.emailAddress;
        } catch (pe) {
          sName = 'Student ' + (i + 1);
        }

        var identity = RosterService.resolveStudentIdentity(sEmail, sName, rosterMap);
        sName = identity.displayName;

        var attachments = [], fileIds = [];
        if (cSub.assignmentSubmission && cSub.assignmentSubmission.attachments) {
          var rawAtts = cSub.assignmentSubmission.attachments;
          for (var a = 0; a < rawAtts.length; a++) {
            var att = rawAtts[a];
            if (att.driveFile) {
              attachments.push({ sourceType: 'driveFile', id: att.driveFile.id, title: att.driveFile.title, alternateLink: att.driveFile.alternateLink, thumbnailUrl: att.driveFile.thumbnailUrl || '' });
              fileIds.push(att.driveFile.id);
            }
          }
        }

        var userHistory = subMap[(cfg.CourseID || '') + '|' + (cfg.CourseWorkID || '') + '|' + uid] || [];
        userHistory.sort(function(a, b) { return (parseInt(b.SubmissionVersion, 10) || 1) - (parseInt(a.SubmissionVersion, 10) || 1); });
        var latest = userHistory.length > 0 ? userHistory[0] : null;

        if (latest && latest.Status === Config.STATUS.NEW && latest.StudentName !== sName) {
          for (var r = 1; r < sData.length; r++) {
            if (sData[r][0] === latest.SubmissionRecordID) {
              sSheet.getRange(r + 1, 6).setValue(sName);
              sData[r][5] = sName; // Update local cache to prevent re-writing
              break;
            }
          }
        }

        var finalRecordId;

        if (!latest) {
          var newId = Utils.generateUuid();
          finalRecordId = newId;
          sSheet.appendRow([
            newId, cfg.CourseID, cfg.CourseWorkID, cSub.id, uid, sName, sEmail, cfg.CourseName || '', taskLabel,
            1, 'Draft', cSub.state || 'NEW', cSub.creationTime || now, cSub.updateTime || now, cSub.late || false,
            attachments.length + ' attachment(s)', Utils.safeJsonStringify(fileIds), Utils.safeJsonStringify(attachments),
            '', Config.STATUS.NEW, '', true, '', '', '', '', cSub.assignedGrade || '', '', '', 'Initial import'
          ]);
          insertFiles(fSheet, newId, attachments);
          stats.newCount++;
        } else {
          finalRecordId = latest.SubmissionRecordID;
          var prevIds = Utils.safeJsonParse(latest.AttachmentFileIDsJSON, []);
          var changed = (JSON.stringify(prevIds.sort()) !== JSON.stringify(fileIds.sort())) || (cSub.updateTime && cSub.updateTime !== latest.UpdateTime);
          if (changed) {
            if (latest.Status !== Config.STATUS.NEW) {
              var nVer = (parseInt(latest.SubmissionVersion, 10) || 1) + 1;
              var vId = Utils.generateUuid();
              finalRecordId = vId;
              var prevRow = Sheets.findRowByKey(sSheet, 'SubmissionRecordID', latest.SubmissionRecordID);
              sSheet.appendRow([
                vId, cfg.CourseID, cfg.CourseWorkID, cSub.id, uid, sName, sEmail, cfg.CourseName || '', taskLabel,
                nVer, 'Reassessment', cSub.state || 'UPDATED', cSub.creationTime || now, cSub.updateTime || now, cSub.late || false,
                attachments.length + ' attachment(s)', Utils.safeJsonStringify(fileIds), Utils.safeJsonStringify(attachments),
                '', Config.STATUS.NEW, latest.SubmissionRecordID, false, '', '', '', '', cSub.assignedGrade || '', '', '', 'Classroom update reassessment'
              ]);
              insertFiles(fSheet, vId, attachments);
              if (prevRow !== -1) Sheets.setCellByHeader(sSheet, prevRow, 'CurrentOfficial', false);
              var nextRow = Sheets.findRowByKey(sSheet, 'SubmissionRecordID', vId);
              if (nextRow !== -1) Sheets.setCellByHeader(sSheet, nextRow, 'CurrentOfficial', true);
              Logging.logHistory(vId, '', 'ClassroomReassessmentImported', '', {},
                { parentId: latest.SubmissionRecordID, version: nVer }, 'Updated Classroom evidence; old assessment preserved.');
              stats.newVersion++;
            } else {
              updateSubRow(sSheet, latest.SubmissionRecordID, cSub, attachments, fileIds);
              stats.updated++;
            }
          } else {
            stats.skipped++;
          }
        }
      }

      return {
        success: true,
        message: 'Classroom sync completed: ' + stats.newCount + ' new, ' + stats.updated + ' updated, ' + stats.newVersion + ' new versions, ' + stats.skipped + ' unchanged.',
        stats: stats
      };
    } catch (err) {
      Logging.logError('importOrRefreshClassroomSubmissions', err);
      return { success: false, message: 'Import failed: ' + Utils.sanitizeError(err) };
    }
  }

  function insertFiles(fSheet, subId, attachments) {
    if (!fSheet || !attachments.length) return;
    var rows = [], now = Utils.formatDate(new Date());
    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i], mime = '', size = 0;
      if (att.id) {
        try { var df = DriveApp.getFileById(att.id); mime = df.getMimeType(); size = df.getSize(); } catch (e) {}
      }
      var eligible = DriveService.isEligibleForAi(mime);
      rows.push([subId, Utils.generateUuid(), att.sourceType, att.id || '', att.title, mime, att.alternateLink || '', att.thumbnailUrl || '', size, eligible, eligible ? 'Pending' : 'Unsupported', '', '', now]);
    }
    if (rows.length) fSheet.getRange(fSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  function updateSubRow(sSheet, recordId, cSub, attachments, fileIds) {
    var data = sSheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] === recordId) {
        sSheet.getRange(r + 1, 12).setValue(cSub.state || 'UPDATED');
        sSheet.getRange(r + 1, 14).setValue(cSub.updateTime || Utils.formatDate(new Date()));
        sSheet.getRange(r + 1, 16).setValue(attachments.length + ' attachment(s)');
        sSheet.getRange(r + 1, 17).setValue(Utils.safeJsonStringify(fileIds));
        sSheet.getRange(r + 1, 18).setValue(Utils.safeJsonStringify(attachments));
        break;
      }
    }
  }

  function syncClassListFromClassroom(courseId) {
    try {
      var course = Classroom.Courses.get(courseId);
      var courseName = course.name || courseId;

      var students = [];
      var pageToken = null;
      do {
        var params = { pageSize: 100 };
        if (pageToken) params.pageToken = pageToken;
        var response = Classroom.Courses.Students.list(String(courseId), params);
        if (response && response.students) students = students.concat(response.students);
        pageToken = (response && response.nextPageToken) ? response.nextPageToken : null;
      } while (pageToken);

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("ClassLists");
      if (!sheet) {
        return { success: false, message: "ClassLists sheet not found." };
      }

      var existingData = sheet.getDataRange().getValues();
      var emailToRowIndex = {};
      for (var r = 1; r < existingData.length; r++) {
        var em = String(existingData[r][3] || "").toLowerCase().trim();
        // Only consider students in the current course we are syncing
        if (em && String(existingData[r][0] || "").toLowerCase().trim() === String(courseName).toLowerCase().trim()) {
          emailToRowIndex[em] = r + 1;
        }
      }

      var classroomEmails = {};
      var newRows = [];
      var updatedCount = 0;
      var newCount = 0;

      for (var i = 0; i < students.length; i++) {
        var s = students[i];
        if (!s.profile || !s.profile.emailAddress) continue;

        var email = s.profile.emailAddress.toLowerCase().trim();
        var officialName = s.profile.name ? s.profile.name.fullName : "Unknown";
        classroomEmails[email] = true;

        if (emailToRowIndex[email]) {
          var rowIdx = emailToRowIndex[email];
          var currentActive = existingData[rowIdx - 1][4];
          if (currentActive !== true && currentActive !== "TRUE" && currentActive !== "true") {
             sheet.getRange(rowIdx, 5).setValue(true);
             updatedCount++;
          }
        } else {
          newRows.push([courseName, officialName, "", email, true]);
          newCount++;
        }
      }

      // Update dropped students (in sheet for this course but NOT in Classroom)
      for (var sheetEmail in emailToRowIndex) {
        if (!classroomEmails[sheetEmail]) {
          var dropIdx = emailToRowIndex[sheetEmail];
          var dropActive = existingData[dropIdx - 1][4];
          if (dropActive === true || dropActive === "TRUE" || dropActive === "true") {
            sheet.getRange(dropIdx, 5).setValue(false);
            updatedCount++;
          }
        }
      }

      if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }

      return { success: true, message: "Class list synced. Added " + newCount + ", updated active status for " + updatedCount + " students." };
    } catch (err) {
      Logging.logError('syncClassListFromClassroom', err);
      return { success: false, message: "Sync failed: " + Utils.sanitizeError(err) };
    }
  }

  function syncAssessmentToClassroom(assessmentId, submissionRecordId) {
    var detail = AssessmentService.getSubmissionForMarking(submissionRecordId);
    var sub = detail.submission;
    if (sub.Status !== Config.STATUS.APPROVED && sub.Status !== Config.STATUS.LOCKED) {
      return { success: false, message: 'Only Approved or Locked assessments can be synced.' };
    }
    try {
      var cw = Classroom.Courses.CourseWork.get(sub.ClassroomCourseID, sub.ClassroomCourseWorkID);
      var maxPoints = cw.maxPoints || 100;
      var weights = GradeScaleService.getGradeWeights();
      var total = 0, maxRubricMarks = 0, defs = Config.CRITERIA_DEFINITIONS;
      for (var i = 0; i < defs.length; i++) {
        var cd = detail.criteriaMap[defs[i].criterionId];
        if (!cd || !cd.finalApprovedGrade || weights[cd.finalApprovedGrade] === undefined) return { success: false, message: 'Missing an approved grade for ' + defs[i].criterionId + '.' };
        total += defs[i].maxMarks * weights[cd.finalApprovedGrade];
        maxRubricMarks += defs[i].maxMarks;
      }
      var assignedGrade = Math.round((total / maxRubricMarks) * maxPoints * 10) / 10;
      Classroom.Courses.CourseWork.StudentSubmissions.patch({ assignedGrade: assignedGrade, draftGrade: assignedGrade }, sub.ClassroomCourseID, sub.ClassroomCourseWorkID, sub.ClassroomSubmissionID, { updateMask: 'assignedGrade,draftGrade' });

      var now = Utils.formatDate(new Date());
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sSheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
      var row = Sheets.findRowByKey(sSheet, 'SubmissionRecordID', submissionRecordId);
      if (row !== -1) {
        Sheets.setCellByHeader(sSheet, row, 'ClassroomAssignedGrade', assignedGrade);
        Sheets.setCellByHeader(sSheet, row, 'LastClassroomSyncAt', now);
        Sheets.setCellByHeader(sSheet, row, 'LastSyncResult', 'Synced ' + assignedGrade + '/' + maxPoints);
      }
      Logging.logHistory(submissionRecordId, assessmentId, 'ClassroomSynced', '', {},
        { assignedGrade: assignedGrade, maxPoints: maxPoints }, 'Explicit teacher sync to Classroom.');
      return { success: true, message: 'Synced grade ' + assignedGrade + ' / ' + maxPoints + ' to Classroom for ' + sub.StudentName + '.' };
    } catch (err) {
      Logging.logError('syncAssessmentToClassroom', err);
      return { success: false, message: 'Classroom sync failed: ' + Utils.sanitizeError(err) };
    }
  }

  return { listTeacherCourses: listTeacherCourses, listCourseWork: listCourseWork, saveClassroomSelection: saveClassroomSelection, saveClassroomSelectionForTask: saveClassroomSelectionForTask, getActiveClassroomConfig: getActiveClassroomConfig, importOrRefreshClassroomSubmissions: importOrRefreshClassroomSubmissions, syncClassListFromClassroom: syncClassListFromClassroom, syncAssessmentToClassroom: syncAssessmentToClassroom, checkAndAutoImportDueAssignments: checkAndAutoImportDueAssignments };
})();

function checkAndAutoImportDueAssignments() {
  WebAccess.requireTeacher(); // trigger owner must be on the teacher allowlist
  return ClassroomService.checkAndAutoImportDueAssignments();
}

function installAutoImportTrigger() {
  WebAccess.requireTeacher();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAndAutoImportDueAssignments') {
      return { success: false, message: 'Auto-import trigger is already installed.' };
    }
  }
  ScriptApp.newTrigger('checkAndAutoImportDueAssignments')
    .timeBased()
    .everyHours(1)
    .create();
  return { success: true, message: 'Auto-import trigger installed successfully. It will check for due assignments every hour.' };
}

/* ============================================================================
 * SECTION 7: HIGH-ACCURACY GEMINI ENGINE (INDEPENDENT BAND TICKING)
 * ============================================================================ */
var GeminiService = (function() {
  function getApiKey() { return PropertiesService.getScriptProperties().getProperty(Config.SCRIPT_PROP_GEMINI_KEY); }
  function setGeminiApiKey(key) {
    if (!key || key.trim().length < 10) return { success: false, message: 'Invalid API key.' };
    PropertiesService.getScriptProperties().setProperty(Config.SCRIPT_PROP_GEMINI_KEY, key.trim());
    return { success: true, message: 'Gemini API key stored securely.' };
  }

  function testGeminiConnection() {
    var key = getApiKey();
    if (!key) return { success: false, message: 'No Gemini key set in Script Properties.' };
    try {
      var res = callGeminiWithFallback([{ role: 'user', parts: [{ text: 'Ping' }] }], false);
      return res.success ? { success: true, message: 'Connected to Gemini successfully using model: ' + res.modelUsed } : { success: false, message: 'Gemini error: ' + res.message };
    } catch (e) { return { success: false, message: 'Connection error: ' + Utils.sanitizeError(e) }; }
  }

  /* preferredModel (optional) is tried first; fallbacks follow. The result's
     `unavailable` flag is true for quota/overload/network failures, so callers
     can offer "queue for later" instead of a hard error. */
  function callGeminiWithFallback(contents, requireJson, preferredModel) {
    var key = getApiKey();
    if (!key) throw new Error('Gemini API key is not configured.');

    var first = /^[a-z0-9.\-]{3,60}$/.test(String(preferredModel || '')) ? preferredModel : Config.PRIMARY_MODEL;
    var modelsToTry = [first].concat([Config.PRIMARY_MODEL].concat(Config.FALLBACK_MODELS).filter(function (m) { return m !== first; }));
    var lastError = null;
    var unavailable = false;

    for (var m = 0; m < modelsToTry.length; m++) {
      var model = modelsToTry[m];
      /* API key travels in the x-goog-api-key header, never in the URL
         (URLs leak into logs and error strings). */
      var url = 'https://generativelanguage.googleapis.com/' + Config.GEMINI_API_VERSION + '/models/' + model + ':generateContent';
      var payload = { contents: contents };
      if (requireJson) {
        payload.generationConfig = { responseMimeType: 'application/json', temperature: 0.2 };
      }

      for (var attempt = 1; attempt <= 3; attempt++) {
        try {
          var resp = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            headers: { 'x-goog-api-key': key },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
          });

          var code = resp.getResponseCode();
          var text = resp.getContentText();

          if (code === 200) {
            var parsed = JSON.parse(text);
            var contentPart = parsed.candidates && parsed.candidates[0] && parsed.candidates[0].content && parsed.candidates[0].content.parts && parsed.candidates[0].content.parts[0] ? parsed.candidates[0].content.parts[0].text : '{}';
            return { success: true, text: contentPart, modelUsed: model };
          }

          if (code === 503 || code === 429 || code === 500) {
            unavailable = true;
            lastError = 'Model ' + model + ' busy (code ' + code + '). Attempt ' + attempt + ' failed.';
            Utilities.sleep(attempt * 1500);
            continue;
          }

          /* Never surface the raw response body or URL: upstream errors may
             echo credentials, student data, or submitted prompt fragments. */
          lastError = 'Gemini returned HTTP ' + code + ' (' + model + '). Check your key, API access, model availability and quota.';
          break;
        } catch (fetchErr) {
          unavailable = true;
          lastError = 'Unable to reach Gemini (' + model + '). Check network access.';
          Utilities.sleep(1500);
        }
      }
    }

    return { success: false, unavailable: unavailable, message: lastError || 'All models exhausted.' };
  }

  function runAiAssessmentWithPreparedFiles(submissionRecordId, preparedFiles) {
    var blocked = LegacySafety.reject(submissionRecordId) || AssessmentService.rejectFinalSubmission(submissionRecordId);
    if (blocked) return blocked;
    if (!preparedFiles || !preparedFiles.length) {
      return { success: false, message: 'No eligible files available for AI analysis.' };
    }

    var rubricContext = Config.getSerializedRubricPrompt();

    var sysPrompt =
      "You are an experienced, encouraging secondary Technological and Applied Studies (TAS) teacher in NSW, Australia assessing Year 9 Design & Technology: Jewellery Design.\n" +
      "Assess student work against the 12 criteria (C01 to C12) strictly using the official rubric below.\n\n" +
      "=== OFFICIAL RUBRIC & OBSERVABLE DEFINITIONS ===\n" +
      rubricContext + "\n" +
      "================================================\n\n" +
      "NON-NEGOTIABLE ASSESSMENT RULES:\n" +
      "1. TICK EVERY OBSERVABLE THE EVIDENCE SUPPORTS, ACROSS ALL BANDS:\n" +
      "   - Evidence for a single criterion is often uneven. A student may satisfy an A-level observable while ALSO only satisfying a D-level observable elsewhere in the same criterion.\n" +
      "   - Tick EVERY individual observable statement across Bands A, B, C, D, and E wherever the submission provides genuine evidence for it.\n" +
      "2. IGNORE BLANK / UNUSED DUPLICATE TEMPLATE SLIDES.\n" +
      "3. VISION - HAND-DRAWN SKETCH DETECTION (C08): tick relevant observables if hand-drawn sketches with dimensions/labels are present.\n" +
      "4. VISION - PRACTICAL PHYSICAL MAKING PHOTOS (C11 & C12): tick relevant observables if photographed finished jewellery is present.\n" +
      "5. NEVER TICK a Missing / Distinct override (Cxx-MD); only the teacher can select it.\n\n" +
      "Respond ONLY in valid JSON matching this schema:\n" +
      "{\n" +
      "  \"criteria\": [\n" +
      "    {\n" +
      "      \"criterionId\": \"C01\",\n" +
      "      \"tickedCheckboxes\": [\"C01-A-1\", \"C01-D-4\"],\n" +
      "      \"evidenceNotes\": [\n" +
      "        {\"checkboxId\": \"C01-A-1\", \"fileName\": \"folio.pdf\", \"note\": \"Six success criteria stated on page 4-5.\"}\n" +
      "      ]\n" +
      "    }\n" +
      "  ]\n" +
      "}";

    var parts = [{ text: sysPrompt }];
    for (var p = 0; p < preparedFiles.length; p++) {
      parts.push({ inlineData: { mimeType: preparedFiles[p].mimeType, data: preparedFiles[p].inlineData.data } });
    }

    var geminiCall = callGeminiWithFallback([{ role: 'user', parts: parts }], true);
    if (!geminiCall.success) {
      return { success: false, message: 'Gemini service busy. Details: ' + geminiCall.message };
    }

    try {
      var parsed = JSON.parse(geminiCall.text);
      var aId = Utils.generateUuid();
      var cList = parsed.criteria;
      if (!Array.isArray(cList) || !cList.length) return { success: false, message: 'Gemini returned no rubric evidence; no draft was saved.' };
      var known = {};
      Config.CRITERIA_DEFINITIONS.forEach(function (c) {
        known[c.criterionId] = {};
        ['A','B','C','D','E'].forEach(function (band) {
          (c.bands[band] || []).forEach(function (obs) { known[c.criterionId][obs[0]] = true; });
        });
      });
      var cleanProposals = [];
      for (var j = 0; j < cList.length; j++) {
        var cRes = cList[j];
        if (!cRes || !known[cRes.criterionId]) continue;
        var map = {};
        var tList = Array.isArray(cRes.tickedCheckboxes) ? cRes.tickedCheckboxes : [];
        for (var t = 0; t < tList.length; t++) {
          if (known[cRes.criterionId][tList[t]]) map[tList[t]] = true; // never auto-tick MD
        }
        cleanProposals.push({ criterionId: cRes.criterionId, checkboxes: map,
          evidence: Array.isArray(cRes.evidenceNotes) ? cRes.evidenceNotes.slice(0, 50) : [] });
      }
      if (!cleanProposals.length) return { success: false, message: 'Gemini returned no valid rubric criteria; no draft was saved.' };
      for (var k = 0; k < cleanProposals.length; k++) {
        var proposal = cleanProposals[k];
        var calc = AssessmentService.deriveGradeFromChecks(proposal.criterionId, proposal.checkboxes);
        var saved = AssessmentService.saveCriterionDraft(aId, submissionRecordId, proposal.criterionId, {
          aiProposedGrade: calc.grade, checkboxes: proposal.checkboxes,
          aiEvidence: proposal.evidence, confidence: 0.85
        });
        if (!saved.success) return saved;
      }

      var fbRes = generateTeacherReviewedFeedback(aId, submissionRecordId);
      var logWarning = '';
      try {
        var aiLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_AI_ASSESSMENTS);
        if (aiLog) aiLog.appendRow([
          aId, submissionRecordId, 'EvidenceProposal', Utils.formatDate(new Date()),
          geminiCall.modelUsed, 'legacy-unified-v2',
          Utils.safeJsonStringify(preparedFiles.map(function (f) { return { fileId: f.fileId, fileName: f.fileName }; })),
          'Legacy Jewellery evidence proposal; teacher review required', Utils.safeJsonStringify(cleanProposals),
          'Success', '', '', Utils.getSafeUserEmail(), 'Pending teacher review'
        ]);
      } catch (logErr) { logWarning = ' The AI run log could not be written; check AIAssessments.'; }

      return {
        success: true,
        message: 'AI assessment & feedback completed using ' + geminiCall.modelUsed + '.' + logWarning,
        assessmentId: aId,
        feedback: fbRes.success ? fbRes.feedback : null
      };
    } catch (err) {
      Logging.logError('runAiAssessmentWithPreparedFiles', err);
      return { success: false, message: 'AI response parsing error: ' + Utils.sanitizeError(err) };
    }
  }

  function runInitialAiAssessment(submissionRecordId) {
    var blocked = LegacySafety.reject(submissionRecordId) || AssessmentService.rejectFinalSubmission(submissionRecordId);
    if (blocked) return blocked;
    var prep = DriveService.prepareSubmissionFilesForAi(submissionRecordId);
    if (!prep.files.length) return { success: false, message: 'No eligible files available for AI analysis.' };
    /* Batch, trigger and Studio use the SAME evidence-ticking policy. AI may
       propose checkboxes but cannot compute or approve a final grade. */
    return runAiAssessmentWithPreparedFiles(submissionRecordId, prep.files);
  }

  function generateTeacherReviewedFeedback(assessmentId, submissionRecordId) {
    var detail = AssessmentService.getSubmissionForMarking(submissionRecordId);
    var summary = [];
    Config.CRITERIA_DEFINITIONS.forEach(function(c) {
      var cd = detail.criteriaMap[c.criterionId] || {};
      summary.push(c.criterionId + ' (' + c.title + '): Grade ' + (cd.derivedGrade || 'E') + ' ' + (cd.teacherWrittenNote ? '| ' + cd.teacherWrittenNote : ''));
    });

    var prompt =
      "You are a supportive, knowledgeable Australian high school TAS teacher writing warm, natural, and constructive report comments for Year 9 Design & Technology student " + detail.submission.StudentName + ".\n\n" +
      "Below is the marking breakdown for their Jewellery Design project:\n" + summary.join('\n') + "\n\n" +
      "CRITICAL VOICE & TONE GUIDELINES:\n" +
      "- Speak naturally like a real teacher talking directly to a Year 9 student.\n" +
      "- Do NOT use robotic, corporate AI jargon (avoid 'commendable', 'testament to', 'pivotal', 'delve into', 'unprecedented').\n" +
      "- Reference real, specific materials and techniques from their portfolio (e.g. FIMO polymer clay, wire wrapping, silver or gold leaf, pliers, bead threading).\n" +
      "- Provide genuine praise for what worked, and clear, practical workshop advice on what to improve next time.\n\n" +
      "Return ONLY JSON matching this structure:\n" +
      "{\n" +
      "  \"whatWentWell\": \"...\",\n" +
      "  \"areasForImprovement\": \"...\",\n" +
      "  \"goalsForNextAssessment\": \"1. ...\\n2. ...\"\n" +
      "}";

    var res = callGeminiWithFallback([{ role: 'user', parts: [{ text: prompt }] }], true);
    if (!res.success) return { success: false, message: 'Feedback generation error: ' + res.message };

    try {
      var out = JSON.parse(res.text);
      var aiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_AI_ASSESSMENTS);
      if (aiSheet) {
        aiSheet.appendRow([
          assessmentId, submissionRecordId, 'FeedbackGeneration', Utils.formatDate(new Date()),
          Config.PRIMARY_MODEL, 'v1', '[]', 'Auto-generated feedback', res.text, 'Success', '', 1.0,
          Utils.getSafeUserEmail(), 'Accepted'
        ]);
      }
      return { success: true, feedback: out };
    } catch (err) {
      return { success: false, message: 'Invalid feedback JSON returned.' };
    }
  }

  return {
  setGeminiApiKey: setGeminiApiKey,
  testGeminiConnection: testGeminiConnection,
  callGeminiWithFallback: callGeminiWithFallback,
  runInitialAiAssessment: runInitialAiAssessment,
  runAiAssessmentWithPreparedFiles: runAiAssessmentWithPreparedFiles,
  generateTeacherReviewedFeedback: generateTeacherReviewedFeedback
};
})();

/* ============================================================================
 * SECTION 8: LEGACY JEWELLERY PDF TEMPLATE
 * Only used by the original spreadsheet dialog. The task-agnostic web report
 * is generated by RubricEngine.renderReportHtml instead.
 * ============================================================================ */
var HtmlTemplates = (function() {

  function getAssessmentPdfReportHtml(detail) {
    var sub = detail.submission;
    var cMap = detail.criteriaMap;
    var fb = detail.feedback || {};
    var defs = Config.CRITERIA_DEFINITIONS;

    var partA = 0, partB = 0;
    var rowsHtml = '';
    var escape = RubricEngine.escapeHtml;
    var weights = GradeScaleService.getGradeWeights();

    for (var i = 0; i < defs.length; i++) {
      var c = defs[i];
      var cd = cMap[c.criterionId] || {};
      var grade = cd.finalApprovedGrade || cd.derivedGrade || 'E';
      var weight = weights[grade];
      if (weight === undefined) throw new Error('Unknown approved grade: ' + grade);
      var mark = (c.maxMarks * weight).toFixed(2);

      if (c.part === 'Part A') partA += parseFloat(mark);
      else partB += parseFloat(mark);

      var note = cd.teacherWrittenNote || cd.teacherAudioTranscript || cd.explanation || 'Thresholds verified.';

      rowsHtml += '<tr>' +
        '<td style="font-weight:bold;padding:6px;border:1px solid #cbd5e1;">' + escape(c.criterionId) + '</td>' +
        '<td style="padding:6px;border:1px solid #cbd5e1;"><b>' + escape(c.title) + '</b><br/><span style="font-size:10px;color:#64748b;">' + escape(c.part) + ' &bull; /' + c.maxMarks + ' Marks</span></td>' +
        '<td style="text-align:center;font-weight:bold;padding:6px;border:1px solid #cbd5e1;">' + escape(grade) + '</td>' +
        '<td style="text-align:center;padding:6px;border:1px solid #cbd5e1;">' + mark + ' / ' + c.maxMarks + '</td>' +
        '<td style="padding:6px;font-size:11px;border:1px solid #cbd5e1;">' + escape(note) + '</td>' +
      '</tr>';
    }

    var total = (partA + partB).toFixed(1);
    var overallGrade = GradeScaleService.getOverallGradeLetter(parseFloat(total), 
    GradeScaleService.getOverallGradeBands());

    return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<style>' +
      'body{font-family:Arial,Helvetica,sans-serif;margin:20px;color:#0f172a;line-height:1.3;font-size:12px;}' +
      'h1{margin:0;font-size:18px;color:#1e3a8a;}' +
      '.header-box{display:flex;justify-content:space-between;border-bottom:2px solid #1e3a8a;padding-bottom:10px;margin-bottom:14px;}' +
      '.score-card{background:#0f172a;color:#fff;padding:8px 14px;border-radius:6px;text-align:right;}' +
      '.details-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;background:#f8fafc;border:1px solid #cbd5e1;padding:8px 12px;border-radius:6px;margin-bottom:14px;font-size:11px;}' +
      'table{width:100%;border-collapse:collapse;margin-bottom:16px;}' +
      'th{background:#0f172a;color:#fff;padding:6px;font-size:11px;border:1px solid #0f172a;text-align:left;}' +
      '.fb-box{background:#f8fafc;border:1px solid #2563eb;border-radius:6px;padding:10px 14px;margin-top:10px;}' +
      '.fb-title{font-weight:bold;color:#1e3a8a;margin-bottom:4px;font-size:11px;}' +
      '</style></head><body>' +
      '<div class="header-box">' +
        '<div>' +
          '<h1>NSW STAGE 5 DESIGN &amp; TECHNOLOGY</h1>' +
          '<div style="font-size:13px;font-weight:bold;color:#2563eb;margin-top:2px;">Year 9 Jewellery Design &mdash; Assessment &amp; Folio Report</div>' +
        '</div>' +
        '<div class="score-card">' +
          '<div style="font-size:10px;color:#94a3b8;">OVERALL RESULT</div>' +
          '<div style="font-size:18px;font-weight:bold;color:#38bdf8;">' + total + ' / 100 (' + overallGrade + ')</div>' +
        '</div>' +
      '</div>' +
      '<div class="details-grid">' +
        '<div><b>Student:</b> ' + escape(sub.StudentName) + '</div>' +
        '<div><b>Class:</b> ' + escape(sub.Class || '9DAT1') + '</div>' +
        '<div><b>Submission Version:</b> v' + escape(sub.SubmissionVersion) + '</div>' +
        '<div><b>Part A (Proposal &amp; Realisation):</b> ' + partA.toFixed(2) + ' / 55</div>' +
        '<div><b>Part B (Creativity &amp; Making):</b> ' + partB.toFixed(2) + ' / 45</div>' +
        '<div><b>Approval Date:</b> ' + escape(String(sub.ApprovedAt || '').split('T')[0]) + '</div>' +
      '</div>' +
      '<table>' +
        '<thead><tr><th>Code</th><th>Criterion Title &amp; Component</th><th style="text-align:center;">Band</th><th style="text-align:center;">Mark</th><th>Observables &amp; Evidence Audit</th></tr></thead>' +
        '<tbody>' + rowsHtml + '</tbody>' +
      '</table>' +
      '<div class="fb-box">' +
        '<div style="font-size:13px;font-weight:bold;color:#0f172a;margin-bottom:8px;border-bottom:1px solid #cbd5e1;padding-bottom:4px;">Teacher Feedback &amp; Future Goals</div>' +
        '<div class="fb-title">What You Did Well:</div>' +
        '<div style="margin-bottom:8px;white-space:pre-wrap;">' + escape(fb.whatWentWell || '') + '</div>' +
        '<div class="fb-title">Areas for Improvement:</div>' +
        '<div style="margin-bottom:8px;white-space:pre-wrap;">' + escape(fb.areasForImprovement || '') + '</div>' +
        '<div class="fb-title">Goals for Next Assessment:</div>' +
        '<div style="white-space:pre-wrap;">' + escape(fb.goalsForNextAssessment || '') + '</div>' +
      '</div>' +
      '</body></html>';
  }

  return {
    getAssessmentPdfReportHtml: getAssessmentPdfReportHtml
  };
})();


/* ============================================================================
 * SECTION 9: TOP-LEVEL ENTRY POINTS & RPC API ENDPOINTS
 * ============================================================================ */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Assessment System')
    .addItem('Open Assessment Studio (Fullscreen)', 'openAssessmentStudio')
    .addItem('Run Class AI Grading (All Students)', 'openClassAiGradingRunner')
    .addItem('Setup New Grading Workbook', 'openGradingWorkbookSetupUi')
    .addItem('Manage Shared Rubrics', 'openManageRubricsUi')
    .addItem('Setup / Migrate Assessment System', 'setupOrMigrateAssessmentSystemUi')
    .addItem('Refresh Marking Sheet', 'refreshMarkingSheetUi')
    .addSeparator()
    .addSubMenu(ui.createMenu('Classroom')
      .addItem('Select Course and Assignment', 'openClassroomPicker')
      .addItem('Import / Refresh Submissions', 'importOrRefreshClassroomSubmissionsUi')
      .addItem('Sync Class List', 'syncClassListFromClassroomUi')
      .addItem('Sync Approved Assessment to Classroom', 'syncApprovedAssessmentToClassroomUi'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Gemini')
      .addItem('Set Gemini API Key', 'setGeminiApiKeyDialog')
      .addItem('Test Gemini Connection', 'testGeminiConnectionUi'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Administration')
      .addItem('Enable Hourly Auto-Import', 'installAutoImportTriggerUi')
      .addItem('Create Reassessment', 'createReassessmentUi')
      .addItem('View Error Log', 'viewErrorLogUi')
      .addItem('Verify Config & Sheets (Diagnostic)', 'verifySubmissionSystem')
      .addItem('Repair System & Re-Import (Diagnostic)', 'repairAssessmentSystemAndImport'))
    .addToUi();
}

function openAssessmentStudio() {
  WebAccess.requireTeacher();
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_MARKING)) {
    SpreadsheetApp.getUi().alert('Use the Lesson Grader web app for project-specific rubrics. This Sheets Studio is for the legacy Jewellery markbook only.');
    return;
  }
  var template = HtmlService.createTemplateFromFile('AssessmentStudio');
  var html = template.evaluate()
    .setWidth(1600)
    .setHeight(1000)
    .setTitle('AI Assessment Studio — Year 9 Jewellery Design');
  SpreadsheetApp.getUi().showModalDialog(html, 'AI Assessment Studio — Year 9 Jewellery Design');
}

function openGradingWorkbookSetupUi() {
  WebAccess.requireTeacher();
  var template = HtmlService.createTemplateFromFile('GradingWorkbookSetup');
  var html = template.evaluate()
    .setWidth(450)
    .setHeight(500)
    .setTitle('Setup Grading Workbook');
  SpreadsheetApp.getUi().showModalDialog(html, 'Setup Grading Workbook');
}

function openManageRubricsUi() {
  WebAccess.requireTeacher();
  var template = HtmlService.createTemplateFromFile('ManageRubrics');
  var html = template.evaluate()
    .setWidth(500)
    .setHeight(400)
    .setTitle('Manage Shared Rubrics');
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Shared Rubrics');
}

function openClassAiGradingRunner() {
  WebAccess.requireTeacher();
  var template = HtmlService.createTemplateFromFile('ClassBatchGrading');
  var html = template.evaluate()
    .setWidth(680)
    .setHeight(420)
    .setTitle('Class-Wide AI Grading Runner');
  SpreadsheetApp.getUi().showModalDialog(html, 'Class-Wide AI Grading Runner');
}

function openClassroomPicker() {
  WebAccess.requireTeacher();
  var template = HtmlService.createTemplateFromFile('CourseworkPicker');
  template.passedCourseCode = '';
  template.passedTaskName = '';
  var html = template.evaluate()
    .setWidth(620)
    .setHeight(480)
    .setTitle('Select Classroom Course and Assignment');
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Classroom Course and Assignment');
}

function openClassroomPickerForTask(courseCode, taskName) {
  WebAccess.requireTeacher();
  var template = HtmlService.createTemplateFromFile('CourseworkPicker');
  template.passedCourseCode = courseCode;
  template.passedTaskName = taskName;
  var html = template.evaluate()
    .setWidth(620)
    .setHeight(480)
    .setTitle('Link Classroom to Task');
  SpreadsheetApp.getUi().showModalDialog(html, 'Link Classroom to Task');
}

function openVoiceDictation(criterionId) {
  WebAccess.requireTeacher();
  var template = HtmlService.createTemplateFromFile('VoiceDictation');
  template.criterionId = criterionId || 'C01';
  var html = template.evaluate()
    .setWidth(420)
    .setHeight(360)
    .setTitle('Voice Dictation');
  SpreadsheetApp.getUi().showModalDialog(html, 'Dictate Observation Notes');
}

function setupOrMigrateAssessmentSystemUi() {
  WebAccess.requireTeacher();
  var res = Sheets.setupOrMigrateAssessmentSystem();
  SpreadsheetApp.getUi().alert('Setup Status', res.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function refreshMarkingSheetUi() {
  WebAccess.requireTeacher();
  Sheets.verifyMarkingSheetFormulas(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Success', 'Formulas refreshed on 9DAT1 Marking.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setGeminiApiKeyDialog() {
  WebAccess.requireTeacher();
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Set Gemini API Key', 'Paste your Google AI Studio Gemini API Key', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) {
    var out = GeminiService.setGeminiApiKey(res.getResponseText());
    ui.alert(out.success ? 'Success' : 'Error', out.message, ui.ButtonSet.OK);
  }
}

function testGeminiConnectionUi() {
  WebAccess.requireTeacher();
  var out = GeminiService.testGeminiConnection();
  SpreadsheetApp.getUi().alert(out.success ? 'Success' : 'Failed', out.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* This now displays the REAL result message from the fixed import function. */
function importOrRefreshClassroomSubmissionsUi() {
  WebAccess.requireTeacher();
  var res = ClassroomService.importOrRefreshClassroomSubmissions();
  SpreadsheetApp.getUi().alert(
    res.success ? 'Classroom Import Complete' : 'Classroom Import Failed',
    res.message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function syncClassListFromClassroomUi() {
  WebAccess.requireTeacher();
  var cfg = ClassroomService.getActiveClassroomConfig();
  if (!cfg) {
    SpreadsheetApp.getUi().alert('Error', 'No active Classroom assignment configured. Please select one via Classroom menu.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var res = ClassroomService.syncClassListFromClassroom(cfg.CourseID);
  SpreadsheetApp.getUi().alert(
    res.success ? 'Class List Sync Complete' : 'Class List Sync Failed',
    res.message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function syncApprovedAssessmentToClassroomUi() {
  WebAccess.requireTeacher();
  SpreadsheetApp.getUi().alert('Classroom Sync', 'Please open Assessment Studio and click "Sync" for the desired student.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function installAutoImportTriggerUi() {
  WebAccess.requireTeacher();
  var out = installAutoImportTrigger();
  SpreadsheetApp.getUi().alert(out.success ? 'Success' : 'Notice', out.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function createReassessmentUi() {
  WebAccess.requireTeacher();
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Create Reassessment', 'Enter parent SubmissionRecordID:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) {
    var out = AssessmentService.createReassessmentFromSubmission(res.getResponseText().trim());
    ui.alert(out.success ? 'Success' : 'Error', out.message, ui.ButtonSet.OK);
  }
}

function viewErrorLogUi() {
  WebAccess.requireTeacher();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(Config.SHEET_ERROR_LOG);
  if (sh) { sh.showSheet(); ss.setActiveSheet(sh); }
}

/* --- Diagnostic helpers (safe to keep permanently) --- */
function verifySubmissionSystem() {
  WebAccess.requireTeacher();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = {
    spreadsheet: ss.getName(),
    resolvedSubmissionsSheetName: Config.SHEET_SUBMISSIONS,
    submissionsSheetExists: !!ss.getSheetByName(Config.SHEET_SUBMISSIONS),
    submissionRows: ss.getSheetByName(Config.SHEET_SUBMISSIONS)
      ? Math.max(0, ss.getSheetByName(Config.SHEET_SUBMISSIONS).getLastRow() - 1)
      : 0,
    activeClassroomConfig: ClassroomService.getActiveClassroomConfig()
  };
  Logger.log(JSON.stringify(report, null, 2));
  SpreadsheetApp.getUi().alert('Diagnostic Report', JSON.stringify(report, null, 2), SpreadsheetApp.getUi().ButtonSet.OK);
  return report;
}

function repairAssessmentSystemAndImport() {
  WebAccess.requireTeacher();
  Sheets.setupOrMigrateAssessmentSystem();
  var result = ClassroomService.importOrRefreshClassroomSubmissions();
  Logger.log(JSON.stringify(result, null, 2));
  SpreadsheetApp.getUi().alert(
    result.success ? 'Repair & Import Complete' : 'Repair & Import Failed',
    result.message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

/* --- Client-Callable RPC Endpoints (google.script.run) --- */
function apiRegenerateRubricSchema(rubricFileId) {
  WebAccess.requireTeacher();
  if (!rubricFileId) return { success: false, message: 'Invalid rubric ID.' };

  try {
    var ss = SpreadsheetApp.openById(rubricFileId);
    var rubricSheet = ss.getSheetByName('Rubric');
    if (!rubricSheet) return { success: false, message: 'Rubric tab not found in file.' };

    var data = rubricSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'Rubric tab is empty.' };

    var headers = data[0].map(function(h) { return String(h).trim(); });
    var expected = ['Criterion', 'Part', 'Section', 'MaxMarks', 'Outcome', 'Band', 'Description'];
    for (var i = 0; i < expected.length; i++) {
      if (headers.indexOf(expected[i]) === -1) {
        return { success: false, message: 'Rubric tab missing required header: ' + expected[i] };
      }
    }

    var cIdx = headers.indexOf('Criterion');
    var pIdx = headers.indexOf('Part');
    var sIdx = headers.indexOf('Section');
    var mIdx = headers.indexOf('MaxMarks');
    var oIdx = headers.indexOf('Outcome');
    var bIdx = headers.indexOf('Band');
    var dIdx = headers.indexOf('Description');

    var criteriaObj = {};
    var cCount = 1;
    for (var r = 1; r < data.length; r++) {
      var critTitle = String(data[r][cIdx]).trim();
      var desc = String(data[r][dIdx]).trim();
      var band = String(data[r][bIdx]).trim();
      if (!critTitle || !desc || !band) continue;

      if (!criteriaObj[critTitle]) {
        var cId = 'C' + (cCount < 10 ? '0' : '') + cCount;
        criteriaObj[critTitle] = {
          id: cId,
          part: data[r][pIdx],
          section: data[r][sIdx],
          maxMarks: data[r][mIdx],
          outcome: data[r][oIdx],
          bands: {}
        };
        cCount++;
      }
      var cData = criteriaObj[critTitle];
      if (!cData.bands[band]) cData.bands[band] = [];
      cData.bands[band].push(desc);
    }

    var rows = [];
    rows.push(['CriterionID', 'CriterionTitle', 'Part', 'Section', 'MaxMarks', 'Outcome', 'Band', 'ThresholdType', 'RequiredCount', 'CheckboxID', 'CheckboxLabel', 'EvidenceFocus', 'IsMissingDistinctOverride']);

    for (var title in criteriaObj) {
      var c = criteriaObj[title];
      var mdId = c.id + '-MD';
      var mdText = "Missing / Distinct: exceptionally refined evidence.";
      rows.push([c.id, title, c.part, c.section, c.maxMarks, c.outcome, 'A', 'Override', 1, mdId, mdText, 'Distinct Evidence', true]);

      var bandsOrder = ['A', 'B', 'C', 'D', 'E'];
      for (var b = 0; b < bandsOrder.length; b++) {
        var bName = bandsOrder[b];
        var items = c.bands[bName] || [];
        var req = (bName === 'A') ? items.length : Math.ceil(items.length * 2 / 3);
        for (var k = 0; k < items.length; k++) {
          var cbId = c.id + '-' + bName + '-' + (k + 1);
          rows.push([c.id, title, c.part, c.section, c.maxMarks, c.outcome, bName, (bName === 'A' ? 'All' : 'TwoThirds'), req, cbId, items[k], 'Observable', false]);
        }
      }
    }

    if (rows.length <= 1) return { success: false, message: 'No valid data found to generate criteria.' };

    var criteriaSheet = ss.getSheetByName('Criteria');
    if (!criteriaSheet) {
      criteriaSheet = ss.insertSheet('Criteria');
      criteriaSheet.hideSheet();
    }

    criteriaSheet.clear();
    criteriaSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    criteriaSheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');

    return { success: true, message: 'Successfully generated ' + (rows.length - 1) + ' criteria rows.' };

  } catch (err) {
    return { success: false, message: 'Error generating schema: ' + err.message };
  }
}

function apiListSharedRubrics() {
  WebAccess.requireTeacher();
  var sharedRubrics = [];
  try {
    RubricsLibrary.ensureRubricsLibraryStructure();
    var root = DriveApp.getRootFolder();
    var libFolders = root.getFoldersByName(RubricsLibrary.ROOT_FOLDER_NAME || 'Graded Assessments');
    if (libFolders.hasNext()) {
      var sharedFolders = libFolders.next().getFoldersByName('Shared Rubrics');
      if (sharedFolders.hasNext()) {
        var files = sharedFolders.next().getFilesByType(MimeType.GOOGLE_SHEETS);
        while (files.hasNext()) {
          var f = files.next();
          sharedRubrics.push({ id: f.getId(), name: f.getName() });
        }
      } else {
        // No rubrics folder yet, return empty list instead of failing
        Logger.log("Shared Rubrics folder missing.");
      }
    } else {
      Logger.log("Root library folder missing.");
    }
  } catch(e) {
    // Keep logger, return empty array without breaking UI
    Logger.log("Error fetching shared rubrics: " + e.message);
  }
  return sharedRubrics;
}

function apiGetGradingSetupData() {
  WebAccess.requireTeacher();
  var stages = ['Stage 4', 'Stage 5', 'Stage 6'];
  var courses = [];
  var categories = [];
  var tasks = [];
  var sharedRubrics = [];

  for (var i = 0; i < stages.length; i++) {
    var stageCourses = RubricsLibrary.getActiveCourses(stages[i]);
    courses = courses.concat(stageCourses);
  }

  // Iterate over all courses to fetch nested data
  for (var i = 0; i < courses.length; i++) {
    var c = courses[i];
    var courseCats = RubricsLibrary.getCategoriesForCourse(c.courseCode);
    categories = categories.concat(courseCats);

    for (var j = 0; j < courseCats.length; j++) {
      var cat = courseCats[j];
      var catTasks = RubricsLibrary.getTasksForCourseCategory(cat.courseCode, cat.categoryName);
      tasks = tasks.concat(catTasks);
    }
  }

  try {
    sharedRubrics = apiListSharedRubrics();
  } catch (e) {
    Logger.log("apiGetGradingSetupData: " + e.message);
    // Continue execution to return courses/categories/tasks even if rubrics fail,
    // letting the UI render but without a populated rubrics list.
  }

  return { courses: courses, categories: categories, tasks: tasks, sharedRubrics: sharedRubrics };
}

function deriveStageFromCourseCode(courseCode) {
  var code = String(courseCode || '').trim();

  if (code.indexOf('7') === 0 || code.indexOf('8') === 0) {
    return 'Stage 4';
  }

  if (code.indexOf('9') === 0 || code.indexOf('10') === 0) {
    return 'Stage 5';
  }

  if (code.indexOf('11') === 0 || code.indexOf('12') === 0) {
    return 'Stage 6';
  }

  return null;
}

function apiSubmitGradingWorkbookSetup(
  courseCode,
  categoryName,
  taskName,
  year,
  rubricSelection,
  newRubricName,
  rubricDocument
) {
  WebAccess.requireTeacher();
  var derivedStage = deriveStageFromCourseCode(courseCode);

  if (!derivedStage) {
    return {
      success: false,
      message: 'Unable to determine Stage from Course Code. Course Code must begin with 7, 8, 9, 10, 11, or 12.'
    };
  }

  // 2. Handle Rubric (Create new or use existing)
  var rubricFileId = "";
  if (rubricSelection === "__UPLOAD__") {
    var createdRubric = apiCreateRubricFromDocument(
      newRubricName,
      rubricDocument
    );
    if (!createdRubric.success) {
      return createdRubric; // Fail early if parsing/upload fails
    }
    rubricFileId = createdRubric.rubricFileId;
  } else if (rubricSelection && rubricSelection !== "__UPLOAD__") {
    rubricFileId = rubricSelection;
  }

  // 3. Call ensureGradingWorkbook
  var result = RubricsLibrary.ensureGradingWorkbook(
    derivedStage,
    courseCode,
    categoryName,
    year,
    taskName
  );
  if (!result.success) {
    return result; // Bubble up error
  }

  // 4. Open workbook and ensure task sheet exists
  try {
    var ss = SpreadsheetApp.openById(result.workbook.id);
    var sheet = ss.getSheetByName(taskName);
    if (!sheet) {
      ss.insertSheet(taskName);
    }

    // Clean up default "Sheet1" if it exists and we added a new task sheet
    var sheet1 = ss.getSheetByName("Sheet1");
    if (sheet1 && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet1);
    }

    // 5. Add task to index
    RubricsLibrary.addTaskToIndex(courseCode, categoryName, taskName, rubricFileId);

    return {
      success: true,
      message: result.message + ' Task tab "' + taskName + '" is ready.',
      workbook: result.workbook,
      rubricFileId: rubricFileId
    };
  } catch (err) {
    return { success: false, message: 'Failed to setup task tab: ' + err.message };
  }
}

function apiCreateRubricFromDocument(newRubricName, rubricDocument) {
  WebAccess.requireTeacher();
  if (!newRubricName || !String(newRubricName).trim()) {
    return { success: false, message: 'A rubric name is required.' };
  }

  if (!rubricDocument || !rubricDocument.base64 || !rubricDocument.mimeType) {
    return { success: false, message: 'A PDF or Markdown rubric document is required.' };
  }

  var mimeType = String(rubricDocument.mimeType).toLowerCase();
  var allowedMimeTypes = [
    'application/pdf',
    'text/markdown',
    'text/plain'
  ];

  if (allowedMimeTypes.indexOf(mimeType) === -1) {
    return {
      success: false,
      message: 'Only PDF, Markdown, and text rubric documents are supported.'
    };
  }

  var EXTRACTION_PROMPT =
    "You are converting a school assessment rubric into structured data.\n" +
    "Return ONLY valid JSON. Do not include Markdown fences or explanation.\n\n" +
    "Return this exact shape:\n" +
    "{\n" +
    "  \"criteria\": [\n" +
    "    {\n" +
    "      \"criterion\": \"Human-readable criterion title\",\n" +
    "      \"part\": \"Optional part label or empty string\",\n" +
    "      \"section\": \"Optional section label or empty string\",\n" +
    "      \"maxMarks\": 10,\n" +
    "      \"outcome\": \"Optional outcome code or empty string\",\n" +
    "      \"bands\": {\n" +
    "        \"A\": [\"observable statement 1\"],\n" +
    "        \"B\": [\"observable statement 1\"],\n" +
    "        \"C\": [\"observable statement 1\"],\n" +
    "        \"D\": [\"observable statement 1\"],\n" +
    "        \"E\": [\"observable statement 1\"]\n" +
    "      }\n" +
    "    }\n" +
    "  ]\n" +
    "}\n\n" +
    "Rules:\n" +
    "- Extract every criterion and every observable statement actually present.\n" +
    "- Preserve the source wording as closely as possible.\n" +
    "- Use A, B, C, D, E only when those bands occur in the source.\n" +
    "- Use [] for an absent band.\n" +
    "- maxMarks must be numeric; use 0 only if the source gives no mark value.\n" +
    "- Do not invent outcomes, marks, criteria, descriptors, or observables.\n" +
    "- Do not merge distinct observable statements.";

  var geminiResponse = GeminiService.callGeminiWithFallback([
    {
      role: 'user',
      parts: [
        { text: EXTRACTION_PROMPT },
        {
          inlineData: {
            mimeType: mimeType,
            data: rubricDocument.base64
          }
        }
      ]
    }
  ], true);

  if (!geminiResponse.success) {
    return { success: false, message: 'Gemini extraction failed: ' + geminiResponse.message };
  }

  var parsed;
  try {
    parsed = JSON.parse(geminiResponse.text);
  } catch(e) {
    return { success: false, message: 'Invalid JSON returned from Gemini.' };
  }

  // Validate parsed JSON before creating a spreadsheet
  if (!parsed.criteria || !Array.isArray(parsed.criteria) || parsed.criteria.length === 0) {
    return { success: false, message: 'Validation failed: parsed.criteria must be a non-empty array.' };
  }

  var rowsToWrite = [];
  var hasAValidCriterion = false;

  for (var i = 0; i < parsed.criteria.length; i++) {
    var c = parsed.criteria[i];
    if (!c.criterion || typeof c.criterion !== 'string' || !c.criterion.trim()) {
      return { success: false, message: 'Validation failed: Each criterion needs a non-empty criterion string.' };
    }
    if (!c.bands || typeof c.bands !== 'object') {
      return { success: false, message: 'Validation failed: Each criterion needs a bands object.' };
    }

    var keys = Object.keys(c.bands);
    for (var k = 0; k < keys.length; k++) {
      var bandName = keys[k];
      var statements = c.bands[bandName];

      if (!Array.isArray(statements)) {
         return { success: false, message: 'Validation failed: Each present band value must be an array.' };
      }

      for (var s = 0; s < statements.length; s++) {
        var stmt = statements[s];
        if (!stmt || typeof stmt !== 'string' || !stmt.trim()) {
           return { success: false, message: 'Validation failed: Each present band value must be an array of non-empty strings.' };
        }

        hasAValidCriterion = true;
        rowsToWrite.push([
          c.criterion,
          c.part || '',
          c.section || '',
          c.maxMarks || 0,
          c.outcome || '',
          bandName,
          stmt
        ]);
      }
    }
  }

  if (!hasAValidCriterion) {
     return { success: false, message: 'Validation failed: No valid criteria/bands found in extraction.' };
  }

  // Schema validated. Proceed to create Drive file.
  try {
    RubricsLibrary.ensureRubricsLibraryStructure();

    var root = DriveApp.getRootFolder();
    var libFolders = root.getFoldersByName(RubricsLibrary.ROOT_FOLDER_NAME || 'Graded Assessments');
    if (!libFolders.hasNext()) return { success: false, message: 'Root library folder not found.' };
    var libFolder = libFolders.next();

    var sharedFolders = libFolder.getFoldersByName('Shared Rubrics');
    if (!sharedFolders.hasNext()) return { success: false, message: 'Shared Rubrics folder not found.' };
    var sharedRubricsFolder = sharedFolders.next();

    var cleanName = String(newRubricName).trim();
    var newRubricSS = SpreadsheetApp.create(cleanName);
    var newRubricFileId = newRubricSS.getId();

    var rubricSheet = newRubricSS.insertSheet('Rubric');
    rubricSheet.appendRow(['Criterion', 'Part', 'Section', 'MaxMarks', 'Outcome', 'Band', 'Description']);

    if (rowsToWrite.length > 0) {
      rubricSheet.getRange(2, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
    }
    rubricSheet.getRange(1, 1, 1, 7).setFontWeight('bold');

    var criteriaSheet = newRubricSS.insertSheet('Criteria');
    criteriaSheet.appendRow([
      'CriterionID', 'CriterionTitle', 'Part', 'Section', 'MaxMarks', 'Outcome',
      'Band', 'ThresholdType', 'RequiredCount', 'CheckboxID', 'CheckboxLabel',
      'EvidenceFocus', 'IsMissingDistinctOverride'
    ]);
    criteriaSheet.getRange(1, 1, 1, 13).setFontWeight('bold');

    newRubricSS.setActiveSheet(rubricSheet);
    criteriaSheet.hideSheet();

    var sheet1 = newRubricSS.getSheetByName('Sheet1');
    if (sheet1) newRubricSS.deleteSheet(sheet1);

    var newRubricFile = DriveApp.getFileById(newRubricFileId);
    newRubricFile.moveTo(sharedRubricsFolder);

    var regenResult = apiRegenerateRubricSchema(newRubricFileId);
    if (!regenResult.success) {
      newRubricFile.setTrashed(true);
      return { success: false, message: 'Schema generation failed during creation: ' + regenResult.message };
    }

    return {
      success: true,
      rubricFileId: newRubricFileId,
      rubricName: cleanName,
      message: 'Created rubric from uploaded document: ' + cleanName
    };

  } catch(e) {
     return { success: false, message: 'File creation failed: ' + e.message };
  }
}

/* The original Sheets dialog and its fixed 12-column markbook are
   Jewellery-only. Refuse other tasks rather than applying the wrong rubric. */
var LegacySafety = (function () {
  function isJewellery(subId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss || !ss.getSheetByName(Config.SHEET_MARKING)) return false;
    var rows = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].SubmissionRecordID) === String(subId)) return /jewell/i.test(String(rows[i].Task || ''));
    }
    return false;
  }
  function reject(subId) {
    return isJewellery(subId) ? null : {
      success: false,
      message: 'The Sheets Studio has only the Jewellery rubric. Activate a task-specific rubric and grade this project in the web app.'
    };
  }
  return { isJewellery: isJewellery, reject: reject };
})();

function apiGetBootstrapData(id) {
  WebAccess.requireTeacher(); return AssessmentService.getAssessmentStudioBootstrapData(id); }
function apiGetSubmissionDetail(id) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(id);
  if (block) return block;
  return AssessmentService.getSubmissionForMarking(id);
}
function apiSaveCriterionDraft(aId, sId, cId, st) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(sId);
  if (block) return block; return AssessmentService.saveCriterionDraft(aId, sId, cId, st); }
function apiSaveAssessmentDraft(aId, sId, cMap, fb) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(sId);
  if (block) return block; return AssessmentService.saveAssessmentDraft(aId, sId, cMap, fb); }
function apiRunInitialAiAssessment(sId) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(sId);
  if (block) return block; return GeminiService.runInitialAiAssessment(sId); }
function apiGenerateFeedback(aId, sId) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(sId);
  if (block) return block; return GeminiService.generateTeacherReviewedFeedback(aId, sId); }
function apiApproveAssessment(aId, sId, fb) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(sId);
  if (block) return block; return AssessmentService.approveAssessment(aId, sId, fb); }
function apiLockAssessment(aId, sId) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(sId);
  if (block) return block; return AssessmentService.lockAssessment(aId, sId); }
function apiCreateReassessment(sId) {
  WebAccess.requireTeacher(); return AssessmentService.createReassessmentFromSubmission(sId); }
function apiSyncToClassroom(aId, sId) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(sId);
  if (block) return block; return ClassroomService.syncAssessmentToClassroom(aId, sId); }
function apiListTeacherCourses() {
  WebAccess.requireTeacher(); return ClassroomService.listTeacherCourses(); }
function apiListCourseWork(cid) {
  WebAccess.requireTeacher(); return ClassroomService.listCourseWork(cid); }
function apiSaveClassroomSelection(cid, cwid) {
  WebAccess.requireTeacher(); return ClassroomService.saveClassroomSelection(cid, cwid); }
function apiSaveClassroomSelectionForTask(cid, cwid, taskName) {
  WebAccess.requireTeacher(); return ClassroomService.saveClassroomSelectionForTask(cid, cwid, taskName); }
function apiImportClassroomSubmissions() {
  WebAccess.requireTeacher(); return ClassroomService.importOrRefreshClassroomSubmissions(); }

function apiPrepareSingleFileForAi(submissionRecordId, driveFileId) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(submissionRecordId);
  if (block) return block;
  var allFiles = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSION_FILES);
  for (var i = 0; i < allFiles.length; i++) {
    if (allFiles[i].SubmissionRecordID === submissionRecordId && allFiles[i].DriveFileID === driveFileId) {
      return DriveService.prepareSingleFileForAi(allFiles[i]);
    }
  }
  return { fileId: driveFileId, fileName: '', status: 'unreadable' };
}

function apiRunAiAssessmentWithFiles(submissionRecordId, preparedFiles) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(submissionRecordId);
  if (block) return block;
  return GeminiService.runAiAssessmentWithPreparedFiles(submissionRecordId, preparedFiles);
}

/* --- Batch Class AI Endpoints --- */
function apiGetClassSubmissions() {
  WebAccess.requireTeacher();
  var subs = Sheets.getSheetDataAsObjects(Config.SHEET_SUBMISSIONS);
  var result = [];
  for (var i = 0; i < subs.length; i++) {
    if (/jewell/i.test(String(subs[i].Task || '')) &&
        subs[i].Status !== Config.STATUS.LOCKED && subs[i].Status !== Config.STATUS.APPROVED) {
      result.push(subs[i]);
    }
  }
  return result;
}
function apiGradeSingleStudentBatch(submissionRecordId) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(submissionRecordId);
  if (block) return block; return GeminiService.runInitialAiAssessment(submissionRecordId); }

/* --- PDF Export Endpoints --- */
function apiGenerateAssessmentPdf(submissionRecordId) {
  WebAccess.requireTeacher();
  var block = LegacySafety.reject(submissionRecordId);
  if (block) return block;
  try {
    var detail = AssessmentService.getSubmissionForMarking(submissionRecordId);
    if (detail.submission.Status !== Config.STATUS.APPROVED && detail.submission.Status !== Config.STATUS.LOCKED) {
      return { success: false, message: 'Approve the assessment before exporting its final PDF.' };
    }
    var htmlContent = HtmlTemplates.getAssessmentPdfReportHtml(detail);
    var blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
    var cleanStudentName = detail.submission.StudentName.replace(/[^a-zA-Z0-9-]/g, '_');
    var fileName = cleanStudentName + '_Year9JewelleryAssessmentReport.pdf';
    blob.setName(fileName);
    return {
      success: true,
      fileName: fileName,
      dataUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    return { success: false, message: 'Report PDF generation failed: ' + Utils.sanitizeError(err) };
  }
}
function apiExportSubmittedDriveFile(driveFileId) {
  WebAccess.requireTeacher(); return DriveService.exportDriveFileAsPdf(driveFileId); }

function runEnsureRubricsLibraryStructure() {
  WebAccess.requireTeacher();
  var result = RubricsLibrary.ensureRubricsLibraryStructure();
  Logger.log(result);
}
