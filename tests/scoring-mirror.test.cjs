/* The browser shows a live score using a copy of the server's scoring code.
   This test extracts that copy from WebAppClient.html and compares it with
   RubricEngine on many random rubrics and selections (F11). */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { context } = require('./helpers.cjs');

const client = fs.readFileSync(path.join(__dirname, '..', 'WebAppClient.html'), 'utf8');
const block = client.split('/* BEGIN SCORING MIRROR')[1].split('/* END SCORING MIRROR */')[0];
const sandbox = vm.createContext({});
vm.runInContext('var Scoring; ' + block.slice(block.indexOf('*/') + 2) + '; this.Scoring = Scoring;', sandbox);
const Mirror = sandbox.Scoring;
const R = context().load('RubricEngine.gs').RubricEngine;

function rng(seed) { let s = seed; return () => (s = (s * 1103515245 + 12345) % 2147483648) / 2147483648; }

test('client live score equals the server engine on 500 random rubrics', () => {
  const rand = rng(42);
  const weightSets = [{ A: 1, B: 0.875, C: 0.7, D: 0.575, E: 0.25 }, { A: 1, B: 0.75, C: 0.5, D: 0.25, E: 0 }];
  const bandSets = [[{ letter: 'A', min: 85, max: 100 }, { letter: 'B', min: 75, max: 84.99 }, { letter: 'C', min: 65, max: 74.99 }, { letter: 'D', min: 50, max: 64.99 }, { letter: 'E', min: 0, max: 49.99 }],
    [{ letter: 'A', min: 80, max: 100 }, { letter: 'B', min: 70, max: 79.99 }, { letter: 'C', min: 50, max: 69.99 }, { letter: 'D', min: 30, max: 49.99 }, { letter: 'E', min: 0, max: 29.99 }]];
  for (let n = 0; n < 500; n++) {
    const rows = [];
    const criteria = 1 + Math.floor(rand() * 4);
    for (let c = 0; c < criteria; c++) {
      const marks = [3, 5, 7, 10, 12.5, 15, 20][Math.floor(rand() * 7)];
      ['A', 'B', 'C', 'D', 'E'].forEach(b => { const k = Math.floor(rand() * 4); for (let o = 0; o < k; o++) rows.push({ criterion: 'K' + c, maxMarks: marks, band: b, description: b + o }); });
      if (!rows.some(r => r.criterion === 'K' + c)) rows.push({ criterion: 'K' + c, maxMarks: marks, band: 'A', description: 'x' });
    }
    const profile = R.buildProfileFromRubricRows(rows, { taskName: 'T', name: 'T', weights: weightSets[n % 2], bands: bandSets[n % 2] });
    const states = {};
    profile.criteria.forEach(c => {
      const checks = {};
      c.bandOrder.forEach(b => c.bands[b].forEach(o => { if (rand() < 0.4) checks[o.id] = true; }));
      if (rand() < 0.05) checks[c.mdOverride.id] = true;
      states[c.criterionId] = { checkboxes: checks };
    });
    const server = R.scoreProfile(profile, states);
    const browser = Mirror.scoreProfile(JSON.parse(JSON.stringify(profile)), JSON.parse(JSON.stringify(states)));
    assert.ok(Math.abs(server.totalPoints - browser.totalPoints) < 1e-9, 'points case ' + n);
    assert.equal(browser.percent, server.percent, 'percent case ' + n);
    assert.equal(browser.letter, server.letter, 'letter case ' + n);
    profile.criteria.forEach(c => assert.equal(browser.criteria[c.criterionId].grade, server.criteria[c.criterionId].grade));
  }
});
