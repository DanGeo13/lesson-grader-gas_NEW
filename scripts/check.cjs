const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');

/* Scripts inside HTML must be plain JS: GAS template tags (<? ... ?>) make a
   legacy dialog file a template, not a parseable script. WebApp uses none. */
function checkHtmlScripts(relPath, content) {
  const scripts = content.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi);
  for (const script of scripts) {
    if (!script[1].includes('<?')) new vm.Script(script[1], { filename: relPath });
  }
}

/* The GitHub Pages launcher must stay deployment-agnostic: the Apps Script
   address is entered in its Settings and kept in the browser, never in code. */
function checkNoHardcodedDeployment(relPath, content) {
  if (/macros\/s\/[A-Za-z0-9_-]{10,}/.test(content)) {
    throw new Error(`${relPath} must not hardcode an Apps Script deployment id; it is configured in the page's Settings.`);
  }
}

const files = fs.readdirSync(root);
let doGet = 0;
for (const file of files) {
  if (file.endsWith('.gs')) {
    const content = fs.readFileSync(path.join(root, file), 'utf8');
    new vm.Script(content, { filename: file });
    /* Apps Script does not guarantee file load order, so every file must load
       on its own without touching another module or a Google service. */
    try { vm.runInContext(content, vm.createContext({}), { filename: file }); }
    catch (err) { throw new Error(file + ' must load without other files or services (load-order independence): ' + err.message); }
    doGet += (content.match(/^function doGet\s*\(/gm) || []).length;
    if (file === 'Code.gs') {
      for (const name of ['Config', 'Sheets', 'AssessmentService', 'ClassroomService', 'GeminiService']) {
        if ((content.match(new RegExp(`^var ${name}\\s*=`, 'gm')) || []).length !== 1) {
          throw new Error(`${name} must be declared exactly once; the duplicated half of Code.gs may have returned.`);
        }
      }
    }
    console.log('✓ ' + file + ' parses');
  } else if (file.endsWith('.html')) {
    const content = fs.readFileSync(path.join(root, file), 'utf8');
    checkHtmlScripts(file, content);
    console.log('✓ ' + file + ' scripts parse');
  }
}
if (doGet !== 1) throw new Error(`Expected one doGet, found ${doGet}`);
console.log('✓ one doGet');

/* Static launcher shipped to GitHub Pages (site/index.html). */
const siteDir = path.join(root, 'site');
if (fs.existsSync(siteDir)) {
  for (const file of fs.readdirSync(siteDir)) {
    if (!file.endsWith('.html')) continue;
    const relPath = 'site/' + file;
    const content = fs.readFileSync(path.join(siteDir, file), 'utf8');
    checkHtmlScripts(relPath, content);
    checkNoHardcodedDeployment(relPath, content);
    console.log('✓ ' + relPath + ' scripts parse with no hardcoded deployment');
  }
}

const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
if (manifest.runtimeVersion !== 'V8' || !(manifest.dependencies?.enabledAdvancedServices || []).some(service => service.userSymbol === 'Classroom')) {
  throw new Error('Apps Script manifest must declare V8 and the Classroom advanced service.');
}
const scopes = manifest.oauthScopes || [];
for (const scope of ['spreadsheets', 'drive', 'documents', 'script.external_request', 'script.scriptapp', 'classroom.coursework.students']) {
  if (!scopes.some(s => s.endsWith('/auth/' + scope))) throw new Error('Manifest must declare the ' + scope + ' OAuth scope explicitly.');
}
if (!manifest.webapp || manifest.webapp.executeAs !== 'USER_ACCESSING' || manifest.webapp.access === 'ANYONE_ANONYMOUS') {
  throw new Error('Web app must execute as the accessing user and never allow anonymous access.');
}
console.log('✓ manifest parses with Classroom service, explicit scopes and a safe web app model');
