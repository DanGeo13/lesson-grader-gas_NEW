/* Resolves the Apps Script include() scriptlets in WebApp.html so the page can be
   loaded by jsdom tests and the local mock preview exactly as HtmlService would
   serve it. Only the closed list of partials used by doGet is supported. */
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const PARTIALS = ['WebAppStyles', 'WebAppClient'];

function renderWebApp(options = {}) {
  let html = fs.readFileSync(path.join(root, 'WebApp.html'), 'utf8');
  html = html.replace(/<\?!=\s*include\('([A-Za-z]+)'\);?\s*\?>/g, (match, name) => {
    if (!PARTIALS.includes(name)) throw new Error('Unknown partial ' + name);
    const content = fs.readFileSync(path.join(root, name + '.html'), 'utf8');
    return name === 'WebAppClient' && options.beforeClient ? options.beforeClient + '\n' + content : content;
  });
  if (/<\?/.test(html)) throw new Error('Unresolved scriptlet in WebApp.html');
  return html;
}
module.exports = { renderWebApp, PARTIALS };
