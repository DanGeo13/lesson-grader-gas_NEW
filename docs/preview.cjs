/* Local, mock-only design preview. Never connects to Google or stores keys.
   Run npm run preview; this server is NOT the Apps Script deployment.
     /            the Apps Script web app UI with mock data
     /launcher    the static GitHub Pages launcher (site/index.html) */
const http=require('node:http');
const fs=require('node:fs');
const path=require('node:path');
const root=path.resolve(__dirname,'..');
const {renderWebApp}=require('../scripts/render-webapp.cjs');
const html=renderWebApp({beforeClient:'<script src="/preview-mock.cjs"></script>'});
const launcher=fs.readFileSync(path.join(root,'site','index.html'),'utf8');
const port=Number(process.env.PORT)||4173;
http.createServer((req,res)=>{
  const pathname=req.url.split('?')[0];
  if(pathname==='/' || pathname==='/index.html'){
    res.writeHead(200,{'content-type':'text/html; charset=utf-8','cache-control':'no-store'});res.end(html);return;
  }
  if(pathname==='/launcher' || pathname==='/launcher/' || pathname==='/launcher/index.html'){
    res.writeHead(200,{'content-type':'text/html; charset=utf-8','cache-control':'no-store'});res.end(launcher);return;
  }
  if(pathname==='/preview-mock.cjs'){
    res.writeHead(200,{'content-type':'application/javascript; charset=utf-8','cache-control':'no-store'});
    res.end(fs.readFileSync(path.join(__dirname,'preview-mock.cjs')));return;
  }
  res.writeHead(404,{'content-type':'text/plain; charset=utf-8'});res.end('Not found.');
}).listen(port,'0.0.0.0',()=>console.log('Lesson Grader design preview (mock data only) listening on port '+port+'\n  /          mock web app UI\n  /launcher  static GitHub Pages launcher'));
