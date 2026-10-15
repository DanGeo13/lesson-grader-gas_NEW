const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');

function context(mocks = {}) {
  const sandbox = vm.createContext({ console: { log() {}, error() {}, warn() {} }, ...mocks });
  return {
    sandbox,
    load(...files) { for (const name of files) vm.runInContext(fs.readFileSync(path.join(root,name),'utf8'),sandbox,{filename:name}); return sandbox; }
  };
}

class Sheet {
  constructor(name, rows = []) { this.name = name; this.rows = rows.map(row=>row.slice()); this.columns = Math.max(30, ...rows.map(row=>row.length)); }
  getName() { return this.name; }
  isSheetHidden() { return false; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(0,...this.rows.map(row=>row.length)); }
  getMaxColumns() { return this.columns; }
  getDataRange() { return this.getRange(1,1,this.getLastRow(),this.getLastColumn()); }
  getRange(r,c,nRows=1,nCols=1) {
    const sh=this;
    return {
      getValues() { const out=[]; for(let i=0;i<nRows;i++){const row=[]; for(let j=0;j<nCols;j++)row.push(sh.rows[r+i-1]?.[c+j-1] ?? '');out.push(row);}return out; },
      getValue() { return sh.rows[r-1]?.[c-1] ?? ''; },
      setValues(data) { for(let i=0;i<data.length;i++){sh.rows[r+i-1] ||= [];for(let j=0;j<data[i].length;j++)sh.rows[r+i-1][c+j-1]=data[i][j];}return this; },
      setValue(v) { sh.rows[r-1] ||= [];sh.rows[r-1][c-1]=v;return this; },
      setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; }
    };
  }
  setFrozenRows() {}
  insertColumnsAfter(_, n) { this.columns += n; }
  appendRow(row) { this.rows.push(row.slice()); }
  clearContents() { this.rows = []; }
}

class Workbook {
  constructor(name='7TECHI [2026] - Rube Goldberg - Grading') { this.name=name; this.id='WORKBOOK'+'0'.repeat(30); this.sheets=new Map(); }
  getId() { return this.id; }
  getName() { return this.name; }
  getSheetByName(name) { return this.sheets.get(name) || null; }
  getSheets() { return [...this.sheets.values()]; }
  insertSheet(name) { const sh=new Sheet(name);this.sheets.set(name,sh);return sh; }
  add(name,rows) { const sh=new Sheet(name,rows);this.sheets.set(name,sh);return sh; }
}

function objectRows(sheet) {
  const [h,...rest]=sheet.rows;
  return rest.map(row=>Object.fromEntries(h.map((key,i)=>[key,row[i]??''])));
}
module.exports={context,Sheet,Workbook,objectRows};
