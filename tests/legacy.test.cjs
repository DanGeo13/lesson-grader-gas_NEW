const {test}=require('node:test');
const assert=require('node:assert/strict');
const {context}=require('./helpers.cjs');

test('legacy markbook feedback helper exists, so approval no longer half-applies (F08)',()=>{
  const built=[];
  const g=context({
    SpreadsheetApp:{
      newTextStyle:()=>({setBold(){return this;},build:()=>({bold:true})}),
      newRichTextValue:()=>{const b={text:'',styles:[]};b.setText=t=>{b.text=t;return b;};b.setTextStyle=(s,e)=>{b.styles.push([s,e]);return b;};b.build=()=>{built.push(b);return b;};return b;}
    }
  }).load('Code.gs');
  const rich=g.Utils.buildRichTextFromMarkdown('Great **folio** work <script>');
  assert.equal(rich.text,'Great folio work <script>');
  assert.deepEqual(rich.styles,[[6,11]]);
  assert.equal(g.Utils.buildRichTextFromMarkdown(null).text,'');
});
