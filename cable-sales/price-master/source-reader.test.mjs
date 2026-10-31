import {test} from 'node:test';
import assert from 'node:assert/strict';
import {prepareRows,targets} from './source-reader.js';
const sheet=values=>[{name:'単価',rows:values.map((values,i)=>({number:i+1,values}))}];
test('product columns agree with the sales report, including OP gaps',()=>{
 assert.equal(targets.find(t=>t.name==='ケーブルプラス電話').col,'AR');
 assert.equal(targets.find(t=>t.name==='ケーブルスマホ').col,'AU');
 assert.equal(targets.find(t=>t.name==='IPスティック').col,'BA');
 assert.equal(targets.find(t=>t.name==='同軸ネット→光ネット').col,'BL');
 assert.equal(new Set(targets.map(t=>t.col)).size,targets.length);
});
test('normalize display variants without changing imported names',()=>{
 const rows=prepareRows(sheet([['TV','既存','SMＴＶ',34000],['TV','既存','BD/デジ録W',34000],['NET','既存','１０ギガ','３２,０００円'],['電話','OP','ケ-ブルプラス電話',8000],['','OP','IPステック',1000]]));
 assert.deepEqual(rows.map(p=>p.col),['K','L','Z','AR','BA']);
 assert.ok(rows.every(p=>p.selected));assert.equal(rows[2].price,32000);
 assert.equal(rows[0].name,'SMＴＶ');
});
test('blank and invalid prices stay unset; explicit zero is valid',()=>{
 const rows=prepareRows(sheet([['NET','おくだけNET','',''],['','OP','IPステック',0],['','OP','NHK',-1],['','OP','STB追加','abc']]));
 assert.deepEqual(rows.map(p=>p.price),[null,0,null,null]);
 assert.deepEqual(rows.map(p=>p.selected),[false,true,false,false]);
 assert.equal(rows[0].col,'AH');
});
test('duplicate target prices require explicit selection',()=>{
 const rows=prepareRows(sheet([['NET','シェアド','５M',3000],['NET','シェアド','５M',35000]]));
 assert.ok(rows.every(p=>!p.selected&&p.col==='AG'&&p.issue.includes('複数')));
});
test('unknown products cannot be silently assigned to report columns',()=>{
 const [row]=prepareRows(sheet([['NET','既存','新商品',5000]]));
 assert.equal(row.col,'');assert.equal(row.selected,false);
});
test('headers are optional and blank sheets are ignored',()=>{
 const rows=prepareRows(sheet([['種別','区分','商品名','単価'],[],['TV','MS','STB',25800]]));
 assert.equal(rows.length,1);assert.equal(rows[0].col,'S');assert.equal(rows[0].source,'単価!3');
 assert.throws(()=>prepareRows(sheet([])),/見つかりません/);
});
