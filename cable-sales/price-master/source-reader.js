import {parseXls} from '../parser.js';

const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const norm=s=>String(s??'').normalize('NFKC').replace(/[\s・/／ーｰ|｜\-‐]/g,'').toLowerCase();
const column=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
// These indices are the product columns in the existing sales report.
export const targets=[
 ...['既存','シェアド','MS'].flatMap((group,g)=>['SMTV','デジ録W','STB'].map((name,i)=>({index:10+g*3+i,category:'TV',group,name}))),
 ...['既存','シェアド'].flatMap((group,g)=>['10ギガ','3ギガ','1ギガ','5M'].map((name,i)=>({index:25+g*4+i,category:'NET',group,name}))),
 {index:33,category:'NET',group:'おくだけNET',name:'おくだけNET'},
 {index:34,category:'NET',group:'MS既存',name:'マンション'},
 ...[[43,'ケーブルプラス電話'],[44,'ケーブルライン'],[46,'ケーブルスマホ'],[47,'NHK'],[48,'ドリームコース'],[49,'J SAFE'],[50,'オプションチャンネル'],[51,'あんしんサポート'],[52,'IPスティック'],[53,'見るプラス見放題パック'],[54,'Netflix./Hulu/DAZN'],[55,'どこでもWifi'],[56,'STB追加'],[57,'ホームタウンでんき'],[58,'防犯カメラ'],[59,'BS再送信'],[60,'同軸再送信→光再送信'],[61,'同軸BSコース→光再送信'],[62,'同軸多ch→光多ch'],[63,'同軸ネット→光ネット']].map(([index,name])=>({index,category:'OP',group:'OP',name}))
].map(p=>({...p,col:column(p.index)}));
function matchTarget(category,group,name){
 let key=norm(name).replace(/^bd/, '').replace('ipステック','ipスティック');
 return targets.find(p=>norm(p.category)===norm(category)&&norm(p.group)===norm(group)&&norm(p.name)===key);
}
export function prepareRows(sheets){
 const rows=[];
 for(const sheet of sheets){
  let category='',group='';
  for(const {number,values:v} of sheet.rows){
   if(!v.some(x=>String(x??'').trim()))continue;
   const a=String(v[0]??'').trim(),b=String(v[1]??'').trim(),c=String(v[2]??'').trim();
   if(['単価','価格','金額'].includes(String(v[3]??'').trim()))continue;
   if(a)category=a;if(b)group=b;
   const cat=norm(group)==='op'?'OP':category;
   const name=c||(norm(group)==='おくだけnet'?group:'');
   if(!name)throw Error(`${sheet.name} ${number}行目：商品名を確認してください。A〜D列が「種別・区分・商品名・単価」の表に対応しています。`);
   let price=null;
   if(v[3]!==undefined&&v[3]!==null&&String(v[3]).trim()!==''){
    const raw=String(v[3]).normalize('NFKC').replace(/[,\s¥￥円]/g,'');
    if(/^(?:\d+)(?:\.0+)?$/.test(raw)&&Number.isSafeInteger(Number(raw)))price=Number(raw);
   }
   const target=matchTarget(cat,group,name);
   rows.push({category:cat,group,name,price,col:target?.col||'',source:`${sheet.name}!${number}`,issue:price===null?'単価が空欄または不正です。':''});
  }
 }
 if(!rows.length)throw Error('単価の行が見つかりません。');
 for(const row of rows){
  if(!row.col)row.issue+=' 対応する元列を選んでください。';
  else if(rows.filter(p=>p.col===row.col).length>1)row.issue+=' 同じ商品に複数の単価があります。使用する行を選んでください。';
  row.selected=!row.issue;
 }
 return rows;
}
function xml(text){const d=new DOMParser().parseFromString(text,'application/xml');if(d.getElementsByTagName('parsererror').length)throw Error('Excelの内部データを読み取れません。');return d;}
function stringValue(node){
 if(!node)return '';
 // Phonetic guides (rPh) are not part of the product name.
 return [...node.children].map(x=>x.localName==='t'?x.textContent:x.localName==='r'?[...x.children].filter(t=>t.localName==='t').map(t=>t.textContent).join(''):'').join('');
}
function resolvePath(target){
 const parts=(target.startsWith('/')?target.slice(1):'xl/'+target).split('/'),out=[];
 for(const part of parts){if(part==='..')out.pop();else if(part&&part!=='.')out.push(part);}
 return out.join('/');
}
export async function readSource(file){
 if(file.size>20*1024*1024)throw Error('20MB以下の単価ファイルを選んでください。');
 if(/\.json$/i.test(file.name))return {json:JSON.parse(await file.text())};
 if(/\.xls$/i.test(file.name)){
  const sheet=parseXls(await file.arrayBuffer());
  return {rows:prepareRows([{name:sheet.name,rows:[...sheet.cells].map(([i,values])=>({number:i+1,values}))}])};
 }
 if(!/\.xlsx$/i.test(file.name))throw Error('Excel（.xlsx／.xls）または単価元データのJSONを選んでください。PDFはExcelに変換してから読み込んでください。');
 if(!globalThis.JSZip)throw Error('Excel読み込み機能を準備できませんでした。ページを再読み込みしてください。');
 const zip=await globalThis.JSZip.loadAsync(await file.arrayBuffer());
 async function document(path){const entry=zip.file(path);if(!entry)throw Error('Excelのブック構造を確認できません。パスワードなしの.xlsx形式で保存してください。');return xml(await entry.async('string'));}
 const book=await document('xl/workbook.xml'),rels=await document('xl/_rels/workbook.xml.rels');
 const shared=zip.file('xl/sharedStrings.xml')?await document('xl/sharedStrings.xml'):null;
 const strings=shared?[...shared.getElementsByTagNameNS(NS,'si')].map(stringValue):[];
 const sheets=[];
 for(const sheet of book.getElementsByTagNameNS(NS,'sheet')){
  const id=sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
  const rel=[...rels.getElementsByTagName('Relationship')].find(r=>r.getAttribute('Id')===id);
  if(!rel||rel.getAttribute('TargetMode')==='External')continue;
  const doc=await document(resolvePath(rel.getAttribute('Target'))),rows=[];
  for(const row of doc.getElementsByTagNameNS(NS,'row')){
   const values=[];
   for(const c of row.getElementsByTagNameNS(NS,'c')){
    const ref=c.getAttribute('r'),col=ref?.match(/^([A-D])\d+$/)?.[1];if(!col)continue;
    const type=c.getAttribute('t'),value=c.getElementsByTagNameNS(NS,'v')[0]?.textContent;
    values[col.charCodeAt(0)-65]=type==='s'?strings[Number(value)]:type==='inlineStr'?stringValue(c.getElementsByTagNameNS(NS,'is')[0]):type==='e'?'Excelエラー':value;
   }
   rows.push({number:Number(row.getAttribute('r')),values});
  }
  sheets.push({name:sheet.getAttribute('name'),rows});
 }
 return {rows:prepareRows(sheets)};
}
