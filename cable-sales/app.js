import {parseXls} from './parser.js';
import {makeWorkbook} from './xlsx.js';
import {readConfirmation} from './confirmed.js';
const $=s=>document.querySelector(s),yen=n=>new Intl.NumberFormat('ja-JP',{style:'currency',currency:'JPY',maximumFractionDigits:0}).format(n);
const html=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;');
const date=n=>{if(!Number.isFinite(n))return '';return new Date(Date.UTC(1899,11,30)+Math.round(n)*86400000).toISOString().slice(0,10);};
const raw=await fetch('./prices.json').then(r=>r.json());
const products=raw.map(p=>({...p,price:Number(localStorage.getItem('cable-price-'+p.col)??p.price)}));
let parsed=null,confirmed=null,sourceName='';
const status=$('#status'),fileInput=$('#file');
function error(msg){status.textContent=msg;status.className='notice error';$('#results').hidden=true;}
function value(cell){return cell??'';}
function canonicalRep(name){
 const raw=String(name||'').replace(/\s+/g,'').trim();
 if(!confirmed)return raw;
 const names=Object.keys(confirmed.reps);
 return names.find(n=>raw===n)||names.find(n=>raw.startsWith(n)&&/^[ァ-ヶーぁ-ゖ]+$/.test(raw.slice(n.length)))||raw;
}
function calculate(){
 if(!parsed||!confirmed){$('#results').hidden=true;status.textContent='営業実績管理表と月次確定ファイルの両方を選んでください。';return;}
 const records=parsed.rows.map(r=>({...r,rep:canonicalRep(r.rep),items:products.map(p=>Number(r.source[p.index])||0)}));
 const cancellations=parsed.cancellations.map(r=>({...r,rep:canonicalRep(r.rep),items:products.map(p=>Number(r.source[p.index-3])||0)}));
 const byRep={},byCategory={シェアド:{count:0,sales:0},既存:{count:0,sales:0},店子:{count:0,sales:0},OP:{count:0,sales:0},キャン:{count:0,sales:0},'電話・その他':{count:0,sales:0}};
 const group=p=>[13,14,15,29,30,31,32].includes(p.index)?'シェアド':([10,11,12,25,26,27,28].includes(p.index)?'既存':([16,17,18,34].includes(p.index)?'店子':(p.index>=46?'OP':'電話・その他')));
 for(const p of products){p.count=records.reduce((a,r)=>a+(Number(r.source[p.index])||0),0);p.sales=p.count*p.price;const x=byCategory[group(p)];x.count+=p.count;x.sales+=p.sales;}
 const unmatched=new Set();
 for(const r of records){r.sales=r.items.reduce((a,n,i)=>a+n*products[i].price,0);if(!confirmed.reps[r.rep]){unmatched.add(r.rep);continue;}const x=byRep[r.rep]??{count:0,gross:0,cancelSales:0,fresh:0,cancelFresh:0,op:0};x.count++;x.fresh+=Number(r.source[64])||0;x.op+=products.reduce((a,p)=>a+(p.index>=46?(Number(r.source[p.index])||0):0),0);byRep[r.rep]=x;}
 const gross=confirmed.amount.sales+confirmed.amount.cancel,total=confirmed.amount.sales,cancelSales=confirmed.amount.cancel,period=parsed.period;
 for(const [name,actual] of Object.entries(confirmed.reps)){
  const x=byRep[name]??{count:0,gross:0,cancelSales:0,fresh:0,cancelFresh:0,op:0};
  x.gross=actual.gross;x.cancelSales=actual.cancelSales;x.cancelFresh=actual.cancelFresh;byRep[name]=x;
 }
 for(const [label,key] of [['シェアド','shared'],['既存','existing'],['店子','tenant'],['OP','op']]){
  byCategory[label].sales=confirmed.amount[key];byCategory[label].count=confirmed.count[key];
 }
 byCategory.キャン={count:confirmed.count.cancel,sales:-cancelSales};
 delete byCategory['電話・その他'];byCategory.固定費={count:0,sales:confirmed.amount.fixed};
 $('#total').textContent=yen(total);$('#contract-count').textContent=records.length+'件';$('#period').textContent=period+'の集計';
 $('#categories').innerHTML=Object.entries(byCategory).map(([k,v])=>`<tr><th>${k}</th><td>${v.count}件</td><td>${v.sales===null?'要確認':yen(v.sales)}</td></tr>`).join('');
 $('#reps').innerHTML=Object.entries(byRep).map(([k,v])=>`<tr><th>${html(k)}</th><td>${yen(v.gross-v.cancelSales)}</td><td>${v.fresh-v.cancelFresh}</td><td>${v.fresh}</td><td>${v.cancelFresh}</td><td>${v.op}</td></tr>`).join('');
 $('#cancel-rows').innerHTML=Object.entries(confirmed.reps).map(([k,v])=>`<tr><th>${html(k)}</th><td>${v.cancelFresh}</td><td>${yen(v.cancelSales)}</td></tr>`).join('')+`<tr><th>合計</th><td>${confirmed.count.cancel}</td><td>${yen(cancelSales)}</td></tr>`;
 $('#items').innerHTML=products.filter(p=>p.count).map(p=>`<tr><th>${html(p.category)} · ${html(p.group)} · ${html(p.name)}</th><td>${p.count}</td><td>${yen(p.price)}</td><td>${yen(p.sales)}</td></tr>`).join('');
 const unknown=[];for(let j=10;j<=63;j++)if(!products.some(p=>p.index===j)){let count=records.reduce((a,r)=>a+(Number(r.source[j])||0),0);if(count)unknown.push(`${column(j)}列 ${count}件`);}
 const warnings=[...parsed.warnings];if(unknown.length)warnings.push('単価未設定の項目：'+unknown.join('、'));
 if(confirmed.month&&confirmed.month!==Number(period.match(/年(\d+)月/)?.[1]))warnings.push(`営業実績表は${period}、確定ファイル名は${confirmed.month}月です。対象月を確認してください。`);
 if(confirmed.adjustment)warnings.push(`確定ファイルの担当者別売上計と営業区分計の差額 ${yen(confirmed.adjustment)} は担当者に配賦していません。固定費 ${yen(confirmed.amount.fixed)} は営業売上合計に含みます。`);
 if(cancellations.reduce((a,c)=>a+c.fresh,0)!==confirmed.count.cancel)warnings.push('営業実績表のキャンセル真水数と確定ファイルが異なるため、確定ファイルの数字を使用しています。');
 if(unmatched.size)warnings.push('確定ファイルに紐付かない担当者：'+[...unmatched].join('、')+'。元ファイルの担当名を確認してください。');
 $('#warnings').textContent=warnings.join('　');$('#warnings').hidden=!warnings.length;
 $('#results').hidden=false;status.textContent=`${sourceName} を読み込みました。集計内容を確認してください。`;status.className='notice ok';
 $('#download').disabled=unmatched.size>0;
 $('#download').onclick=()=>{const blob=makeWorkbook({period,products,records,cancellations,byRep,byCategory,total,gross,cancelSales,fileName:sourceName,confirmed});const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=`${period.replace(/[^0-9]/g,'')}_ケーブル売上管理表.xlsx`;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000);};
}
function column(j){let s='',n=j+1;while(n){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26);}return s;}
window.addEventListener('pageshow',()=>{for(const p of products)p.price=Number(localStorage.getItem('cable-price-'+p.col)??raw.find(x=>x.col===p.col).price);calculate();});
$('#confirmed').onchange=async e=>{const f=e.target.files[0];confirmed=null;$('#confirmed-name').textContent=f?.name||'未選択';if(!f){calculate();return;}try{confirmed=await readConfirmation(f);calculate();}catch(err){error(err.message);}};
async function openFile(file){if(!file)return;sourceName=file.name;
 if(!/\.xls$/i.test(file.name)){error('実績管理表の .xls ファイルを選んでください。');return;}
 status.textContent='読み込み中…';status.className='notice';
 try{
  const sheet=parseXls(await file.arrayBuffer());const h=sheet.cells.get(2)||[];
  if(String(h[1]||'').trim()!=='契約日'||String(h[4]||'').trim()!=='担当'||String(h[5]||'').trim()!=='名前'||String(h[10]||'').trim()!=='TV'||String(h[25]||'').trim()!=='NET')throw Error('実績管理表の列配置が想定と異なります。元ファイルを確認してください。');
  const rows=[...sheet.cells.entries()].filter(([i,r])=>i>=5&&Number.isFinite(r[1])&&String(r[5]||'').trim()&&String(r[4]||'').trim()).map(([i,r])=>({sourceRow:i+1,date:date(r[1]),workDate:date(r[2]),rep:String(r[4]).trim(),customer:String(r[5]).trim(),area:String(value(r[9])).trim(),source:r}));
  if(!rows.length)throw Error('契約行が見つかりません。');
  const months=[...new Set(rows.map(r=>r.date.slice(0,7)))];const warnings=[];if(months.length!==1)warnings.push('契約日が複数月にまたがっています。');
  const cancelRows=[...sheet.cancellations.entries()].filter(([i,r])=>i>=5&&Number.isFinite(r[1])&&String(r[2]||'').trim()&&String(r[3]||'').trim()).map(([i,r])=>({sourceRow:i+1,rep:String(r[2]).trim(),customer:String(r[3]).trim(),source:r,fresh:Number(r[61])||0,note:String(r[65]||'')}));
  if(!sheet.cancellations.size)warnings.push('キャンセルシートがありません。キャンセル売上・真水は0で表示しています。');
  const first=rows[0].date;parsed={rows,cancellations:cancelRows,period:`${Number(first.slice(0,4))}年${Number(first.slice(5,7))}月`,warnings};calculate();
 }catch(e){console.error(e);error(e.message||'ファイルを読み取れませんでした。');}
}
fileInput.onchange=e=>openFile(e.target.files[0]);
const drop=$('#drop');drop.ondragover=e=>{e.preventDefault();drop.classList.add('drag');};drop.ondragleave=()=>drop.classList.remove('drag');drop.ondrop=e=>{e.preventDefault();drop.classList.remove('drag');openFile(e.dataTransfer.files[0]);};
