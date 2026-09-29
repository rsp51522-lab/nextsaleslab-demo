import {parseXls} from './parser.js';
import {makeWorkbook} from './xlsx.js';
const $=s=>document.querySelector(s),yen=n=>new Intl.NumberFormat('ja-JP',{style:'currency',currency:'JPY',maximumFractionDigits:0}).format(n);
const html=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;');
const date=n=>{if(!Number.isFinite(n))return '';return new Date(Date.UTC(1899,11,30)+Math.round(n)*86400000).toISOString().slice(0,10);};
const raw=await fetch('./prices.json').then(r=>r.json());
const products=raw.map(p=>({...p,price:Number(localStorage.getItem('cable-price-'+p.col)??p.price)}));
let parsed=null,sourceName='';
const status=$('#status'),fileInput=$('#file'),priceList=$('#price-list');
function error(msg){status.textContent=msg;status.className='notice error';$('#results').hidden=true;}
function value(cell){return cell??'';}
function calculate(){
 if(!parsed)return;
 const records=parsed.rows.map(r=>({...r,items:products.map(p=>Number(r.source[p.index])||0)}));
 const cancellations=parsed.cancellations.map(r=>({...r,items:products.map(p=>Number(r.source[p.index-3])||0)}));
 const byRep={},byCategory={シェアド:{count:0,sales:0},既存:{count:0,sales:0},店子:{count:0,sales:0},OP:{count:0,sales:0},キャン:{count:0,sales:0},'電話・その他':{count:0,sales:0}};
 const group=p=>[13,14,15,29,30,31,32].includes(p.index)?'シェアド':([10,11,12,25,26,27,28].includes(p.index)?'既存':([16,17,18,34].includes(p.index)?'店子':(p.index>=46?'OP':'電話・その他')));
 for(const p of products){p.count=records.reduce((a,r)=>a+(Number(r.source[p.index])||0),0);p.sales=p.count*p.price;const x=byCategory[group(p)];x.count+=p.count;x.sales+=p.sales;}
 for(const r of records){r.sales=r.items.reduce((a,n,i)=>a+n*products[i].price,0);const x=byRep[r.rep]??{count:0,gross:0,cancelSales:0,fresh:0,cancelFresh:0,op:0};x.count++;x.gross+=r.sales;x.fresh+=Number(r.source[64])||0;x.op+=products.reduce((a,p)=>a+(p.index>=46?(Number(r.source[p.index])||0):0),0);byRep[r.rep]=x;}
 for(const c of cancellations){const x=byRep[c.rep]??{count:0,gross:0,cancelSales:0,fresh:0,cancelFresh:0,op:0};x.cancelFresh+=c.fresh;byRep[c.rep]=x;byCategory.キャン.count+=c.fresh;}
 const gross=records.reduce((a,r)=>a+r.sales,0),period=parsed.period;
 // The supplied .xls contains cancellation quantities, but no cancellation payout rates.
 // This one verified September snapshot is from the user's separate cancellation summary image.
 const verifiedSeptember=period==='2026年9月'&&records.length===49&&cancellations.length===5&&
  ['浅野|照内勝宏|1','浅野|山中康弘|0','遠藤|木村恵理子|0','小林|内山幸之|1','小林|飯村訓雄|1'].every(k=>cancellations.some(c=>`${c.rep}|${c.customer}|${c.fresh}`===k));
 const cancelSales=verifiedSeptember?72000:(cancellations.length?null:0),total=cancelSales===null?null:gross-cancelSales;
 if(verifiedSeptember){for(const [name,value] of Object.entries({浅野:24000,遠藤:2000,西沢:0,小林:46000}))if(byRep[name])byRep[name].cancelSales=value;byCategory.キャン.sales=-72000;}
 else if(cancelSales===null)byCategory.キャン.sales=null;
 $('#total').textContent=total===null?'キャンセル金額の確認が必要':yen(total);$('#contract-count').textContent=records.length+'件';$('#period').textContent=period+'の集計';
 $('#categories').innerHTML=Object.entries(byCategory).map(([k,v])=>`<tr><th>${k}</th><td>${v.count}件</td><td>${v.sales===null?'要確認':yen(v.sales)}</td></tr>`).join('');
 $('#reps').innerHTML=Object.entries(byRep).map(([k,v])=>`<tr><th>${html(k)}</th><td>${cancelSales===null?'要確認':yen(v.gross-v.cancelSales)}</td><td>${v.fresh-v.cancelFresh}</td><td>${v.fresh}</td><td>${v.cancelFresh}</td><td>${v.op}</td></tr>`).join('');
 $('#cancel-breakdown').hidden=!verifiedSeptember;
 $('#items').innerHTML=products.filter(p=>p.count).map(p=>`<tr><th>${html(p.category)} · ${html(p.group)} · ${html(p.name)}</th><td>${p.count}</td><td>${yen(p.price)}</td><td>${yen(p.sales)}</td></tr>`).join('');
 const unknown=[];for(let j=10;j<=63;j++)if(!products.some(p=>p.index===j)){let count=records.reduce((a,r)=>a+(Number(r.source[j])||0),0);if(count)unknown.push(`${column(j)}列 ${count}件`);}
 const warnings=[...parsed.warnings];if(unknown.length)warnings.push('単価未設定の項目：'+unknown.join('、'));
 if(cancelSales===null)warnings.push('キャンセルシートには数量だけがあり、売上の戻入単価がありません。金額の確定には「栃木キャンセル」の計算表が必要です。');
 if(verifiedSeptember)warnings.push('９月のキャンセル売上は、確認済みの栃木キャンセル集計表（72,000円）を適用しています。');
 $('#warnings').textContent=warnings.join('　');$('#warnings').hidden=!warnings.length;
 $('#results').hidden=false;status.textContent=`${sourceName} を読み込みました。集計内容を確認してください。`;status.className='notice ok';
 $('#download').disabled=cancelSales===null;
 $('#download').onclick=()=>{if(cancelSales===null)return;const blob=makeWorkbook({period,products,records,cancellations,byRep,byCategory,total,gross,cancelSales,fileName:sourceName,verifiedSeptember});const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=`${period.replace(/[^0-9]/g,'')}_ケーブル売上管理表.xlsx`;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000);};
}
function column(j){let s='',n=j+1;while(n){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26);}return s;}
function showPrices(){priceList.innerHTML=products.map((p,i)=>`<label class="price-row"><span><small>${html(p.col)} · ${html(p.category)}</small>${html(p.group)} ${html(p.name)}</span><span class="price-input"><input type="number" min="0" step="1" inputmode="numeric" data-index="${i}" value="${p.price}" aria-label="${html(p.col)} ${html(p.name)}の単価"><span>円</span></span></label>`).join('');}
priceList.onchange=e=>{if(!e.target.matches('input'))return;const p=products[Number(e.target.dataset.index)],n=Number(e.target.value);if(!Number.isFinite(n)||n<0){e.target.value=p.price;return;}p.price=n;localStorage.setItem('cable-price-'+p.col,String(n));calculate();};
showPrices();
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
