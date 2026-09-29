import {parseXls} from '../parser.js';
const $=s=>document.querySelector(s), yen=n=>new Intl.NumberFormat('ja-JP',{style:'currency',currency:'JPY',maximumFractionDigits:0}).format(n);
const esc=s=>String(s??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const xml=s=>new DOMParser().parseFromString(s,'application/xml');
const serialize=x=>new XMLSerializer().serializeToString(x);
const files={sales:null,confirmed:null,annual:null};
const month=$('#month');for(const n of [6,7,8,10,11,12,1,2,3,4,5])month.add(new Option(`${n}月`,n));
function monthCol(n,base){return String.fromCharCode(base+[6,7,8,9,10,11,12,1,2,3,4,5].indexOf(Number(n)));}
function value(map,ref){return Number(map[ref])||0;}
function repName(name){const raw=String(name||'').replace(/\s+/g,'');return ['浅野','小林','西沢','遠藤'].find(n=>raw===n||raw.startsWith(n)&&/^[ァ-ヶーぁ-ゖ]+$/.test(raw.slice(n.length)))||raw;}
async function loadBook(file){
 const zip=await JSZip.loadAsync(await file.arrayBuffer());
 const book=xml(await zip.file('xl/workbook.xml').async('string'));
 const rel=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
 const relMap=Object.fromEntries([...rel.getElementsByTagName('Relationship')].map(x=>[x.getAttribute('Id'),x.getAttribute('Target')]));
 const shared=zip.file('xl/sharedStrings.xml')?xml(await zip.file('xl/sharedStrings.xml').async('string')):null;
 const strings=shared?[...shared.getElementsByTagNameNS(ns,'si')].map(x=>[...x.getElementsByTagNameNS(ns,'t')].map(t=>t.textContent).join('')):[];
 const sheets={};for(const s of book.getElementsByTagNameNS(ns,'sheet')){
  const target=relMap[s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')];
  if(!target)continue;
  const path=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');
  const entry=zip.file(path);if(!entry)continue;
  const doc=xml(await entry.async('string')), map={};
  for(const c of doc.getElementsByTagNameNS(ns,'c')){
   const v=c.getElementsByTagNameNS(ns,'v')[0], inline=c.getElementsByTagNameNS(ns,'is')[0];
   map[c.getAttribute('r')]=c.getAttribute('t')==='s'?strings[Number(v?.textContent)]:inline?inline.textContent:v?.textContent??'';
  }
  sheets[s.getAttribute('name')]={path,doc,map};
 }
 return {zip,book,sheets};
}
function salesData(s){
 const head=s.cells.get(2)||[];
 if(String(head[1]||'').trim()!=='契約日'||String(head[4]||'').trim()!=='担当'||String(head[5]||'').trim()!=='名前')throw Error('営業実績表の列配置が異なります。');
 const rows=[...s.cells.entries()].filter(([i,r])=>i>=5&&Number.isFinite(r[1])&&r[4]&&r[5]);
 const reps={};for(const [,r] of rows){const k=repName(r[4]),x=reps[k]??={fresh:0,cancel:0,op:0};x.fresh+=Number(r[64])||0;for(let col=46;col<=62;col++)x.op+=Number(r[col])||0;}
 for(const [i,r] of s.cancellations)if(i>=5&&Number.isFinite(r[1])&&r[2]&&r[3]){const k=repName(r[2]),x=reps[k]??={fresh:0,cancel:0};x.cancel+=Number(r[61])||0;}
 const serial=rows[0]?.[1]?.[1];if(!serial)throw Error('営業の契約行が見つかりません。');
 const date=new Date(Date.UTC(1899,11,30)+Math.round(serial)*86400000);
 return {reps,rows:rows.length,sourceMonth:date.getUTCMonth()+1};
}
function confirmedData(book){const s=book.sheets['実績']?.map;if(!s)throw Error('確定ファイルに「実績」シートがありません。');
 const g=(ref)=>value(s,ref);
 const sales={shared:g('D30'),existing:g('D31'),tenant:g('D32'),op:g('D33'),fixed:g('D34'),cancel:g('D35'),total:g('D13')};
 const m={tochigi:[g('E48'),g('E49'),g('E50'),g('E51')],koga:[g('F48'),g('F49'),g('F50'),g('F51')],total:g('D14')};
 const counts={result:g('G26'),gain:g('G28'),cancel:g('G35'),shared:g('G30'),existing:g('G31'),tenant:g('G32'),op:g('G33'),tochigi:[g('H48'),g('H49'),g('H50'),g('H51')],koga:[g('I48'),g('I49'),g('I50'),g('I51')]};
 if(!sales.total||!m.total||Math.abs(sales.shared+sales.existing+sales.tenant+sales.op+sales.fixed-sales.cancel-sales.total)>1||Math.abs([...m.tochigi,...m.koga].reduce((a,b)=>a+b,0)-m.total)>1)throw Error('確定ファイルの営業・メンテ内訳が合計と一致しません。');
 const cancelSheet=book.sheets['栃木キャン']?.map,pk=book.sheets['栃木PK']?.map;
 if(!cancelSheet||!pk||value(cancelSheet,'H66')!==sales.cancel)throw Error('確定ファイルの栃木キャンと実績の金額が一致しません。');
 const repCancel={};for(let r=61;r<=64;r++)if(cancelSheet[`B${r}`])repCancel[repName(cancelSheet[`B${r}`])]={fresh:value(cancelSheet,`G${r}`),sales:value(cancelSheet,`H${r}`)};
 if(Object.values(repCancel).reduce((a,x)=>a+x.sales,0)!==sales.cancel)throw Error('キャンセル担当別の合計が一致しません。');
 const repSales={};for(let r=180;r<=183;r++)if(pk[`C${r}`])repSales[repName(pk[`C${r}`])]=value(pk,`J${r}`);
 return {sales,m,counts,overall:g('D9'),repCancel,repSales};
}
function setCell(sheet,ref,n){
 const doc=sheet.doc,data=doc.getElementsByTagNameNS(ns,'sheetData')[0];if(!data)throw Error('Excelのシート構造を読み取れません。');
 const rowNum=Number(ref.match(/\d+/)[0]);let row=[...data.getElementsByTagNameNS(ns,'row')].find(x=>Number(x.getAttribute('r'))===rowNum);
 if(!row){row=doc.createElementNS(ns,'row');row.setAttribute('r',rowNum);const next=[...data.children].find(x=>x.localName==='row'&&Number(x.getAttribute('r'))>rowNum);data.insertBefore(row,next||null);}
 let cell=[...row.children].find(x=>x.localName==='c'&&x.getAttribute('r')===ref);
 if(!cell){cell=doc.createElementNS(ns,'c');cell.setAttribute('r',ref);const next=[...row.children].find(x=>x.localName==='c'&&columnIndex(x.getAttribute('r'))>columnIndex(ref));row.insertBefore(cell,next||null);}
 cell.removeAttribute('t');for(const child of [...cell.children])if(child.localName==='f'||child.localName==='v'||child.localName==='is')cell.removeChild(child);
 const v=doc.createElementNS(ns,'v');v.textContent=String(n);cell.appendChild(v);
}
function setFormulaCache(sheet,ref,n){const cell=[...sheet.doc.getElementsByTagNameNS(ns,'c')].find(x=>x.getAttribute('r')===ref);if(!cell||![...cell.children].some(x=>x.localName==='f'))return;let v=[...cell.children].find(x=>x.localName==='v');if(!v){v=sheet.doc.createElementNS(ns,'v');cell.appendChild(v);}v.textContent=String(n);}
function columnIndex(ref){return [...ref.match(/^[A-Z]+/)[0]].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0);}
const fiscalMonths=[6,7,8,9,10,11,12,1,2,3,4,5];
function showView(){
 const view=location.hash==='#sales'?'sales-view':location.hash==='#maintenance'?'maintenance':'dashboard';
 for(const id of ['dashboard','sales-view','maintenance'])$(`#${id}`).hidden=id!==view;
 $('#import-panel').hidden=view!=='dashboard';$('#empty-view').hidden=!$('#result').hidden||view==='dashboard';
 if(!$('#result').hidden)window.scrollTo({top:0,behavior:'auto'});
}
function progressBar(actual,target){const ratio=target>0?actual/target:0,pct=target>0?Math.round(ratio*100):0;return `<div class="track" role="img" aria-label="進捗 ${pct}%"><span class="${ratio>=1?'over':''}" style="width:${Math.max(0,Math.min(100,pct))}%"></span></div>`;}
function renderProgress(annual,d,n){
 const sheet=annual.sheets['部署売上']?.map;if(!sheet)throw Error('2026年ファイルに「部署売上」シートがありません。');
 const selected=fiscalMonths.indexOf(n);
 const data=fiscalMonths.map((m,i)=>{const col=monthCol(m,69),target=value(sheet,`${col}10`),saved=value(sheet,`${col}20`);return {month:m,target,actual:i>selected?0:i===selected?d.overall:saved,future:i>selected};});
 $('#monthly-progress').innerHTML=data.map(x=>`<div class="progress-row"><span>${x.month}月</span><div>${progressBar(x.actual,x.target)}</div><strong>${x.future?'未到来':yen(x.actual)} / ${yen(x.target)}</strong></div>`).join('');
 const sum=xs=>xs.reduce((a,x)=>({target:a.target+x.target,actual:a.actual+x.actual}),{target:0,actual:0});
 const periods=[['上期（6〜11月）',sum(data.slice(0,6))],['下期（12〜5月）',sum(data.slice(6))],['年間（6〜5月）',sum(data)]];
 $('#period-progress').innerHTML=periods.map(([label,x])=>`<div class="progress-card"><strong>${label}</strong><div class="line"><span>実績 ${yen(x.actual)}</span><strong>${x.target?Math.round(x.actual/x.target*100):0}%</strong></div>${progressBar(x.actual,x.target)}<small>計画 ${yen(x.target)}</small></div>`).join('');
}
async function updateAnnual(book,d,s,n){
 const annual=book.sheets['部署売上'],sales=book.sheets['営業'];if(!annual||!sales||!book.sheets['メンテ'])throw Error('2026年ファイルに「部署売上」「営業」「メンテ」が揃っていません。');
 const c=monthCol(n,69),sc=monthCol(n,68),put=(r,v)=>setCell(annual,`${c}${r}`,v);
 for(const [r,k] of [[40,'shared'],[41,'existing'],[42,'tenant'],[43,'op'],[44,'fixed'],[45,'cancel']])put(r,d.sales[k]);
 [d.m.tochigi,d.m.koga].forEach((list,i)=>list.forEach((v,j)=>put((i?59:53)+j,v)));
 for(const [r,v] of [[75,d.counts.result],[76,d.counts.gain],[77,d.counts.cancel],[78,d.counts.shared+d.counts.existing+d.counts.tenant],[79,d.counts.shared],[80,d.counts.existing],[81,d.counts.tenant],[82,d.counts.op],...[85,86,87,88].map((r,i)=>[r,d.counts.tochigi[i]]),...[90,91,92,93].map((r,i)=>[r,d.counts.koga[i]])])put(r,v);
 for(const [r,v] of [[14,d.overall],[15,d.sales.total],[16,d.m.total],[20,d.overall],[37,d.sales.total],[47,d.m.total]])setFormulaCache(annual,`${c}${r}`,v);
 const reps=['浅野','小林','西沢','遠藤'];reps.forEach((name,i)=>{const x=s.reps[name]||{fresh:0,cancel:0};setCell(sales,`${sc}${13+i}`,x.fresh);setCell(sales,`${sc}${21+i}`,d.repCancel[name]?.fresh||0);});
 // Formula cells and unrelated months stay intact. Excel refreshes dependent cached values on open.
 let calc=book.book.getElementsByTagNameNS(ns,'calcPr')[0];if(!calc){calc=book.book.createElementNS(ns,'calcPr');book.book.documentElement.appendChild(calc);}calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');calc.setAttribute('calcMode','auto');
 book.zip.file('xl/workbook.xml',serialize(book.book));
 for(const sh of [annual,sales])book.zip.file(sh.path,serialize(sh.doc));
 return book.zip.generateAsync({type:'blob',compression:'DEFLATE'});
}
async function render(){const status=$('#status');$('#save').disabled=true;$('#result').hidden=true;try{
 if(!files.sales||!files.confirmed||!files.annual){status.textContent='3つのファイルを選んでください。';showView();return;}
 status.textContent='読み込み中…';const n=Number(month.value),sales=salesData(parseXls(await files.sales.arrayBuffer()));
 const [confirmed,annual]=await Promise.all([loadBook(files.confirmed),loadBook(files.annual)]);
 const d=confirmedData(confirmed);
 const warnings=[];if(sales.sourceMonth!==n)warnings.push(`営業実績表は${sales.sourceMonth}月の契約日です。反映先は${n}月です。`);
 const unallocated=d.sales.shared+d.sales.existing+d.sales.tenant+d.sales.op-Object.values(d.repSales).reduce((a,b)=>a+b,0);
 if(unallocated)warnings.push(`担当者別元売上と営業区分売上の差額 ${yen(unallocated)} は担当者に配賦していません。固定費も担当者別には含めません。`);
 if(Math.abs(d.overall-d.sales.total-d.m.total)>1)warnings.push('確定ファイルの合計に営業・メンテ以外の売上が含まれます。');
 const existing=annual.sheets['部署売上'].map;if(existing[`${monthCol(n,69)}40`]||existing[`${monthCol(n,69)}53`])warnings.push('2026年ファイルのこの月には既存値があります。保存時は選択月の対象項目を置き換えます。');
 $('#warning').hidden=!warnings.length;$('#warning').textContent=warnings.join('　');
 $('#total').textContent=yen(d.overall);$('#sales-total').textContent=yen(d.sales.total);$('#maintenance-total').textContent=yen(d.m.total);
 $('#sales-kpi').textContent=yen(d.sales.total);$('#shared-kpi').textContent=`${d.counts.shared}件`;$('#existing-kpi').textContent=`${d.counts.existing}件`;
 $('#maintenance-kpi').textContent=yen(d.m.total);$('#tochigi-kpi').textContent=yen(d.m.tochigi.reduce((a,b)=>a+b,0));$('#koga-kpi').textContent=yen(d.m.koga.reduce((a,b)=>a+b,0));
 renderProgress(annual,d,n);
 $('#sales-rows').innerHTML=[['シェアド',d.counts.shared,d.sales.shared],['既存',d.counts.existing,d.sales.existing],['店子',d.counts.tenant,d.sales.tenant],['OP',d.counts.op,d.sales.op],['固定費',0,d.sales.fixed],['キャンセル控除',d.counts.cancel,-d.sales.cancel]].map(([name,count,amount])=>`<tr><th>${name}</th><td>${count}</td><td>${yen(amount)}</td></tr>`).join('');
 $('#maintenance-rows').innerHTML=[['栃木',d.m.tochigi,d.counts.tochigi],['古河',d.m.koga,d.counts.koga]].flatMap(([name,amounts,counts])=>['訪問','時間','延長','OP'].map((label,i)=>`<tr><th>${name} ${label}</th><td>${counts[i]}</td><td>${yen(amounts[i])}</td></tr>`)).join('');
 $('#reps').innerHTML=Object.entries(sales.reps).map(([name,x])=>`<tr><th>${esc(name)}</th><td>${d.repSales[name]===undefined?'要確認':yen(d.repSales[name]-(d.repCancel[name]?.sales||0))}</td><td>${x.fresh-(d.repCancel[name]?.fresh||0)}</td><td>${x.fresh}</td><td>${d.repCancel[name]?.fresh||0}</td><td>${x.op||0}</td></tr>`).join('');
 $('#result').hidden=false;status.textContent=`確定ファイルの営業 ${yen(d.sales.total)} とメンテ ${yen(d.m.total)} を確認しました。`;status.className='notice ok';showView();
 $('#save').disabled=false;$('#save').onclick=async()=>{try{status.textContent='Excelファイルを作成中…';$('#save').disabled=true;const blob=await updateAnnual(annual,d,sales,n),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=`2026年_${n}月反映.xlsx`;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);status.textContent='保存しました。Excelで開いて数式を再計算してください。';}catch(e){status.textContent=e.message;status.className='notice error';}finally{$('#save').disabled=false;}};
 }catch(e){console.error(e);status.textContent=e.message||'ファイルを読み込めませんでした。';status.className='notice error';showView();}}
for(const key of Object.keys(files))$(`#${key}`).addEventListener('change',e=>{files[key]=e.target.files[0]||null;$(`#${key}-name`).textContent=files[key]?.name||'未選択';render();});month.addEventListener('change',render);
window.addEventListener('hashchange',showView);showView();
